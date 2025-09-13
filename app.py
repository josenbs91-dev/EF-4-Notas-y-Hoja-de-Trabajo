import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (incluye Hoja de Trabajo y HT EF-4)", type=["xlsx"])

if uploaded_file and equiv_file:
    # -------------------------------
    # 1. Procesar archivo principal
    # -------------------------------
    df = pd.read_excel(uploaded_file)

    # Mantener formato original
    df = df.copy()

    # Crear exp_contable
    df["exp_contable"] = (
        df["ano_eje"].astype(str) + "-" +
        df["nro_not_exp"].astype(str) + "-" +
        df["ciclo"].astype(str) + "-" +
        df["fase"].astype(str)
    )

    # Identificar exp_contables con mayor=1101
    exp_con_1101 = df.loc[df["mayor"] == 1101, "exp_contable"].unique()

    # Crear columnas ajustadas
    df["debe_adj"] = df.apply(
        lambda x: x["haber"] if x["exp_contable"] in exp_con_1101 else x["debe"],
        axis=1
    )
    df["haber_adj"] = df.apply(
        lambda x: x["debe"] if x["exp_contable"] in exp_con_1101 else x["haber"],
        axis=1
    )

    # Crear clave para equivalencias
    df["clave_cta"] = df["mayor"].astype(str) + "." + df["sub_cta"].astype(str)

    # -------------------------------
    # 2. Procesar archivo de equivalencias
    # -------------------------------
    xls_equiv = pd.ExcelFile(equiv_file)

    # Cargar Hoja de Trabajo
    df_equiv = pd.read_excel(xls_equiv, sheet_name="Hoja de Trabajo")

    # Normalizar valores
    df_equiv["Cuentas Contables"] = df_equiv["Cuentas Contables"].astype(str).str.strip()
    df_equiv["Rubros"] = df_equiv["Rubros"].astype(str).str.strip()

    # Evitar duplicados
    df_equiv = df_equiv.drop_duplicates(subset=["Cuentas Contables"], keep="first")

    # Merge con equivalencias
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    # -------------------------------
    # 3. Preparar hojas de salida
    # -------------------------------
    df_original = pd.read_excel(uploaded_file)

    df_tipo1 = df[df["tipo_ctb"] == 1]
    df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # -------------------------------
    # 4. Actualizar HT EF-4
    # -------------------------------
    try:
        df_ht = pd.read_excel(xls_equiv, sheet_name="HT EF-4")

        # Agrupamos sumas desde Tipo1_sin_1101
        agrupado = df_tipo1_sin1101.groupby("Rubros")[["debe_adj", "haber_adj"]].sum().reset_index()

        # Reemplazar columnas F y G en df_ht según Rubros
        for i, row in df_ht.iterrows():
            rubro = str(row.iloc[0]).strip()  # Columna A = Rubros
            if rubro in agrupado["Rubros"].values:
                valores = agrupado[agrupado["Rubros"] == rubro]
                df_ht.at[i, df_ht.columns[5]] = valores["debe_adj"].values[0]  # Columna F
                df_ht.at[i, df_ht.columns[6]] = valores["haber_adj"].values[0] # Columna G

    except Exception as e:
        st.error(f"No se pudo procesar HT EF-4: {e}")
        df_ht = None

    # -------------------------------
    # 5. Construir Excel final
    # -------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hojas del archivo principal
        df_original.to_excel(writer, index=False, sheet_name="Original")
        df.to_excel(writer, index=False, sheet_name="Resultado General")
        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # Guardar todas las hojas del archivo de equivalencias
        for sheet_name in xls_equiv.sheet_names:
            if sheet_name == "HT EF-4" and df_ht is not None:
                df_ht.to_excel(writer, index=False, sheet_name="HT EF-4")  # hoja modificada
            else:
                df_temp = pd.read_excel(xls_equiv, sheet_name=sheet_name)
                df_temp.to_excel(writer, index=False, sheet_name=sheet_name)

    # -------------------------------
    # 6. Botón de descarga
    # -------------------------------
    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
