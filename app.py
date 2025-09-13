import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo + HT EF-4)", type=["xlsx"])

if uploaded_file and equiv_file:
    # -------------------------------
    # 1. Cargar archivo principal
    # -------------------------------
    df = pd.read_excel(uploaded_file)
    df = df.copy()  # mantener formato original

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
    # 2. Cargar archivo equivalencias
    # -------------------------------
    xls_equiv = pd.ExcelFile(equiv_file)
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo")

    # Normalizar valores
    df_equiv["Cuentas Contables"] = df_equiv["Cuentas Contables"].astype(str).str.strip()
    df_equiv["Rubros"] = df_equiv["Rubros"].astype(str).str.strip()

    # Evitar duplicados en equivalencias
    df_equiv = df_equiv.drop_duplicates(subset=["Cuentas Contables"], keep="first")

    # Merge con equivalencias
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    # -------------------------------
    # 3. Preparar resultados
    # -------------------------------
    df_tipo1 = df[df["tipo_ctb"] == 1]
    df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # -------------------------------
    # 4. Actualizar HT EF-4
    # -------------------------------
    if "HT EF-4" in xls_equiv.sheet_names:
        df_ht = pd.read_excel(equiv_file, sheet_name="HT EF-4")

        # Agrupar por rubros desde Tipo1_sin_1101
        agrupado = df_tipo1_sin1101.groupby("Rubros").agg({
            "debe_adj": "sum",
            "haber_adj": "sum"
        }).reset_index()

        # Reemplazar valores en HT EF-4
        if "Rubros" in df.columns:
            df_ht["DEUDOR"] = df_ht.iloc[:, 0].map(
                agrupado.set_index("Rubros")["debe_adj"]
            ).fillna(0)
            df_ht["ACREEDOR"] = df_ht.iloc[:, 0].map(
                agrupado.set_index("Rubros")["haber_adj"]
            ).fillna(0)

        # Insertar columna de AJUSTE con correlativos
        df_ht["AJUSTE"] = range(1, len(df_ht) + 1)
    else:
        df_ht = None

    # -------------------------------
    # 5. Generar archivo de salida
    # -------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hojas generadas
        df_original = pd.read_excel(uploaded_file)
        df_original.to_excel(writer, index=False, sheet_name="Original")
        df.to_excel(writer, index=False, sheet_name="Resultado General")
        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # Copiar todas las hojas de equivalencias
        for sheet_name in xls_equiv.sheet_names:
            df_temp = pd.read_excel(equiv_file, sheet_name=sheet_name)
            if sheet_name == "HT EF-4" and df_ht is not None:
                df_ht.to_excel(writer, index=False, sheet_name=sheet_name)
            else:
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
