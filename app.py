import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and equiv_file:
    # Cargar archivo principal
    df = pd.read_excel(uploaded_file, dtype=str).copy()

    # Normalizar columnas
    df.columns = df.columns.str.strip().str.lower()

    # Convertir campos numéricos
    for col in ["debe", "haber", "saldo"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Crear exp_contable
    df["exp_contable"] = (
        df["ano_eje"].astype(str) + "-" +
        df["nro_not_exp"].astype(str) + "-" +
        df["ciclo"].astype(str) + "-" +
        df["fase"].astype(str)
    )

    # Identificar exp_contables con mayor=1101
    exp_con_1101 = df.loc[df["mayor"] == "1101", "exp_contable"].unique()

    # Crear columnas ajustadas
    df["debe_adj"] = df.apply(
        lambda x: x["haber"] if x["exp_contable"] not in exp_con_1101 else x["debe"],
        axis=1
    )
    df["haber_adj"] = df.apply(
        lambda x: x["debe"] if x["exp_contable"] not in exp_con_1101 else x["haber"],
        axis=1
    )

    # Crear clave para equivalencias
    if "mayor" in df.columns and "sub_cta" in df.columns:
        df["clave_cta"] = df["mayor"].astype(str) + "." + df["sub_cta"].astype(str)

    # Cargar equivalencias
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo")
    df_equiv["Cuentas Contables"] = df_equiv["Cuentas Contables"].astype(str).str.strip()
    df_equiv["Rubros"] = df_equiv["Rubros"].astype(str).str.strip()
    df_equiv = df_equiv.drop_duplicates(subset=["Cuentas Contables"], keep="first")

    # Merge con equivalencias
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    # Filtrar tipo_ctb = 1 y separar en dos hojas
    df_tipo1 = df[df["tipo_ctb"] == "1"]
    df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # Agrupación por Rubros desde Tipo1_sin_1101
    resumen = df_tipo1_sin1101.groupby("Rubros").agg(
        debe_total=("debe_adj", "sum"),
        haber_total=("haber_adj", "sum")
    ).reset_index()

    # Cargar hoja HT EF-4 del archivo de equivalencias
    wb_equiv = openpyxl.load_workbook(equiv_file)
    if "HT EF-4" in wb_equiv.sheetnames:
        ws = wb_equiv["HT EF-4"]

        # Buscar Rubros en columna A y reemplazar en columnas F y G
        for row in range(2, ws.max_row + 1):  # asumiendo encabezados en fila 1
            rubro = str(ws.cell(row=row, column=1).value).strip()
            if rubro in resumen["Rubros"].values:
                valor_debe = resumen.loc[resumen["Rubros"] == rubro, "debe_total"].values[0]
                valor_haber = resumen.loc[resumen["Rubros"] == rubro, "haber_total"].values[0]
                ws.cell(row=row, column=6, value=valor_debe)  # Columna F
                ws.cell(row=row, column=7, value=valor_haber)  # Columna G

    # Guardar todo en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hoja original
        df_original = pd.read_excel(uploaded_file)
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Guardar resultado general
        df.to_excel(writer, index=False, sheet_name="Resultado General")

        # Guardar hojas filtradas
        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # Guardar copia de HT EF-4 actualizada como hoja adicional
        if "HT EF-4" in wb_equiv.sheetnames:
            data_ht = pd.DataFrame(ws.values)
            data_ht.to_excel(writer, index=False, header=False, sheet_name="HT EF-4")

    # Botón de descarga
    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
