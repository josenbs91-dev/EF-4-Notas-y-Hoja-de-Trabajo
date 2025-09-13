import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and equiv_file:
    # ------------------------------
    # 1. Cargar archivo principal
    # ------------------------------
    df_raw = pd.read_excel(uploaded_file, dtype=str)

    # Forzar debe, haber y saldo a numérico
    for col in ["debe", "haber", "saldo"]:
        if col in df_raw.columns:
            df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce")

    df = df_raw.copy()

    # Crear exp_contable
    df["exp_contable"] = (
        df["ano_eje"].astype(str) + "-" +
        df["nro_not_exp"].astype(str) + "-" +
        df["ciclo"].astype(str) + "-" +
        df["fase"].astype(str)
    )

    # Identificar exp_contables con mayor=1101
    exp_con_1101 = df.loc[df["mayor"].astype(str) == "1101", "exp_contable"].unique()

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

    # ------------------------------
    # 2. Cargar equivalencias
    # ------------------------------
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo", dtype=str)
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

    # ------------------------------
    # 3. Crear Excel en memoria
    # ------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hoja original cargada
        df_original = pd.read_excel(uploaded_file, dtype=str)
        for col in ["debe", "haber", "saldo"]:
            if col in df_original.columns:
                df_original[col] = pd.to_numeric(df_original[col], errors="coerce")
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Guardar resultado general con equivalencias
        df.to_excel(writer, index=False, sheet_name="Resultado General")

        # Filtrar tipo_ctb = 1 y separar en dos hojas
        df_tipo1 = df[df["tipo_ctb"].astype(str) == "1"]

        df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
        df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # Guardar equivalencias (Hoja de Trabajo)
        df_equiv.to_excel(writer, index=False, sheet_name="Equivalencias")

    # ------------------------------
    # 4. Actualizar hoja HT EF-4 (manteniendo formato original)
    # ------------------------------
    output.seek(0)
    book_result = load_workbook(output)
    book_equiv = load_workbook(equiv_file)

    # Copiar hoja HT EF-4 del archivo equivalencias al resultado final
    if "HT EF-4" in book_equiv.sheetnames:
        sheet_equiv = book_equiv["HT EF-4"]
        sheet_result = book_result.copy_worksheet(sheet_equiv)
        sheet_result.title = "HT EF-4"

        # Calcular sumas por Rubros en Tipo1_sin_1101
        df_tipo1_sin1101_grouped = df_tipo1_sin1101.groupby("Rubros").agg(
            debe_total=("debe_adj", "sum"),
            haber_total=("haber_adj", "sum")
        ).reset_index()

        # Insertar valores en columnas F (DEUDOR) y G (ACREEDOR)
        for row in range(1, sheet_result.max_row + 1):
            rubro = str(sheet_result.cell(row=row, column=1).value).strip()
            if rubro in df_tipo1_sin1101_grouped["Rubros"].values:
                valores = df_tipo1_sin1101_grouped[
                    df_tipo1_sin1101_grouped["Rubros"] == rubro
                ]
                debe_val = float(valores["debe_total"].values[0])
                haber_val = float(valores["haber_total"].values[0])

                sheet_result.cell(row=row, column=6, value=debe_val)   # Columna F
                sheet_result.cell(row=row, column=7, value=haber_val)  # Columna G

    # Guardar el archivo final en memoria
    final_output = BytesIO()
    book_result.save(final_output)

    # ------------------------------
    # 5. Botón de descarga
    # ------------------------------
    st.download_button(
        label="Descargar resultado en Excel",
        data=final_output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
