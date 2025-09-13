import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and equiv_file:
    # Cargar archivo principal
    df = pd.read_excel(uploaded_file)

    # Mantener formato original de los datos
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

    # 🚨 Cargar equivalencias con filtro
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo")

    # Normalizar
    df_equiv["Cuentas Contables"] = df_equiv["Cuentas Contables"].astype(str).str.strip()

    # Filtrar filas vacías y separadores
    df_equiv = df_equiv[df_equiv["Cuentas Contables"].notna()]
    df_equiv = df_equiv[~df_equiv["Cuentas Contables"].str.contains("TOTAL|---|^nan$", case=False, na=False)]

    # Eliminar duplicados
    df_equiv = df_equiv.drop_duplicates(subset=["Cuentas Contables"], keep="first")

    # Merge con equivalencias
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    # Crear Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hoja original cargada
        df_original = pd.read_excel(uploaded_file)
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Guardar resultado general con equivalencias
        df.to_excel(writer, index=False, sheet_name="Resultado General")

        # Filtrar tipo_ctb = 1 y separar en dos hojas
        df_tipo1 = df[df["tipo_ctb"] == 1]

        df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
        df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # 🚨 Copiar hoja de equivalencias y HT EF-4 tal cual están
        wb_equiv = load_workbook(equiv_file)

        for sheet_name in wb_equiv.sheetnames:
            ws = wb_equiv[sheet_name]
            new_ws = writer.book.create_sheet(title=sheet_name)

            for row in ws.iter_rows(values_only=False):
                new_row = [cell.value for cell in row]
                new_ws.append(new_row)

    # Botón de descarga
    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
