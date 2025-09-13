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
    df = pd.read_excel(uploaded_file)
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

    # Cargar equivalencias
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo")

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

    # Crear Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hoja original
        df_original = pd.read_excel(uploaded_file)
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Guardar resultado general
        df.to_excel(writer, index=False, sheet_name="Resultado General")

        # Filtrar tipo_ctb = 1
        df_tipo1 = df[df["tipo_ctb"] == 1]
        df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
        df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

    # --- Ahora actualizamos la hoja HT EF-4 del archivo de equivalencias ---
    wb = openpyxl.load_workbook(equiv_file)
    if "HT EF-4" in wb.sheetnames:
        ws = wb["HT EF-4"]

        # Agrupar debe_adj y haber_adj por Rubros desde Tipo1_sin_1101
        totales = df_tipo1_sin1101.groupby("Rubros")[["debe_adj", "haber_adj"]].sum().reset_index()

        # Recorrer la columna A de HT EF-4 (Rubros)
        for row in range(2, ws.max_row + 1):  # Asumimos que fila 1 son encabezados
            rubro = str(ws.cell(row=row, column=1).value).strip() if ws.cell(row=row, column=1).value else None
            if rubro and rubro in totales["Rubros"].values:
                fila = totales[totales["Rubros"] == rubro].iloc[0]
                ws.cell(row=row, column=6, value=float(fila["debe_adj"]))   # Columna F
                ws.cell(row=row, column=7, value=float(fila["haber_adj"]))  # Columna G

        # Guardar cambios en memoria junto con lo demás
        wb.save(output)

    # Botón de descarga
    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
