import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

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
    # 2. Cargar equivalencias
    # -------------------------------
    xls_equiv = pd.ExcelFile(equiv_file)
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo")

    # Normalizar
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

    # -------------------------------
    # 3. Preparar resultados
    # -------------------------------
    df_tipo1 = df[df["tipo_ctb"] == 1]
    df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # Agrupar para HT EF-4
    agrupado = df_tipo1_sin1101.groupby("Rubros").agg({
        "debe_adj": "sum",
        "haber_adj": "sum"
    }).reset_index()

    # -------------------------------
    # 4. Modificar HT EF-4 con openpyxl
    # -------------------------------
    wb_equiv = load_workbook(equiv_file)
    if "HT EF-4" in wb_equiv.sheetnames:
        ws_ht = wb_equiv["HT EF-4"]

        # Iterar filas de la columna A (Rubros) y asignar valores en F y G
        for row in range(2, ws_ht.max_row + 1):  # desde fila 2 (asumo fila 1 cabecera)
            rubro = str(ws_ht.cell(row=row, column=1).value).strip()
            if rubro in list(agrupado["Rubros"]):
                valor_debe = float(agrupado.loc[agrupado["Rubros"] == rubro, "debe_adj"].values[0])
                valor_haber = float(agrupado.loc[agrupado["Rubros"] == rubro, "haber_adj"].values[0])
                ws_ht.cell(row=row, column=6, value=valor_debe)  # Columna F = 6
                ws_ht.cell(row=row, column=7, value=valor_haber)  # Columna G = 7

        # Insertar columna de AJUSTE al final (después de la última columna)
        last_col = ws_ht.max_column + 1
        ws_ht.cell(row=1, column=last_col, value="AJUSTE")
        for row in range(2, ws_ht.max_row + 1):
            ws_ht.cell(row=row, column=last_col, value=row - 1)

    # -------------------------------
    # 5. Generar archivo de salida
    # -------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Guardar hojas procesadas
        df_original = pd.read_excel(uploaded_file)
        df_original.to_excel(writer, index=False, sheet_name="Original")
        df.to_excel(writer, index=False, sheet_name="Resultado General")
        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # Guardar todas las hojas de equivalencias tal cual (incluyendo HT EF-4 modificada)
        for sheet_name in wb_equiv.sheetnames:
            ws = wb_equiv[sheet_name]
            data = ws.values
            cols = next(data)
            df_temp = pd.DataFrame(data, columns=cols)
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
