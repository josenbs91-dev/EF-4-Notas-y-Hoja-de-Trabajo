import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from copy import copy

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and equiv_file:
    # -------------------------
    # 1. Cargar archivo principal
    # -------------------------
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Asegurar que debe, haber y saldo sean numéricos; lo demás texto
    for col in df.columns:
        if col.strip().lower() in ["debe", "haber", "saldo"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = df[col].astype(str).str.strip()

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
        lambda x: x["haber"] if x["exp_contable"] in exp_con_1101 else x["debe"],
        axis=1
    )
    df["haber_adj"] = df.apply(
        lambda x: x["debe"] if x["exp_contable"] in exp_con_1101 else x["haber"],
        axis=1
    )

    # Crear clave para equivalencias
    df["clave_cta"] = df["mayor"].astype(str) + "." + df["sub_cta"].astype(str)

    # -------------------------
    # 2. Cargar equivalencias
    # -------------------------
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

    # -------------------------
    # 3. Crear Excel en memoria
    # -------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Hoja Original
        df_original = pd.read_excel(uploaded_file, sheet_name=0)
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Resultado General
        df.to_excel(writer, index=False, sheet_name="Resultado General")

        # Filtrar tipo_ctb = 1
        df_tipo1 = df[df["tipo_ctb"] == "1"]

        df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
        df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

    # -------------------------
    # 4. Insertar hoja HT EF-4 con formato
    # -------------------------
    output.seek(0)
    book_result = load_workbook(output)
    book_equiv = load_workbook(equiv_file)

    if "HT EF-4" in book_equiv.sheetnames:
        sheet_equiv = book_equiv["HT EF-4"]

        # Eliminar si ya existe
        if "HT EF-4" in book_result.sheetnames:
            del book_result["HT EF-4"]
        sheet_result = book_result.create_sheet("HT EF-4")

        # Copiar dimensiones
        for col_letter, col_dim in sheet_equiv.column_dimensions.items():
            sheet_result.column_dimensions[col_letter].width = col_dim.width
        for row_idx, row_dim in sheet_equiv.row_dimensions.items():
            sheet_result.row_dimensions[row_idx].height = row_dim.height

        # Copiar valores y estilos
        for row in sheet_equiv.iter_rows():
            for cell in row:
                new_cell = sheet_result.cell(row=cell.row, column=cell.col_idx, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

        # Copiar celdas combinadas
        for merged_range in sheet_equiv.merged_cells.ranges:
            sheet_result.merge_cells(str(merged_range))

        # Calcular sumas por Rubros
        df_tipo1_sin1101_grouped = df_tipo1_sin1101.groupby("Rubros").agg(
            debe_total=("debe_adj", "sum"),
            haber_total=("haber_adj", "sum")
        ).reset_index()

        # Insertar sumas en columnas F (6) y G (7)
        correlativo = 1
        for row in range(1, sheet_result.max_row + 1):
            rubro = str(sheet_result.cell(row=row, column=1).value).strip()
            if rubro in df_tipo1_sin1101_grouped["Rubros"].values:
                valores = df_tipo1_sin1101_grouped[df_tipo1_sin1101_grouped["Rubros"] == rubro]
                debe_val = float(valores["debe_total"].values[0])
                haber_val = float(valores["haber_total"].values[0])

                sheet_result.cell(row=row, column=6, value=debe_val)   # F
                sheet_result.cell(row=row, column=7, value=haber_val)  # G
                sheet_result.cell(row=row, column=8, value=correlativo)  # Columna nueva AJUSTE
                correlativo += 1

    # Guardar cambios finales
    output_final = BytesIO()
    book_result.save(output_final)

    # -------------------------
    # 5. Botón de descarga
    # -------------------------
    st.download_button(
        label="Descargar resultado en Excel",
        data=output_final.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
