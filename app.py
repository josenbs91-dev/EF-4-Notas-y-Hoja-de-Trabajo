import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from copy import copy

st.title("Procesamiento de Archivos Contables")

# Subida de archivos
uploaded_file = st.file_uploader("Sube tu archivo contable", type=["xlsx"])
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and equiv_file:
    try:
        # Leer archivo contable
        df = pd.read_excel(uploaded_file, sheet_name="exp_contable")

        # Asegurar tipos de datos
        for col in df.columns:
            if col in ["debe", "haber", "saldo", "debe_adj", "haber_adj"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = df[col].astype(str)

        # Proceso: separar en dos hojas
        df_tipo1_con1101 = df[df["mayor"] == "1101"].copy()
        df_tipo1_sin1101 = df[df["mayor"] != "1101"].copy()

        # Guardar resultados iniciales en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Original", index=False)
            df_tipo1_con1101.to_excel(writer, sheet_name="Resultado_Tipo1_con_1101", index=False)
            df_tipo1_sin1101.to_excel(writer, sheet_name="Resultado_Tipo1_sin_1101", index=False)

        # Reabrir libro para seguir editando
        output.seek(0)
        book_result = load_workbook(output)
        book_equiv = load_workbook(equiv_file)

        # Copiar hoja HT EF-4 con formato original
        if "HT EF-4" in book_equiv.sheetnames:
            sheet_equiv = book_equiv["HT EF-4"]

            # Borrar si ya existía
            if "HT EF-4" in book_result.sheetnames:
                del book_result["HT EF-4"]
            sheet_result = book_result.create_sheet("HT EF-4")

            # Copiar dimensiones
            for col_letter, col_dim in sheet_equiv.column_dimensions.items():
                sheet_result.column_dimensions[col_letter].width = col_dim.width
            for row_idx, row_dim in sheet_equiv.row_dimensions.items():
                sheet_result.row_dimensions[row_idx].height = row_dim.height

            # Copiar valores + estilos
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

            # Calcular sumas por Rubros en Tipo1_sin_1101
            df_tipo1_sin1101_grouped = df_tipo1_sin1101.groupby("Rubros").agg(
                debe_total=("debe_adj", "sum"),
                haber_total=("haber_adj", "sum")
            ).reset_index()

            # Insertar valores en columnas F (6) y G (7)
            for row in range(1, sheet_result.max_row + 1):
                rubro = str(sheet_result.cell(row=row, column=1).value).strip()
                if rubro in df_tipo1_sin1101_grouped["Rubros"].values:
                    valores = df_tipo1_sin1101_grouped[
                        df_tipo1_sin1101_grouped["Rubros"] == rubro
                    ]
                    debe_val = float(valores["debe_total"].values[0])
                    haber_val = float(valores["haber_total"].values[0])

                    # F (Debe) y G (Haber) en formato numérico
                    c_debe = sheet_result.cell(row=row, column=6, value=debe_val)
                    c_haber = sheet_result.cell(row=row, column=7, value=haber_val)
                    c_debe.number_format = "#,##0.00"
                    c_haber.number_format = "#,##0.00"

        # Agregar también hoja Equivalencias completa
        for sheetname in book_equiv.sheetnames:
            if sheetname != "HT EF-4":  # ya la copiamos arriba
                if sheetname in book_result.sheetnames:
                    del book_result[sheetname]
                sheet_equiv = book_equiv[sheetname]
                sheet_copy = book_result.create_sheet(sheetname)

                for col_letter, col_dim in sheet_equiv.column_dimensions.items():
                    sheet_copy.column_dimensions[col_letter].width = col_dim.width
                for row_idx, row_dim in sheet_equiv.row_dimensions.items():
                    sheet_copy.row_dimensions[row_idx].height = row_dim.height

                for row in sheet_equiv.iter_rows():
                    for cell in row:
                        new_cell = sheet_copy.cell(row=cell.row, column=cell.col_idx, value=cell.value)
                        if cell.has_style:
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.number_format = copy(cell.number_format)
                            new_cell.protection = copy(cell.protection)
                            new_cell.alignment = copy(cell.alignment)

                for merged_range in sheet_equiv.merged_cells.ranges:
                    sheet_copy.merge_cells(str(merged_range))

        # Guardar archivo final
        output_final = io.BytesIO()
        book_result.save(output_final)
        output_final.seek(0)

        st.success("Proceso completado con éxito.")
        st.download_button(
            label="Descargar Resultado en Excel",
            data=output_final,
            file_name="resultado_exp_contable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
