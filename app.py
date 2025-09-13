# Crear Excel en memoria
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # Guardar hoja original cargada
    df_original = pd.read_excel(uploaded_file, dtype=str)
    for col in ["debe", "haber", "saldo"]:
        if col in df_original.columns:
            df_original[col] = pd.to_numeric(df_original[col], errors="coerce").fillna(0)
    df_original.to_excel(writer, index=False, sheet_name="Original")

    # Guardar resultado general
    df.to_excel(writer, index=False, sheet_name="Resultado General")

    # Filtrar tipo_ctb = 1
    df_tipo1_con1101, df_tipo1_sin1101 = pd.DataFrame(), pd.DataFrame()
    if "tipo_ctb" in df.columns:
        df_tipo1 = df[df["tipo_ctb"].astype(str) == "1"]

        df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)].copy()
        df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)].copy()

        for dfx in [df_tipo1_con1101, df_tipo1_sin1101]:
            for col in numeric_cols_result:
                if col in dfx.columns:
                    dfx[col] = pd.to_numeric(dfx[col], errors="coerce").fillna(0)
            for col in dfx.columns:
                if col not in numeric_cols_result:
                    dfx[col] = dfx[col].astype(str)

        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

    # === Copiar hoja HT EF-4 y sumar rubros ===
    book_equiv = openpyxl.load_workbook(equiv_file)
    book_result = writer.book

    if "HT EF-4" in book_equiv.sheetnames:
        # Copiar directamente la hoja
        sheet_equiv = book_equiv["HT EF-4"]
        sheet_copy = book_result.copy_worksheet(sheet_equiv)
        sheet_copy.title = "HT EF-4"

        # Crear DataFrame auxiliar con sumatorias por rubro
        if not df_tipo1_sin1101.empty and 'Rubros' in df_tipo1_sin1101.columns:
            df_tipo1_sin1101_sum = df_tipo1_sin1101.groupby("Rubros")[["debe_adj", "haber_adj"]].sum().reset_index()
            dict_debe = dict(zip(df_tipo1_sin1101_sum["Rubros"], df_tipo1_sin1101_sum["debe_adj"]))
            dict_haber = dict(zip(df_tipo1_sin1101_sum["Rubros"], df_tipo1_sin1101_sum["haber_adj"]))

            # Iterar filas desde la 2 (asumiendo encabezado en fila 1)
            for row in sheet_copy.iter_rows(min_row=2):
                rubro = str(row[1].value).strip()  # Columna B = índice 1
                if rubro:
                    debe_sum = dict_debe.get(rubro, 0)
                    haber_sum = dict_haber.get(rubro, 0)

                    # Escribir en columnas G (índice 6) y H (índice 7)
                    row[6].value = debe_sum
                    row[7].value = haber_sum

# Botón de descarga
st.download_button(
    label="Descargar resultado en Excel",
    data=output.getvalue(),
    file_name="resultado_exp_contable.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
