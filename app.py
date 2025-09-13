import streamlit as st
import pandas as pd

st.title("Procesamiento EF-4 con Equivalencias y HT EF-4")

# Subida de archivos
file = st.file_uploader("Sube tu archivo principal (con mayor, sub_cta, etc.)", type=["xls", "xlsx", "xlsm"])
equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo)", type=["xls", "xlsx", "xlsm"])

if file and equiv_file:
    # ==============================
    # 1. Cargar archivo principal
    # ==============================
    df = pd.read_excel(file, sheet_name=0, dtype=str)

    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip().str.lower()

    # Convertir debe/haber/saldo a numéricos si existen
    for col in ["debe", "haber", "saldo"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Crear clave_cta = mayor.sub_cta
    if "mayor" in df.columns and "sub_cta" in df.columns:
        df["clave_cta"] = df["mayor"].astype(str) + "." + df["sub_cta"].astype(str)

    # ==============================
    # 2. Cargar equivalencias
    # ==============================
    df_equiv = pd.read_excel(equiv_file, sheet_name="Hoja de Trabajo", dtype=str)
    df_equiv.columns = df_equiv.columns.str.strip()

    if "Cuentas Contables" in df_equiv.columns and "Rubros" in df_equiv.columns:
        df = df.merge(
            df_equiv[["Cuentas Contables", "Rubros"]],
            left_on="clave_cta",
            right_on="Cuentas Contables",
            how="left"
        )

    # ==============================
    # 3. Identificar exp_contables con mayor=1101
    # ==============================
    if "exp_contable" not in df.columns:
        df["exp_contable"] = (
            df.get("ano_eje", "").astype(str) + "-" +
            df.get("nro_not_exp", "").astype(str) + "-" +
            df.get("ciclo", "").astype(str) + "-" +
            df.get("fase", "").astype(str)
        )

    exp_con_1101 = df.loc[df["mayor"] == "1101", "exp_contable"].unique()

    # ==============================
    # 4. Ajuste de debe/haber
    # ==============================
    if "debe" in df.columns and "haber" in df.columns:
        df["debe_adj"] = df.apply(
            lambda x: x["haber"] if x["exp_contable"] not in exp_con_1101 else x["debe"],
            axis=1
        )
        df["haber_adj"] = df.apply(
            lambda x: x["debe"] if x["exp_contable"] not in exp_con_1101 else x["haber"],
            axis=1
        )

    # ==============================
    # 5. Dividir por tipo_ctb
    # ==============================
    df_tipo1 = df[df.get("tipo_ctb") == "1"]
    df_tipo1_con1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # ==============================
    # 6. Guardar resultados
    # ==============================
    with pd.ExcelWriter("resultado_final.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado General")
        df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

        # 🚨 Procesar HT EF-4 ya existente
        try:
            # Leemos todo sin encabezados
            df_raw = pd.read_excel(equiv_file, sheet_name="HT EF-4", header=None)

            # Buscar fila con "DESCRIPCION"
            header_row = df_raw.index[df_raw.apply(lambda r: r.astype(str).str.contains("DESCRIPCION", case=False).any(), axis=1)]
            if not header_row.empty:
                header_idx = header_row[0]
                df_ht = pd.read_excel(equiv_file, sheet_name="HT EF-4", header=header_idx)
            else:
                raise ValueError("No se encontró la fila con 'DESCRIPCION' en HT EF-4")

            # Agrupar importes desde Tipo1_sin_1101
            resumen = df_tipo1_sin1101.groupby("Rubros").agg(
                DEUDOR=("debe_adj", "sum"),
                ACREEDOR=("haber_adj", "sum")
            ).reset_index().rename(columns={"Rubros": "DESCRIPCION"})

            # Unir con HT EF-4
            if "DESCRIPCION" not in df_ht.columns:
                df_ht["DESCRIPCION"] = None

            df_ht = df_ht.merge(
                resumen,
                on="DESCRIPCION",
                how="left"
            )

            df_ht.to_excel(writer, index=False, sheet_name="HT EF-4")

        except Exception as e:
            st.warning(f"No se pudo procesar la hoja HT EF-4: {e}")

    st.success("Archivo resultado_final.xlsx generado con éxito")
    st.download_button("Descargar archivo procesado", data=open("resultado_final.xlsx", "rb").read(), file_name="resultado_final.xlsx")
