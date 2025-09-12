import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"])

# Subir archivo de equivalencias
uploaded_equiv = st.file_uploader("Sube tu archivo de equivalencias (Hoja de Trabajo)", type=["xlsx"])

if uploaded_file and uploaded_equiv:
    # =======================
    # Cargar archivos como TEXTO para mantener formatos
    # =======================
    df = pd.read_excel(uploaded_file, dtype=str)
    try:
        df_equiv = pd.read_excel(uploaded_equiv, sheet_name="Hoja de Trabajo", dtype=str)
    except Exception:
        st.error("El archivo de equivalencias no contiene una hoja llamada 'Hoja de Trabajo'.")
        st.stop()

    # =======================
    # Convertir columnas numéricas (debe, haber, saldo)
    # =======================
    for col in ["debe", "haber", "saldo"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    st.subheader("Vista previa de los datos originales")
    st.dataframe(df.head(20))

    # =======================
    # Crear exp_contable
    # =======================
    df["exp_contable"] = (
        df["ano_eje"].astype(str) + "-" +
        df["nro_not_exp"].astype(str) + "-" +
        df["ciclo"].astype(str) + "-" +
        df["fase"].astype(str)
    )

    # =======================
    # Identificar exp_contables con mayor=1101
    # =======================
    exp_con_1101 = df.loc[df["mayor"] == "1101", "exp_contable"].unique()

    # =======================
    # Crear columnas ajustadas (invertir solo si NO pertenece a mayor=1101)
    # =======================
    df["debe_adj"] = df.apply(
        lambda x: x["haber"] if x["exp_contable"] not in exp_con_1101 else x["debe"],
        axis=1
    )
    df["haber_adj"] = df.apply(
        lambda x: x["debe"] if x["exp_contable"] not in exp_con_1101 else x["haber"],
        axis=1
    )

    # =======================
    # Construir clave_cta con punto
    # =======================
    df["clave_cta"] = df["mayor"].map(str) + "." + df["sub_cta"].map(str)

    # =======================
    # Preparar equivalencias
    # =======================
    df_equiv["Cuentas Contables"] = df_equiv["Cuentas Contables"].map(str).str.strip()

    # =======================
    # Unir equivalencias (merge)
    # =======================
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    st.subheader("Resultado con exp_contable, debe/haber ajustados y Rubros")
    st.dataframe(df.head(20))

    # =======================
    # Filtrar tipo_ctb = 1 en dos hojas (con mayor=1101 y sin mayor=1101)
    # =======================
    df_tipo1 = df[df["tipo_ctb"] == "1"]

    df_tipo1_con_1101 = df_tipo1[df_tipo1["exp_contable"].isin(exp_con_1101)]
    df_tipo1_sin_1101 = df_tipo1[~df_tipo1["exp_contable"].isin(exp_con_1101)]

    # =======================
    # Crear Excel en memoria para descarga
    # =======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado General")
        df_tipo1_con_1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
        df_tipo1_sin_1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")

    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
