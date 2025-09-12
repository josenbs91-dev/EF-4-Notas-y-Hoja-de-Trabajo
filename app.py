import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Procesador de Exp Contable - SIAF con Equivalencias")

# Subir archivo principal
uploaded_file = st.file_uploader("Sube tu archivo Excel principal (movimientos)", type=["xlsx", "xlsm"])

# Subir archivo con equivalencias
uploaded_equiv = st.file_uploader("Sube tu archivo Excel con equivalencias (Hoja de Trabajo)", type=["xlsx", "xlsm"])

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

    st.subheader("Vista previa de los datos originales")
    st.dataframe(df.head(20))

    st.subheader("Vista previa de las equivalencias")
    st.dataframe(df_equiv.head(20))

    # =======================
    # Crear exp_contable (manteniendo formatos originales)
    # =======================
    df["exp_contable"] = (
        df["ano_eje"] + "-" +
        df["nro_not_exp"] + "-" +
        df["ciclo"] + "-" +
        df["fase"]
    )

    # =======================
    # Identificar exp_contables con mayor=1101
    # =======================
    exp_con_1101 = df.loc[df["mayor"] == "1101", "exp_contable"].unique()

    # =======================
    # Ajuste debe/haber
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
    # Validar columnas de equivalencias
    # =======================
    if "Cuentas Contables" not in df_equiv.columns or "Rubros" not in df_equiv.columns:
        st.error("La hoja 'Hoja de Trabajo' debe tener las columnas 'Cuentas Contables' y 'Rubros'.")
        st.stop()

    # =======================
    # Construir clave_cta manteniendo formato original
    # =======================
    df["clave_cta"] = df["mayor"] + "." + df["sub_cta"]

    # =======================
    # Unir equivalencias
    # =======================
    df = df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left"
    )

    # =======================
    # Dividir en dos hojas según condición
    # =======================
    df_ctb1_1101 = df[(df["tipo_ctb"] == "1") & (df["exp_contable"].isin(exp_con_1101))]
    df_ctb1_no1101 = df[(df["tipo_ctb"] == "1") & (~df["exp_contable"].isin(exp_con_1101))]

    st.subheader("Vista previa - tipo_ctb=1 y exp_con_1101")
    st.dataframe(df_ctb1_1101.head(20))

    st.subheader("Vista previa - tipo_ctb=1 y NO exp_con_1101")
    st.dataframe(df_ctb1_no1101.head(20))

    # =======================
    # Exportar a Excel (con formatos originales)
    # =======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado_Completo")
        df_ctb1_1101.to_excel(writer, index=False, sheet_name="TipoCTB1_1101")
        df_ctb1_no1101.to_excel(writer, index=False, sheet_name="TipoCTB1_No1101")

    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
