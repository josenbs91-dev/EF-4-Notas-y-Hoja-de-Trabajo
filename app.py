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
    # Cargar archivos
    # =======================
    df = pd.read_excel(uploaded_file)

    # Buscar la hoja "Hoja de Trabajo" en el archivo de equivalencias
    try:
        df_equiv = pd.read_excel(uploaded_equiv, sheet_name="Hoja de Trabajo")
    except Exception:
        st.error("El archivo de equivalencias no contiene una hoja llamada 'Hoja de Trabajo'.")
        st.stop()

    st.subheader("Vista previa de los datos originales")
    st.dataframe(df.head(20))

    st.subheader("Vista previa de las equivalencias")
    st.dataframe(df_equiv.head(20))

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
    exp_con_1101 = df.loc[df["mayor"] == 1101, "exp_contable"].unique()

    # =======================
    # Ajuste debe/haber
    # (solo se invierte si NO pertenece a exp_con_1101)
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
    # Crear clave de unión mayor.sub_cta
    # =======================
    df["clave_cta"] = df["mayor"].astype(str) + "." + df["sub_cta"].astype(str)

    # Aseguramos que la hoja de equivalencias tenga esas columnas
    if "Cuentas Contables" not in df_equiv.columns or "Rubros" not in df_equiv.columns:
        st.error("La hoja 'Hoja de Trabajo' debe tener las columnas 'Cuentas Contables' y 'Rubros'.")
        st.stop()

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
    # 1. tipo_ctb = 1 y pertenece a exp_con_1101
    df_ctb1_1101 = df[(df["tipo_ctb"] == 1) & (df["exp_contable"].isin(exp_con_1101))]

    # 2. tipo_ctb = 1 y NO pertenece a exp_con_1101
    df_ctb1_no1101 = df[(df["tipo_ctb"] == 1) & (~df["exp_contable"].isin(exp_con_1101))]

    st.subheader("Vista previa - tipo_ctb=1 y exp_con_1101")
    st.dataframe(df_ctb1_1101.head(20))

    st.subheader("Vista previa - tipo_ctb=1 y NO exp_con_1101")
    st.dataframe(df_ctb1_no1101.head(20))

    # =======================
    # Exportar a Excel
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
