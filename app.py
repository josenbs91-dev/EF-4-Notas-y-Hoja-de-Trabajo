import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Procesador de Exp Contable - SIAF")

# Subir archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Vista previa de los datos originales")
    st.dataframe(df.head(20))

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
        lambda x: x["haber"] if x["exp_contable"] not in exp_con_1101 else x["debe"],
        axis=1
    )
    df["haber_adj"] = df.apply(
        lambda x: x["debe"] if x["exp_contable"] not in exp_con_1101 else x["haber"],
        axis=1
    )

    st.subheader("Resultado con exp_contable y debe/haber ajustados")
    st.dataframe(df.head(20))

    # Filtrar registros con tipo_ctb = 1
    df_tipo_ctb1 = df[df["tipo_ctb"] == 1]

    # Crear Excel en memoria para descarga con dos hojas
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
        df_tipo_ctb1.to_excel(writer, index=False, sheet_name="Tipo_CTB_1")

    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_exp_contable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
