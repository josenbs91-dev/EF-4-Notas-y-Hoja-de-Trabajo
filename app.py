import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =============================
# Configuración básica
# =============================
st.set_page_config(page_title="Procesador de Exp Contable - SIAF", layout="wide")
st.title("Procesador de Exp Contable - SIAF (Excel simplificado)")

# Constantes
REQUIRED_FOR_EXP = ["ano_eje", "nro_not_exp", "ciclo", "fase"]
NUMERIC_COLS = ["debe", "haber", "saldo"]
EQUIV_SHEET = "Hoja de Trabajo"

# =============================
# Utilidades
# =============================
@st.cache_data(show_spinner=False)
def _read_file_bytes(uploaded_file) -> bytes:
    return uploaded_file.read()

@st.cache_data(show_spinner=False)
def load_main_df(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), dtype=str, engine="openpyxl")
    # Normalizamos nombres de columnas
    df.columns = [c.strip().lower() for c in df.columns]

    # Coerción numérica segura
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Validación de columnas requeridas
    missing = [c for c in REQUIRED_FOR_EXP if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas en el archivo principal: {', '.join(missing)}"
        )

    # Construcción de exp_contable
    parts = [df[c].astype(str).fillna("") for c in REQUIRED_FOR_EXP]
    df["exp_contable"] = parts[0] + "-" + parts[1] + "-" + parts[2] + "-" + parts[3]

    # Clave contable
    mayor = df.get("mayor", "").astype(str)
    sub_cta = df.get("sub_cta", "").astype(str)
    df["clave_cta"] = mayor.str.strip() + "." + sub_cta.str.strip()

    return df

@st.cache_data(show_spinner=False)
def load_equiv_df(file_bytes: bytes) -> pd.DataFrame:
    try:
        df_e = pd.read_excel(BytesIO(file_bytes), sheet_name=EQUIV_SHEET, dtype=str, engine="openpyxl")
    except ValueError as e:
        raise ValueError(f"No se encontró la hoja '{EQUIV_SHEET}' en el archivo de Equivalencias.") from e

    df_e.columns = [str(c).strip() for c in df_e.columns]
    required_cols = {"Cuentas Contables", "Rubros"}
    if not required_cols.issubset(df_e.columns):
        raise ValueError("La hoja de Equivalencias debe contener las columnas 'Cuentas Contables' y 'Rubros'.")

    df_e["Cuentas Contables"] = df_e["Cuentas Contables"].astype(str).str.strip()
    df_e["Rubros"] = df_e["Rubros"].astype(str).str.strip()
    df_e = df_e.drop_duplicates(subset=["Cuentas Contables"], keep="first").reset_index(drop=True)
    return df_e

def compute_adjusted(df: pd.DataFrame) -> pd.DataFrame:
    """Ajusta debe/haber si el exp_contable pertenece a un mayor 1101."""
    mask_1101 = df.get("mayor", "").astype(str).eq("1101")
    exp_con_1101 = set(df.loc[mask_1101, "exp_contable"].unique())
    in_1101 = df["exp_contable"].isin(exp_con_1101)

    debe = pd.to_numeric(df.get("debe", 0), errors="coerce").fillna(0.0)
    haber = pd.to_numeric(df.get("haber", 0), errors="coerce").fillna(0.0)

    df["debe_adj"] = np.where(in_1101, haber, debe)
    df["haber_adj"] = np.where(in_1101, debe, haber)

    # Tipos finales
    for col in ["debe_adj", "haber_adj", "debe", "haber", "saldo"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    for col in df.columns:
        if col not in ["debe_adj", "haber_adj", "debe", "haber", "saldo"]:
            df[col] = df[col].astype(str)

    return df

def merge_equivalences(df: pd.DataFrame, df_equiv: pd.DataFrame) -> pd.DataFrame:
    return df.merge(
        df_equiv[["Cuentas Contables", "Rubros"]],
        left_on="clave_cta",
        right_on="Cuentas Contables",
        how="left",
    )

def build_simple_excel(main_bytes: bytes, df_result: pd.DataFrame) -> BytesIO:
    """Genera un Excel liviano con hojas: Original, Resultado General y particiones tipo_ctb."""
    output = BytesIO()

    # ⚠️ Importante: NO pasar 'options' al ExcelWriter (evita el error reportado)
    # Requiere tener instalado 'xlsxwriter' para mejor rendimiento de escritura.
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Hoja Original (sin coerción extra para velocidad)
            df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
            df_original.to_excel(writer, index=False, sheet_name="Original")

            # Resultado General
            df_result.to_excel(writer, index=False, sheet_name="Resultado General")

            # Particiones tipo_ctb == 1
            if "tipo_ctb" in df_result.columns:
                df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()

                in_1101 = df_tipo1["exp_contable"].isin(
                    set(df_result.loc[df_result.get("mayor", "").astype(str).eq("1101"), "exp_contable"].unique())
                )
                df_tipo1_con1101 = df_tipo1[in_1101].copy()
                df_tipo1_sin1101 = df_tipo1[~in_1101].copy()

                df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
                df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
            else:
                pd.DataFrame(
                    {"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}
                ).to_excel(writer, index=False, sheet_name="Avisos")
    except Exception:
        # Fallback: si no tienes xlsxwriter instalado, usa openpyxl
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
            df_original.to_excel(writer, index=False, sheet_name="Original")
            df_result.to_excel(writer, index=False, sheet_name="Resultado General")
            if "tipo_ctb" in df_result.columns:
                df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()
                in_1101 = df_tipo1["exp_contable"].isin(
                    set(df_result.loc[df_result.get("mayor", "").astype(str).eq("1101"), "exp_contable"].unique())
                )
                df_tipo1_con1101 = df_tipo1[in_1101].copy()
                df_tipo1_sin1101 = df_tipo1[~in_1101].copy()
                df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
                df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
            else:
                pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                    writer, index=False, sheet_name="Avisos"
                )

    output.seek(0)
    return output

# =============================
# UI
# =============================
col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"], key="main")
with col2:
    equiv_file = st.file_uploader(f"Sube tu archivo de Equivalencias ({EQUIV_SHEET})", type=["xlsx"], key="equiv")

if uploaded_file and equiv_file:
    try:
        main_bytes = _read_file_bytes(uploaded_file)
        equiv_bytes = _read_file_bytes(equiv_file)

        # Carga y procesamiento
        df_main = load_main_df(main_bytes)
        df_equiv = load_equiv_df(equiv_bytes)
        df_proc = compute_adjusted(df_main.copy())
        df_final = merge_equivalences(df_proc, df_equiv)

        # Excel simplificado
        xls_bytes = build_simple_excel(main_bytes, df_final)

        st.success("Procesamiento completado.")
        st.download_button(
            label="Descargar resultado en Excel (simplificado)",
            data=xls_bytes.getvalue(),
            file_name="resultado_exp_contable_simplificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Ver vista previa (primeras 200 filas)"):
            st.dataframe(df_final.head(200))

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")

else:
    st.info("Sube ambos archivos para iniciar el procesamiento.")
