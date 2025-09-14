import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from copy import copy

# =============================
# Configuración básica
# =============================
st.set_page_config(page_title="Procesador de Exp Contable - SIAF", layout="wide")
st.title("Procesador de Exp Contable - SIAF")

# Constantes
REQUIRED_FOR_EXP = ["ano_eje", "nro_not_exp", "ciclo", "fase"]
NUMERIC_COLS = ["debe", "haber", "saldo"]
EQUIV_SHEET = "Hoja de Trabajo"
COPIABLE_SHEET = "HT EF-4"

# =============================
# Utilidades / Helpers
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

def copy_sheet_with_styles(src_ws: openpyxl.worksheet.worksheet.Worksheet,
                           dst_ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Copia valores, estilos y rangos combinados."""
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

def is_inside_merged_area(row: int, col: int, merged_ranges) -> bool:
    for rng in merged_ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return True
    return False

# =============================
# NUEVA hoja: HT EF-4 (Compilada) a partir de EF-1 Apertura y EF-1 Final
# =============================
def write_ht_ef4_compilada(writer, equiv_bytes: bytes, sheet_name: str = "HT EF-4 (Compilada)"):
    """
    Construye una hoja con estructura estilo HT EF-4:
      - Sección "EF-1 Apertura" (si existe): Rubro en columna B, desagregado de Cuentas Contables en columna C (y Descripción en D si existe).
      - Sección "EF-1 Final" (si existe), inmediatamente después.
      - Columnas E/F (o 'Debe'/'Haber') quedan disponibles en blanco para usos posteriores.
    """
    # Cargar hojas si existen
    sections = []
    for label in ["EF-1 Apertura", "EF-1 Final"]:
        try:
            df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=label, dtype=str, engine="openpyxl")
            df_sec.columns = [str(c).strip() for c in df_sec.columns]
            if not {"Rubros", "Cuentas Contables"}.issubset(df_sec.columns):
                sections.append((label, None, f"La hoja '{label}' no contiene columnas 'Rubros' y 'Cuentas Contables'."))
                continue
            # elegir columna de descripción (opcional)
            desc_col = next((c for c in ["Descripción", "Descripcion", "DESCRIPCION", "Nombre", "Detalle", "Glosa"] if c in df_sec.columns), None)
            df_norm = pd.DataFrame({
                "Rubros": df_sec["Rubros"].astype(str).str.strip(),
                "Cuentas Contables": df_sec["Cuentas Contables"].astype(str).str.strip(),
                "Descripción": df_sec[desc_col].astype(str).str.strip() if desc_col else ""
            }).drop_duplicates().reset_index(drop=True)
            sections.append((label, df_norm, None))
        except Exception:
            sections.append((label, None, f"No se encontró la hoja '{label}'."))

    # Si ninguna existe, crear aviso
    if all(df is None for _, df, _ in sections):
        pd.DataFrame({"info": ["No se encontraron 'EF-1 Apertura' ni 'EF-1 Final' en el archivo de Equivalencias."]}) \
            .to_excel(writer, index=False, sheet_name=sheet_name)
        return

    # Construir filas con estructura: [ (colA vacía), Rubros(B), Cuenta(C), Descripción(D), Debe(E), Haber(F) ]
    rows = []
    def add_section_rows(label: str, df_sec: pd.DataFrame, err: str | None):
        # Título de sección
        rows.append(["", label, "", "", "", ""])
        if err:
            rows.append(["", err, "", "", "", ""])
            rows.append(["", "", "", "", "", ""])
            return
        if df_sec is None or df_sec.empty:
            rows.append(["", f"(sin datos en {label})", "", "", "", ""])
            rows.append(["", "", "", "", "", ""])
            return
        # Ordenar por Rubros y cuenta
        for rubro, g in df_sec.groupby("Rubros"):
            rubro = str(rubro).strip()
            rows.append(["", rubro, "", "", "", ""])
            g_sorted = g.sort_values(["Rubros", "Cuentas Contables", "Descripción"], na_position="last")
            for _, r in g_sorted.iterrows():
                cuenta = str(r["Cuentas Contables"]).strip()
                desc = str(r.get("Descripción", "")).strip()
                rows.append(["", "", cuenta, desc, "", ""])
        rows.append(["", "", "", "", "", ""])  # línea en blanco entre secciones

    # Apertura primero, luego Final
    for label in ["EF-1 Apertura", "EF-1 Final"]:
        for lab, df_sec, err in sections:
            if lab == label:
                add_section_rows(lab, df_sec, err)

    # Volcar a Excel
    compiled_df = pd.DataFrame(rows, columns=["", "Rubros", "Cuenta Contable", "Descripción", "Debe", "Haber"])
    compiled_df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)

# =============================
# Exportadores
# =============================
def build_excel_without_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes) -> BytesIO:
    """
    Excel rápido (sin copiar HT EF-4 original):
    - Original, Resultado General, Tipo1_con_1101/Tipo1_sin_1101 o Avisos
    - HT EF-4 (Compilada) hecha desde 'EF-1 Apertura' y 'EF-1 Final'
    """
    output = BytesIO()
    try:
        # Writer rápido
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Original
            df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
            df_original.to_excel(writer, index=False, sheet_name="Original")
            # Resultado
            df_result.to_excel(writer, index=False, sheet_name="Resultado General")
            # Particiones
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

            # === Nueva hoja compilada desde equivalencias ===
            write_ht_ef4_compilada(writer, equiv_bytes, sheet_name="HT EF-4 (Compilada)")

    except Exception:
        # Fallback a openpyxl
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

            # === Nueva hoja compilada desde equivalencias ===
            write_ht_ef4_compilada(writer, equiv_bytes, sheet_name="HT EF-4 (Compilada)")

    output.seek(0)
    return output

def build_excel_with_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes) -> BytesIO:
    """
    Excel con copia de hoja HT EF-4 (original, con estilos) y sumas por Rubro (G/H),
    más la hoja nueva HT EF-4 (Compilada) construída desde 'EF-1 Apertura' + 'EF-1 Final'.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Original
        df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # 2) Resultado General
        df_result.to_excel(writer, index=False, sheet_name="Resultado General")

        # 3) Particiones tipo_ctb
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
            df_tipo1_sin1101 = pd.DataFrame()
            pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                writer, index=False, sheet_name="Avisos"
            )

        # 4) Copiar hoja HT EF-4 original desde equivalencias y escribir sumas G/H
        book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
        book_result = writer.book

        if COPIABLE_SHEET in book_equiv.sheetnames:
            src_ws = book_equiv[COPIABLE_SHEET]
            dst_ws = book_result.create_sheet(COPIABLE_SHEET)
            copy_sheet_with_styles(src_ws, dst_ws)

            # Si tenemos df_tipo1_sin1101 y Rubros, sumar y escribir
            if 'df_tipo1_sin1101' in locals() and not df_tipo1_sin1101.empty and ("Rubros" in df_tipo1_sin1101.columns):
                df_sum = df_tipo1_sin1101.groupby("Rubros")[["debe_adj", "haber_adj"]].sum(numeric_only=True).reset_index()
                dict_debe = dict(zip(df_sum["Rubros"], df_sum["debe_adj"]))
                dict_haber = dict(zip(df_sum["Rubros"], df_sum["haber_adj"]))

                merged_ranges = dst_ws.merged_cells.ranges

                # Recorremos filas a partir de la 2 (asumiendo títulos en fila 1)
                # Rubro esperado en columna B (índice 2). Sumas en G (7) y H (8)
                for i, row in enumerate(dst_ws.iter_rows(min_row=2), start=2):
                    rubro_val = row[1].value  # columna B
                    rubro = str(rubro_val).strip() if rubro_val is not None else ""
                    if not rubro:
                        continue
                    debe_sum = float(dict_debe.get(rubro, 0.0))
                    haber_sum = float(dict_haber.get(rubro, 0.0))
                    if not is_inside_merged_area(i, 7, merged_ranges):
                        dst_ws.cell(row=i, column=7, value=debe_sum)
                    if not is_inside_merged_area(i, 8, merged_ranges):
                        dst_ws.cell(row=i, column=8, value=haber_sum)
        else:
            # Hoja de aviso si no existe la hoja copiable
            ws = writer.book.create_sheet("Aviso_HT_EF4")
            ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")

        # 5) === Nueva hoja compilada desde equivalencias ===
        write_ht_ef4_compilada(writer, equiv_bytes, sheet_name="HT EF-4 (Compilada)")

    output.seek(0)
    return output

# =============================
# UI
# =============================
opt_col1, opt_col2 = st.columns([1, 1])
with opt_col1:
    copy_ht = st.checkbox("Copiar hoja HT EF-4 (original) y llenar sumas (más pesado)", value=True, help="Si tu archivo es grande y se demora, desactívalo para generar un Excel rápido sin HT EF-4 original. La HT EF-4 (Compilada) se crea igual.")
with opt_col2:
    st.caption("Se copia la hoja 'HT EF-4' original (si está activado) y SIEMPRE se genera la hoja 'HT EF-4 (Compilada)' a partir de EF-1 Apertura y EF-1 Final.")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"], key="main")
with col2:
    equiv_file = st.file_uploader(f"Sube tu archivo de Equivalencias ({EQUIV_SHEET}, EF-1 Apertura, EF-1 Final)", type=["xlsx"], key="equiv")

if uploaded_file and equiv_file:
    try:
        main_bytes = _read_file_bytes(uploaded_file)
        equiv_bytes = _read_file_bytes(equiv_file)

        # Carga y procesamiento
        df_main = load_main_df(main_bytes)
        df_equiv = load_equiv_df(equiv_bytes)
        df_proc = compute_adjusted(df_main.copy())
        df_final = merge_equivalences(df_proc, df_equiv)

        # Excel según opción
        if copy_ht:
            xls_bytes = build_excel_with_ht(main_bytes, df_final, equiv_bytes)
            fname = "resultado_exp_contable_con_HT_EF4.xlsx"
        else:
            xls_bytes = build_excel_without_ht(main_bytes, df_final, equiv_bytes)
            fname = "resultado_exp_contable_simplificado.xlsx"

        st.success("Procesamiento completado.")
        st.download_button(
            label=f"Descargar {fname}",
            data=xls_bytes.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Ver vista previa (primeras 200 filas)"):
            st.dataframe(df_final.head(200))

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")

else:
    st.info("Sube ambos archivos para iniciar el procesamiento.")
