import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from copy import copy
import re
import unicodedata

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
    # Siempre se lee del SEGUNDO archivo (Equivalencias) la hoja "Hoja de Trabajo"
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
    # Merge con "Hoja de Trabajo" del SEGUNDO archivo (Equivalencias)
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

# --------- Normalizadores y utilidades para nombres y columnas ---------
def _norm_text(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _pick_col(df: pd.DataFrame, candidates: list[str], must_contain: str | None = None) -> str | None:
    cand_norm = {_norm_text(c) for c in candidates}
    for c in df.columns:
        if _norm_text(c) in cand_norm:
            return c
    if must_contain:
        mc = _norm_text(must_contain)
        for c in df.columns:
            if mc in _norm_text(c):
                return c
    return None

def _find_sheet_name(equiv_bytes: bytes, patterns: list[str]) -> str | None:
    wb = openpyxl.load_workbook(BytesIO(equiv_bytes), read_only=True, data_only=True)
    norm_map = {name: _norm_text(name) for name in wb.sheetnames}
    for name, n in norm_map.items():
        for pat in patterns:
            if re.search(pat, n):
                return name
    return None

# =============================
# (1) HT EF-4 (Compilada): EF-1 Apertura + EF-1 Final con columna Rubros añadida (desde "Hoja de Trabajo")
# =============================
def write_ht_ef4_compilada(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Compilada)"):
    # Map de equivalencias: Cuentas Contables -> Rubros
    map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))

    apertura_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*apert"])
    final_name    = _find_sheet_name(equiv_bytes, [r"ef\s*1.*final"])

    rows_offset = 0
    ws_df = None  # Pandas uses to_excel; we manage startrow.

    with pd.ExcelWriter(writer.book, engine="openpyxl") as _:
        # (We won't use this context; writer is already open. This dummy ensures no accidental close.)

        target_ws = writer.book.create_sheet(sheet_name)

    # Helper to dump a section
    def dump_section(label: str, sheet_real: str | None, start_row: int) -> int:
        # Title row
        target_ws = writer.book[sheet_name]
        target_ws.cell(row=start_row + 1, column=2, value=label)  # Col B
        row_ptr = start_row + 2

        if not sheet_real:
            target_ws.cell(row=row_ptr, column=2, value="(No se encontró la hoja)")
            return row_ptr + 2

        # Read original EF-1 sheet
        df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=sheet_real, dtype=str, engine="openpyxl")
        df_sec.columns = [str(c).strip() for c in df_sec.columns]

        # Identify key columns
        cta_col = _pick_col(df_sec, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")

        # Build output dataframe = original + Rubros (added at right)
        df_out = df_sec.copy()
        if cta_col:
            df_out["Rubros"] = df_out[cta_col].astype(str).map(map_cta_to_rubro).fillna("")
        else:
            df_out["Rubros"] = ""

        # Write headers
        for j, col in enumerate(df_out.columns, start=1):
            target_ws.cell(row=row_ptr, column=j, value=col)
        row_ptr += 1

        # Write data
        for _, row in df_out.iterrows():
            for j, col in enumerate(df_out.columns, start=1):
                target_ws.cell(row=row_ptr, column=j, value=str(row[col]) if pd.notna(row[col]) else "")
            row_ptr += 1

        return row_ptr + 1  # blank line

    # dump Apertura then Final
    target_ws = writer.book[sheet_name]
    rows_offset = dump_section("EF-1 Apertura", apertura_name, rows_offset)
    rows_offset = dump_section("EF-1 Final", final_name, rows_offset)

# =============================
# (2) HT EF-4 (Ctas): copia de HT EF-4 + lista de cuentas debajo de cada Rubro (col B)
# =============================
def write_ht_ef4_ctas(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Ctas)"):
    book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
    if COPIABLE_SHEET not in book_equiv.sheetnames:
        ws = writer.book.create_sheet(sheet_name)
        ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")
        return

    # 1) Copiar plantilla HT EF-4 a nueva hoja
    src_ws = book_equiv[COPIABLE_SHEET]
    dst_ws = writer.book.create_sheet(sheet_name)
    copy_sheet_with_styles(src_ws, dst_ws)

    # 2) Construir mapping Rubro -> set(Cuentas) desde EF-1 Apertura + EF-1 Final
    def accounts_by_rubro(equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> dict:
        map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))
        result = {}
        for pat, label in ([r"ef\s*1.*apert", "Apertura"], [r"ef\s*1.*final", "Final"]):
            sheet_name = _find_sheet_name(equiv_bytes, [pat])
            if not sheet_name:
                continue
            df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=sheet_name, dtype=str, engine="openpyxl")
            df_sec.columns = [str(c).strip() for c in df_sec.columns]
            cta_col = _pick_col(df_sec, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
            if not cta_col:
                continue
            # Map a rubro via equivalencias
            df_sec["__rubro__"] = df_sec[cta_col].astype(str).map(map_cta_to_rubro)
            for _, r in df_sec[["__rubro__", cta_col]].dropna().iterrows():
                rub = str(r["__rubro__"]).strip()
                cta = str(r[cta_col]).strip()
                if not rub or not cta:
                    continue
                result.setdefault(_norm_text(rub), set()).add(cta)
        return result

    rubro_to_accounts = accounts_by_rubro(equiv_bytes, df_equiv_ht)

    # 3) Rellenar debajo de cada Rubro (columna B) la lista de cuentas en columna C
    merged_ranges = dst_ws.merged_cells.ranges

    # Detectar filas donde hay Rubro: col B no vacía
    rubro_rows = []
    max_row = dst_ws.max_row
    for i in range(2, max_row + 1):  # asumir encabezados en fila 1
        val = dst_ws.cell(row=i, column=2).value  # Col B
        if val is not None and str(val).strip() != "":
            rubro_rows.append(i)

    # Para cada rubro, escribir las cuentas únicas dentro del bloque (hasta el siguiente rubro)
    for idx, start_row in enumerate(rubro_rows):
        rubro_val = dst_ws.cell(row=start_row, column=2).value
        rubro_key = _norm_text(str(rubro_val).strip())
        accounts = sorted(rubro_to_accounts.get(rubro_key, []))

        # Bloque disponible: desde start_row+1 hasta next_rubro_row-1 (o fin de hoja)
        end_row = rubro_rows[idx + 1] - 1 if idx + 1 < len(rubro_rows) else max_row
        write_row = start_row + 1

        if not accounts:
            continue

        # Escribir cuentas en Col C respetando celdas combinadas
        for k, cta in enumerate(accounts):
            if write_row > end_row:
                # no hay más espacio; indicar sobrantes
                dst_ws.cell(row=end_row, column=3, value=f"... (+{len(accounts) - k} más)")
                break
            if not is_inside_merged_area(write_row, 3, merged_ranges):
                dst_ws.cell(row=write_row, column=3, value=cta)
            write_row += 1

# =============================
# Exportadores
# =============================
def build_excel_without_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> BytesIO:
    """
    Excel sin copiar HT EF-4 original, pero:
    - Original, Resultado General, Tipo1_con_1101/Tipo1_sin_1101 o Avisos
    - HT EF-4 (Compilada) = EF-1 Apertura/Final con 'Rubros' añadidos
    - HT EF-4 (Ctas) = copia de HT EF-4 + cuentas debajo de cada Rubro
    (Usamos openpyxl para poder copiar la plantilla HT EF-4 con estilos)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Original
        df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Resultado General
        df_result.to_excel(writer, index=False, sheet_name="Resultado General")

        # Particiones tipo_ctb
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

        # Nueva hoja: HT EF-4 (Compilada)
        write_ht_ef4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")

        # Nueva hoja: HT EF-4 (Ctas)
        write_ht_ef4_ctas(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Ctas)")

    output.seek(0)
    return output

def build_excel_with_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> BytesIO:
    """
    Excel con:
    - Copia de HT EF-4 original (con estilos) y sumas por Rubro (G/H) si lo deseas en el futuro
    - HT EF-4 (Compilada) = EF-1 Apertura/Final con 'Rubros' añadidos
    - HT EF-4 (Ctas) = copia de HT EF-4 + cuentas debajo de cada Rubro
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

        # 4) Copiar hoja HT EF-4 original desde el 2º archivo (Equivalencias)
        book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
        book_result = writer.book

        if COPIABLE_SHEET in book_equiv.sheetnames:
            src_ws = book_equiv[COPIABLE_SHEET]
            dst_ws = book_result.create_sheet(COPIABLE_SHEET)
            copy_sheet_with_styles(src_ws, dst_ws)
        else:
            ws = writer.book.create_sheet("Aviso_HT_EF4")
            ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")

        # 5) Nueva hoja: HT EF-4 (Compilada)
        write_ht_ef4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")

        # 6) Nueva hoja: HT EF-4 (Ctas)
        write_ht_ef4_ctas(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Ctas)")

    output.seek(0)
    return output

# =============================
# UI
# =============================
opt_col1, opt_col2 = st.columns([1, 1])
with opt_col1:
    copy_ht = st.checkbox(
        "Incluir copia de HT EF-4 (original, con estilos)",
        value=True,
        help="Si no marcas esta opción, igual se crearán HT EF-4 (Compilada) y HT EF-4 (Ctas)."
    )
with opt_col2:
    st.caption("El 2º archivo (Equivalencias) debe contener 'Hoja de Trabajo' y, si es posible, 'EF-1 Apertura' y 'EF-1 Final'.")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"], key="main")
with col2:
    equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo, EF-1 Apertura, EF-1 Final)", type=["xlsx"], key="equiv")

if uploaded_file and equiv_file:
    try:
        main_bytes = _read_file_bytes(uploaded_file)
        equiv_bytes = _read_file_bytes(equiv_file)

        # Carga y procesamiento
        df_main = load_main_df(main_bytes)
        df_equiv_ht = load_equiv_df(equiv_bytes)  # Mapeo Rubros desde "Hoja de Trabajo"
        df_proc = compute_adjusted(df_main.copy())
        df_final = merge_equivalences(df_proc, df_equiv_ht)

        # Excel según opción
        if copy_ht:
            xls_bytes = build_excel_with_ht(main_bytes, df_final, equiv_bytes, df_equiv_ht)
            fname = "resultado_exp_contable_con_HT_EF4.xlsx"
        else:
            xls_bytes = build_excel_without_ht(main_bytes, df_final, equiv_bytes, df_equiv_ht)
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
