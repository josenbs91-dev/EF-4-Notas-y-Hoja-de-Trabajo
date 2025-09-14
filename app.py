import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from copy import copy
import re
import unicodedata

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
MISSING_RUBRO_LABEL = "(Sin Rubro)"
MISSING_EF4_LABEL = "(Sin EF-4)"
MISSING_ACT_LABEL = "(Sin Actividad)"

# =============================
# Utilidades / Helpers
# =============================
@st.cache_data(show_spinner=False)
def _read_file_bytes(uploaded_file) -> bytes:
    return uploaded_file.read()

@st.cache_data(show_spinner=False)
def load_main_df(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), dtype=str, engine="openpyxl")
    df.columns = [c.strip().lower() for c in df.columns]

    # Coerción numérica segura
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Validación de columnas requeridas
    missing = [c for c in REQUIRED_FOR_EXP if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas en el archivo principal: {', '.join(missing)}")

    # Construcción de exp_contable
    parts = [df[c].astype(str).fillna("") for c in REQUIRED_FOR_EXP]
    df["exp_contable"] = parts[0] + "-" + parts[1] + "-" + parts[2] + "-" + parts[3]

    # Clave contable
    mayor = df["mayor"].astype(str) if "mayor" in df.columns else pd.Series("", index=df.index)
    sub_cta = df["sub_cta"].astype(str) if "sub_cta" in df.columns else pd.Series("", index=df.index)
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

    # No eliminamos columnas extra (necesitamos EF-4 y Actividad si existen)
    df_e["Cuentas Contables"] = df_e["Cuentas Contables"].astype(str).str.strip()
    df_e["Rubros"] = df_e["Rubros"].astype(str).str.strip()
    df_e = df_e.drop_duplicates(subset=["Cuentas Contables"], keep="first").reset_index(drop=True)
    return df_e

def compute_adjusted(df: pd.DataFrame) -> pd.DataFrame:
    """Ajusta debe/haber si el exp_contable pertenece a un mayor 1101."""
    mayor_series = df["mayor"].astype(str) if "mayor" in df.columns else pd.Series("", index=df.index)
    mask_mayor_1101 = mayor_series.eq("1101")
    exp_con_1101 = set(df.loc[mask_mayor_1101, "exp_contable"].unique())
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

# ---------- Normalizadores / búsqueda tolerante ----------
def _norm_text(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _norm_account_code(s: str) -> str:
    s = str(s or "")
    s = s.strip()
    s = re.sub(r"\s+", "", s)        # quita espacios
    s = re.sub(r"[^0-9.]", "", s)    # solo dígitos y puntos
    s = re.sub(r"\.+", ".", s)       # colapsa .. en .
    return s.strip(".")

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

def _find_structure_sheet_name(equiv_bytes: bytes) -> str | None:
    """Encuentra la hoja de estructura de Rubros de forma robusta."""
    wb = openpyxl.load_workbook(BytesIO(equiv_bytes), read_only=True, data_only=True)
    exact_norm = {"estructura del archivo", "estructura_del_archivo"}
    candidates = []
    for name in wb.sheetnames:
        n = _norm_text(name)
        if n in exact_norm:
            return name
        if "estruct" in n:
            candidates.append(name)
    # Validar candidatos por columnas compatibles
    for name in candidates:
        try:
            df_test = pd.read_excel(BytesIO(equiv_bytes), sheet_name=name, dtype=str, engine="openpyxl", nrows=5)
            cols_norm = [_norm_text(c) for c in df_test.columns]
            if any(
                cn in {"estructura", "descripcion", "descripción", "rubros", "rubro"} or
                ("estruct" in cn) or ("descr" in cn)
                for cn in cols_norm
            ):
                return name
        except Exception:
            pass
    return None

def _find_estructura_ef4_sheet_name(equiv_bytes: bytes) -> str | None:
    """Encuentra la hoja 'Estructura EF-4' (o similar) para ordenar EF-4 y Actividades."""
    wb = openpyxl.load_workbook(BytesIO(equiv_bytes), read_only=True, data_only=True)
    for name in wb.sheetnames:
        n = _norm_text(name)
        if "estruct" in n and ("ef4" in n or "ef 4" in n or "ef-4" in n or "ef_4" in n):
            return name
    # fallback: buscar un nombre que contenga "ef-4"
    for name in wb.sheetnames:
        n = _norm_text(name)
        if "ef-4" in n or "ef4" in n or "ef 4" in n or "ef_4" in n:
            return name
    return None

# =============================
# (1) HT EF-4 (Compilada): EF-1 Apertura + EF-1 Final con columna "Rubros"
# =============================
def write_ht_ef_4_compilada(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Compilada)"):
    # Map Cuentas -> Rubros (desde "Hoja de Trabajo")
    map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))

    ap_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*apert"])
    fi_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*final"])

    ws = writer.book.create_sheet(sheet_name)
    row_ptr = 1

    def dump_section(label: str, real_name: str | None):
        nonlocal row_ptr, ws
        ws.cell(row=row_ptr, column=2, value="Sección:")
        ws.cell(row=row_ptr, column=3, value=label)
        row_ptr += 1

        if not real_name:
            ws.cell(row=row_ptr, column=2, value="(No se encontró la hoja solicitada)")
            ws.cell(row=row_ptr+1, column=2, value="(Revise que el nombre coincida con 'EF-1 Apertura' o 'EF-1 Final')")
            row_ptr += 3
            return

        df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=real_name, dtype=str, engine="openpyxl")
        df_sec.columns = [str(c).strip() for c in df_sec.columns]
        cta_col = _pick_col(df_sec, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
        df_out = df_sec.copy()
        if cta_col:
            df_out["Rubros"] = df_out[cta_col].astype(str).map(map_cta_to_rubro).fillna("")
        else:
            df_out["Rubros"] = ""

        # Escribir encabezados
        for j, col in enumerate(df_out.columns, start=1):
            ws.cell(row=row_ptr, column=j, value=col)
        row_ptr_local = row_ptr + 1

        # Escribir datos
        for _, r in df_out.iterrows():
            for j, col in enumerate(df_out.columns, start=1):
                ws.cell(row=row_ptr_local, column=j, value=str(r[col]) if pd.notna(r[col]) else "")
            row_ptr_local += 1

        row_ptr = row_ptr_local + 1  # línea en blanco

    dump_section("EF-1 Apertura", ap_name)
    dump_section("EF-1 Final", fi_name)

# =============================
# EF-2: Variaciones POR CUENTA (desde hoja EF-2 Final) -> para integrarlas en filas de cuenta
# =============================
def _compute_ef2_variaciones_por_cuenta(equiv_bytes: bytes, df_equiv_ht: pd.DataFrame):
    """
    Lee 'EF-2 Final':
      - Cuenta_Nombre -> primer token (antes del 1er espacio), p.ej.: '4104.010104 Arancel...' -> '4104.010104'
      - Normaliza código de cuenta (solo dígitos y puntos)
      - Agrupa importes por CUENTA normalizada (abs)
      - Prefijo '5' -> Variación +, prefijo '4' -> Variación -
    Retorna dicts por cuenta normalizada.
    """
    sheet = _find_sheet_name(equiv_bytes, [r"ef\s*2.*final"])
    if not sheet:
        return {}, {}

    df = pd.read_excel(BytesIO(equiv_bytes), sheet_name=sheet, dtype=str, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    cuenta_nombre_col = _pick_col(df, ["Cuenta_Nombre", "Cuenta Nombre", "CUENTA_NOMBRE", "Cuenta", "Descripción"], must_contain="cuenta")
    importe_col = _pick_col(df, ["Importe", "Importes", "Monto", "Valor", "Importe S/.", "Importe S"], must_contain="import")
    if cuenta_nombre_col is None or importe_col is None:
        return {}, {}

    cuentas_raw = df[cuenta_nombre_col].astype(str).str.strip().str.split().str[0]  # primer token
    cuentas = cuentas_raw.fillna("").map(_norm_account_code)
    importes = pd.to_numeric(df[importe_col], errors="coerce").fillna(0.0).abs()

    tmp = pd.DataFrame({"Cuenta": cuentas, "Importe": importes})
    tmp = tmp[tmp["Cuenta"] != ""].copy()
    by_acc = tmp.groupby("Cuenta", as_index=False)["Importe"].sum()

    # Clasificación por prefijo
    by_acc["pref"] = by_acc["Cuenta"].str[:1]
    plus_map = dict(zip(by_acc.loc[by_acc["pref"] == "5", "Cuenta"], by_acc.loc[by_acc["pref"] == "5", "Importe"]))
    minus_map = dict(zip(by_acc.loc[by_acc["pref"] == "4", "Cuenta"], by_acc.loc[by_acc["pref"] == "4", "Importe"]))

    return plus_map, minus_map

# =============================
# (2) HT EF-4 (Estructura) + Variaciones, Debe/Haber (HT) a nivel cuenta y Saldos Ajustados
# (integra EF-2 en filas de cuenta) + Hoja de Auditoría + Totales y Diferencia
# =============================
def write_ht_ef4_estructura(
    writer,
    equiv_bytes: bytes,
    df_equiv_ht: pd.DataFrame,
    sheet_name: str = "HT EF-4 (Estructura)",
    acc_debe_map: dict | None = None,
    acc_haber_map: dict | None = None,
    rub_debe_map: dict | None = None,
    rub_haber_map: dict | None = None,
    ef2_acc_plus_map: dict | None = None,
    ef2_acc_minus_map: dict | None = None,
):
    acc_debe_map = acc_debe_map or {}
    acc_haber_map = acc_haber_map or {}
    rub_debe_map = rub_debe_map or {}
    rub_haber_map = rub_haber_map or {}
    ef2_acc_plus_map = ef2_acc_plus_map or {}
    ef2_acc_minus_map = ef2_acc_minus_map or {}

    # Mapeo de Cuentas -> Rubros desde "Hoja de Trabajo" (con fallback normalizado)
    map_cta_to_rubro_raw = dict(zip(df_equiv_ht["Cuentas Contables"].astype(str).str.strip(),
                                    df_equiv_ht["Rubros"].astype(str).str.strip()))
    map_cta_to_rubro_norm = { _norm_account_code(k): v for k, v in map_cta_to_rubro_raw.items() }
    def _get_rubro_for_account(cta_code: str) -> str:
        return map_cta_to_rubro_raw.get(cta_code) or map_cta_to_rubro_norm.get(_norm_account_code(cta_code), MISSING_RUBRO_LABEL)

    # ---- Mapas (EXCLUSIVOS Tipo1_sin_1101) ----
    acc_debe_map_norm = { _norm_account_code(k): v for k, v in (acc_debe_map or {}).items() }
    acc_haber_map_norm = { _norm_account_code(k): v for k, v in (acc_haber_map or {}).items() }
    rub_debe_map_norm = { _norm_text(k): v for k, v in (rub_debe_map or {}).items() }
    rub_haber_map_norm = { _norm_text(k): v for k, v in (rub_haber_map or {}).items() }

    # Localizar hojas EF-1
    ap_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*apert"])
    fi_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*final"])

    # --- Lectura de importes por cuenta (Apertura / Final) NORMALIZANDO cuentas ---
    def read_importes(sheet_name: str) -> pd.DataFrame:
        if not sheet_name:
            return pd.DataFrame(columns=["Cuenta", "Importe"])
        df = pd.read_excel(BytesIO(equiv_bytes), sheet_name=sheet_name, dtype=str, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        cta_col = _pick_col(df, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
        imp_col = _pick_col(df, ["Importes", "Importe", "Monto", "Valor", "Importe S/.", "Importe S"], must_contain="import")
        if not cta_col:
            return pd.DataFrame(columns=["Cuenta", "Importe"])
        cuentas = df[cta_col].astype(str).map(_norm_account_code)
        vals = pd.to_numeric(df[imp_col], errors="coerce").fillna(0.0) if imp_col else 0.0
        out = pd.DataFrame({"Cuenta": cuentas, "Importe": vals})
        out = out[out["Cuenta"] != ""]
        out = out.groupby("Cuenta", as_index=False)["Importe"].sum()
        return out

    ap_df = read_importes(ap_name)
    fi_df = read_importes(fi_name)

    # EF-2 POR CUENTA
    ef2_acc_plus_map, ef2_acc_minus_map = (ef2_acc_plus_map or {}), (ef2_acc_minus_map or {})

    # --- Conjunto de CUENTAS a mostrar (incluye EF-2 y TODAS las del principal Tipo1_sin_1101) ---
    cuentas_ef1 = set(ap_df["Cuenta"]).union(set(fi_df["Cuenta"]))
    cuentas_ef2 = set(ef2_acc_plus_map.keys()).union(set(ef2_acc_minus_map.keys()))
    cuentas_main_norm = set(acc_debe_map_norm.keys()).union(set(acc_haber_map_norm.keys()))
    cuentas = sorted(cuentas_ef1.union(cuentas_ef2).union(cuentas_main_norm))

    # --- Construcción por cuenta + auditoría ---
    audit_rows = []
    rows_data = []
    for cta in cuentas:
        rub = _get_rubro_for_account(cta)
        ap_val = float(ap_df.loc[ap_df["Cuenta"] == cta, "Importe"].sum()) if not ap_df.empty else 0.0
        fi_val = float(fi_df.loc[fi_df["Cuenta"] == cta, "Importe"].sum()) if not fi_df.empty else 0.0

        # Variaciones por prefijo de cuenta (modelo EF-1)
        diff = fi_val - ap_val
        starts = str(cta).strip()[:1]
        if starts == "1":
            var_plus = max(diff, 0.0)
            var_minus = abs(min(diff, 0.0))
        elif starts in {"2", "3"}:
            var_plus = abs(min(diff, 0.0))
            var_minus = max(diff, 0.0)
        else:
            var_plus = 0.0
            var_minus = 0.0

        # SUMA EF-2 a NIVEL CUENTA
        ef2_plus = float(ef2_acc_plus_map.get(cta, 0.0))
        ef2_minus = float(ef2_acc_minus_map.get(cta, 0.0))
        var_plus += ef2_plus
        var_minus += ef2_minus

        # Debe/Haber (HT) por cuenta -> EXCLUSIVAMENTE Tipo1_sin_1101
        debe_ht = float(acc_debe_map_norm.get(_norm_account_code(cta), 0.0))
        haber_ht = float(acc_haber_map_norm.get(_norm_account_code(cta), 0.0))

        # === Saldos Ajustados NUEVA DEFINICIÓN ===
        saldo_aj = float(var_plus + debe_ht - var_minus - haber_ht)

        rows_data.append({
            "Rubros": rub,
            "Rubros_norm": _norm_text(rub),
            "Cuenta Contable": cta,
            "EF-1 Final": fi_val,
            "EF-1 Apertura": ap_val,
            "Variación +": var_plus,
            "Variación -": var_minus,
            "Debe (HT EF-4)": debe_ht,
            "Haber (HT EF-4)": haber_ht,
            "Saldos Ajustados": saldo_aj
        })

        if rub == MISSING_RUBRO_LABEL:
            audit_rows.append({
                "Cuenta Contable": cta,
                "EF-1 Final": fi_val,
                "EF-1 Apertura": ap_val,
                "EF-2 Variación +": ef2_plus,
                "EF-2 Variación -": ef2_minus,
                "Observación": "Cuenta sin Rubro en Hoja de Trabajo"
            })

    df_all = pd.DataFrame(rows_data)
    if df_all.empty:
        pd.DataFrame({"info": ["No se pudieron consolidar Importes de EF-1/EF-2."]}).to_excel(
            writer, index=False, sheet_name=sheet_name
        )
        return df_all  # <-- devolver igualmente

    # === Orden estricto desde 'Estructura del archivo' (Rubros) ===
    def struct_order_strict() -> list | None:
        struct_name = _find_structure_sheet_name(equiv_bytes)
        if not struct_name:
            return None
        try:
            df_struct = pd.read_excel(BytesIO(equiv_bytes), sheet_name=struct_name, dtype=str, engine="openpyxl")
            df_struct.columns = [str(c).strip() for c in df_struct.columns]
            rub_col = _pick_col(df_struct, ["Rubros", "Rubro", "Estructura", "DESCRIPCION", "Descripción", "Descripcion"])
            if rub_col is None:
                rub_col = _pick_col(df_struct, [], must_contain="estruct") or _pick_col(df_struct, [], must_contain="descr")
            if rub_col is None:
                return None
            ord_col = _pick_col(df_struct, ["Orden", "Order", "Ordenamiento", "N°", "No", "Nro", "Nro."], must_contain="orden")

            if ord_col:
                tmp = df_struct[[ord_col, rub_col]].copy()
                tmp[ord_col] = pd.to_numeric(tmp[ord_col], errors="coerce")
                tmp["_row"] = np.arange(len(tmp))
                tmp = tmp.sort_values([ord_col, "_row"], na_position="last", kind="mergesort")
                ordered = [str(x).strip() for x in tmp[rub_col].tolist()]
            else:
                ordered = [str(x).strip() for x in df_struct[rub_col].tolist()]
            seen, final = set(), []
            for r in ordered:
                if r and r not in seen:
                    seen.add(r)
                    final.append(r)
            return final
        except Exception:
            return None

    strict_order = struct_order_strict()
    if strict_order and len([x for x in strict_order if str(x).strip()]) > 0:
        rubros_order = [str(r).strip() for r in strict_order if str(r).strip() != ""]
    else:
        rubros_from_equiv = [str(x).strip() for x in df_equiv_ht["Rubros"].astype(str).tolist() if str(x).strip() != ""]
        if rubros_from_equiv:
            seen, rubros_order = set(), []
            for r in rubros_from_equiv:
                if r not in seen:
                    seen.add(r)
                    rubros_order.append(r)
        else:
            rubros_presentes = [str(x).strip() for x in df_all["Rubros"].astype(str).unique() if str(x).strip() != ""]
            rubros_order = sorted(rubros_presentes) if rubros_presentes else []

    if MISSING_RUBRO_LABEL in set(df_all["Rubros"].astype(str)):
        if MISSING_RUBRO_LABEL not in rubros_order:
            rubros_order.append(MISSING_RUBRO_LABEL)

    # Totales por rubro (para pintar fila rubro)
    df_all["_rub_norm"] = df_all["Rubros"].astype(str).map(_norm_text)
    totals_norm = (
        df_all.groupby("_rub_norm")[["EF-1 Final", "EF-1 Apertura", "Variación +", "Variación -", "Saldos Ajustados"]]
        .sum(numeric_only=True)
        .to_dict()
    )

    # --- Construcción HT EF-4 (Estructura) ---
    header = [
        "", "Rubro", "Cuenta Contable",
        "EF-1 Final", "EF-1 Apertura", "Variación +", "Variación -",
        "Debe (HT EF-4)", "Haber (HT EF-4)", "Saldos Ajustados"
    ]
    out_rows = [header]
    for rub in rubros_order:
        rub_norm = _norm_text(rub)
        out_rows.append(["", rub, "", "", "", "", "", "", "", ""])  # fila Rubro

        block = df_all[df_all["_rub_norm"] == rub_norm].copy()
        block = block.sort_values(["Cuenta Contable"]).drop_duplicates(subset=["Cuenta Contable"], keep="first")
        if block.empty:
            out_rows.append(["", "", "(sin cuentas)", 0.0, 0.0, 0.0, 0.0,
                             0.0, 0.0, 0.0])
        else:
            for _, r in block.iterrows():
                out_rows.append([
                    "", "",
                    r["Cuenta Contable"],
                    float(r["EF-1 Final"]),
                    float(r["EF-1 Apertura"]),
                    float(r["Variación +"]),
                    float(r["Variación -"]),
                    float(r["Debe (HT EF-4)"]),
                    float(r["Haber (HT EF-4)"]),
                    float(r["Saldos Ajustados"]),
                ])
        out_rows.append(["", "", "", "", "", "", "", "", "", ""])  # línea en blanco

    out_df = pd.DataFrame(out_rows[1:], columns=out_rows[0])
    out_df.to_excel(writer, index=False, sheet_name=sheet_name)

    # --------- FORMATO + Totales por Rubro ---------
    ws = writer.book[sheet_name]
    max_row = ws.max_row

    widths = {2: 42, 3: 22, 4: 18, 5: 18, 6: 16, 7: 16, 8: 16, 9: 16, 10: 18}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFEFEFEF")
    center = Alignment(horizontal="center", vertical="center")
    for c in range(2, 11):  # B..J
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    num_align = Alignment(horizontal="right")
    for r in range(2, max_row + 1):
        for c in [4, 5, 6, 7, 8, 9, 10]:
            ws.cell(row=r, column=c).number_format = '#,##0.00'
            ws.cell(row=r, column=c).alignment = num_align

    rubro_fill = PatternFill("solid", fgColor="FFF7F7F7")
    rubro_font = Font(bold=True)
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r in range(2, max_row + 1):
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        if (b is not None and str(b).strip() != "") and (c is None or str(c).strip() == ""):
            rub = str(b).strip()
            rub_norm = _norm_text(rub)
            for col in range(2, 11):
                ws.cell(row=r, column=col).fill = rubro_fill
            ws.cell(row=r, column=2).font = rubro_font
            ws.cell(row=r, column=4, value=float(totals_norm.get("EF-1 Final", {}).get(rub_norm, 0.0))).font = rubro_font
            ws.cell(row=r, column=5, value=float(totals_norm.get("EF-1 Apertura", {}).get(rub_norm, 0.0))).font = rubro_font
            ws.cell(row=r, column=6, value=float(totals_norm.get("Variación +", {}).get(rub_norm, 0.0))).font = rubro_font
            ws.cell(row=r, column=7, value=float(totals_norm.get("Variación -", {}).get(rub_norm, 0.0))).font = rubro_font
            # Para vista, no reescribo Debe/Haber por rubro aquí (ya vienen por cuenta)
            ws.cell(row=r, column=10, value=float(totals_norm.get("Saldos Ajustados", {}).get(rub_norm, 0.0))).font = rubro_font

    for r in range(1, max_row + 1):
        for c in range(2, 11):
            ws.cell(row=r, column=c).border = border

    # --------- TOTALES GENERALES (al pie) ---------
    cols_sum = ["EF-1 Final","EF-1 Apertura","Variación +","Variación -","Debe (HT EF-4)","Haber (HT EF-4)","Saldos Ajustados"]
    tot_cuentas = {col: float(df_all[col].sum()) for col in cols_sum}

    row_total_cta = ws.max_row + 1
    total_fill = PatternFill("solid", fgColor="FFE9F5FF")
    total_font = Font(bold=True)
    for col_idx, key in zip([4,5,6,7,8,9,10], cols_sum):
        ws.cell(row=row_total_cta, column=col_idx, value=tot_cuentas[key]).font = total_font
        ws.cell(row=row_total_cta, column=col_idx).number_format = '#,##0.00'
    ws.cell(row=row_total_cta, column=2, value="TOTAL CUENTAS").font = total_font
    for c in range(2, 11):
        ws.cell(row=row_total_cta, column=c).fill = total_fill
        ws.cell(row=row_total_cta, column=c).border = border

    ws.auto_filter.ref = f"B1:J{ws.max_row}"
    ws.freeze_panes = "B2"

    # --------- Hoja de Auditoría: cuentas sin rubro ---------
    if audit_rows:
        df_aud = pd.DataFrame(audit_rows)
        agg_cols = ["EF-1 Final", "EF-1 Apertura", "EF-2 Variación +", "EF-2 Variación -"]
        df_aud = df_aud.groupby("Cuenta Contable", as_index=False)[agg_cols].sum(numeric_only=True)
        df_aud["Observación"] = "Cuenta sin Rubro en Hoja de Trabajo"
        df_aud.to_excel(writer, index=False, sheet_name="Auditoría (Sin Rubro)")

    # === DEVOLVER df_all para usarlo en 'Estructura EF-4 (Detalle)' ===
    return df_all

# =============================
# Nueva hoja: Estructura EF-4 (Detalle)
# =============================
def write_estructura_ef4_detalle(
    writer,
    equiv_bytes: bytes,
    df_equiv_ht: pd.DataFrame,
    df_ht_estructura_all: pd.DataFrame,
    sheet_name: str = "Estructura EF-4 (Detalle)",
):
    # Validar columnas EF-4 y Actividad en Hoja de Trabajo
    ef4_col = _pick_col(df_equiv_ht, ["EF-4", "EF4", "EF_4"], must_contain="ef")
    act_col = _pick_col(df_equiv_ht, ["Actividad", "Actividad EF-4", "Actividad_EF4", "Actividad EF4"], must_contain="activ")
    cta_col = "Cuentas Contables"

    if ef4_col is None or act_col is None or cta_col not in df_equiv_ht.columns:
        ws = writer.book.create_sheet("Aviso_Estructura_EF4")
        ws.cell(row=1, column=1, value="No se encontraron columnas 'EF-4' y/o 'Actividad' en 'Hoja de Trabajo'.")
        return

    # Mapear cuenta -> EF-4 / Actividad
    df_map = df_equiv_ht[[cta_col, ef4_col, act_col]].copy()
    df_map[cta_col] = df_map[cta_col].astype(str).map(_norm_account_code)
    df_map[ef4_col] = df_map[ef4_col].astype(str).str.strip().replace({"": np.nan}).fillna(MISSING_EF4_LABEL)
    df_map[act_col] = df_map[act_col].astype(str).str.strip().replace({"": np.nan}).fillna(MISSING_ACT_LABEL)

    map_cta_to_ef4 = dict(zip(df_map[cta_col], df_map[ef4_col]))
    map_cta_to_act = dict(zip(df_map[cta_col], df_map[act_col]))

    # Base: tomar Cuenta Contable + Saldos Ajustados desde HT EF-4 (Estructura)
    if df_ht_estructura_all is None or df_ht_estructura_all.empty:
        ws = writer.book.create_sheet("Aviso_Estructura_EF4")
        ws.cell(row=1, column=1, value="No se pudo leer el detalle de 'HT EF-4 (Estructura)'.")
        return

    base = df_ht_estructura_all.copy()
    base = base[["Cuenta Contable", "Saldos Ajustados"]].copy()
    base = base[base["Cuenta Contable"].astype(str).str.strip() != ""]
    base["Cuenta_norm"] = base["Cuenta Contable"].astype(str).map(_norm_account_code)

    # Adjuntar EF-4 / Actividad
    base["EF-4"] = base["Cuenta_norm"].map(map_cta_to_ef4).fillna(MISSING_EF4_LABEL)
    base["Actividad"] = base["Cuenta_norm"].map(map_cta_to_act).fillna(MISSING_ACT_LABEL)

    # Agrupar por cuenta por si está repetida
    base = (base.groupby(["EF-4", "Actividad", "Cuenta Contable"], as_index=False)["Saldos Ajustados"]
                 .sum(numeric_only=True))

    # ORDEN: usar hoja "Estructura EF-4" si existe
    ef4_struct_name = _find_estructura_ef4_sheet_name(equiv_bytes)

    def order_from_struct():
        if not ef4_struct_name:
            return None, None
        try:
            dfs = pd.read_excel(BytesIO(equiv_bytes), sheet_name=ef4_struct_name, dtype=str, engine="openpyxl")
            dfs.columns = [str(c).strip() for c in dfs.columns]
            ef4c = _pick_col(dfs, ["EF-4", "EF4", "EF_4"], must_contain="ef")
            actc = _pick_col(dfs, ["Actividad", "Actividad EF-4", "Actividad_EF4", "Actividad EF4"], must_contain="activ")
            ordc = _pick_col(dfs, ["Orden", "Order", "No", "Nro", "N°"], must_contain="orden")
            if ef4c is None:
                return None, None
            tmp = dfs.copy()
            tmp["_row"] = np.arange(len(tmp))
            if ordc:
                tmp["__ord"] = pd.to_numeric(tmp[ordc], errors="coerce")
            else:
                tmp["__ord"] = np.nan
            # Orden EF-4
            order_ef4 = []
            seen = set()
            tmp_ef4 = tmp[[ef4c, "__ord", "_row"]].dropna(subset=[ef4c])
            tmp_ef4 = tmp_ef4.sort_values(["__ord", "_row"], na_position="last", kind="mergesort")
            for v in tmp_ef4[ef4c].astype(str):
                if v not in seen:
                    seen.add(v)
                    order_ef4.append(v)
            # Orden Actividad por EF-4
            order_act = {}
            if actc:
                for ef in order_ef4:
                    blk = tmp[tmp[ef4c].astype(str) == ef]
                    seen2 = set()
                    acts = []
                    blk2 = blk[[actc, "__ord", "_row"]].dropna(subset=[actc]).sort_values(
                        ["__ord", "_row"], na_position="last", kind="mergesort"
                    )
                    for a in blk2[actc].astype(str):
                        if a not in seen2:
                            seen2.add(a)
                            acts.append(a)
                    order_act[ef] = acts
            return order_ef4, order_act
        except Exception:
            return None, None

    order_ef4, order_act_map = order_from_struct()

    # Si no hay estructura, usar orden de aparición en Hoja de Trabajo
    if not order_ef4:
        tmp_ht = df_map[[ef4_col, act_col]].copy()
        tmp_ht["_row"] = np.arange(len(tmp_ht))
        order_ef4 = []
        seen = set()
        for v in tmp_ht[ef4_col].astype(str):
            if v not in seen and v.strip() != "":
                seen.add(v)
                order_ef4.append(v)
        order_act_map = {}
        for ef in order_ef4:
            blk = tmp_ht[tmp_ht[ef4_col].astype(str) == ef]
            seen2, acts = set(), []
            for a in blk[act_col].astype(str):
                if a not in seen2 and a.strip() != "":
                    seen2.add(a)
                    acts.append(a)
            order_act_map[ef] = acts

    # Último recurso: usar orden alfabético
    if not order_ef4:
        order_ef4 = sorted(base["EF-4"].unique())
    for ef in order_ef4:
        if ef not in order_act_map:
            order_act_map[ef] = sorted(base.loc[base["EF-4"] == ef, "Actividad"].unique())

    # Construir salida
    ws = writer.book.create_sheet(sheet_name)
    headers = ["EF-4", "Actividad", "Cuenta Contable", "Saldos Ajustados"]
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=j, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFEFEFEF")
        cell.alignment = Alignment(horizontal="center")

    row_ptr = 2
    # bordes y formatos
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = '#,##0.00'

    def write_total_row(r, label_a, label_b, amount, fill_color):
        ws.cell(row=r, column=1, value=label_a).font = Font(bold=True)
        ws.cell(row=r, column=2, value=label_b).font = Font(bold=True)
        ws.cell(row=r, column=4, value=float(amount)).font = Font(bold=True)
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=fill_color)
            ws.cell(row=r, column=c).border = border
            if c == 4:
                ws.cell(row=r, column=c).number_format = num_fmt
                ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")

    total_general = 0.0

    for ef in order_ef4:
        df_ef = base[base["EF-4"] == ef].copy()
        if df_ef.empty:
            continue

        # Totales por EF-4 se suman al final del bloque
        ef_total = 0.0

        acts = order_act_map.get(ef, sorted(df_ef["Actividad"].unique()))
        for act in acts:
            df_act = df_ef[df_ef["Actividad"] == act].copy()
            if df_act.empty:
                continue

            # Total por actividad
            act_total = float(df_act["Saldos Ajustados"].sum())
            write_total_row(row_ptr, ef, act, act_total, "FFF7F7F7")
            row_ptr += 1

            # Detalle por cuenta
            df_act = df_act.sort_values(["Cuenta Contable"]).drop_duplicates(subset=["Cuenta Contable"], keep="first")
            for _, r in df_act.iterrows():
                ws.cell(row=row_ptr, column=1, value="")  # EF-4 vacío en detalle
                ws.cell(row=row_ptr, column=2, value="")  # Actividad vacía en detalle
                ws.cell(row=row_ptr, column=3, value=str(r["Cuenta Contable"]))
                ws.cell(row=row_ptr, column=4, value=float(r["Saldos Ajustados"]))
                ws.cell(row=row_ptr, column=4).number_format = num_fmt
                ws.cell(row=row_ptr, column=4).alignment = Alignment(horizontal="right")
                for c in range(1, 5):
                    ws.cell(row=row_ptr, column=c).border = border
                row_ptr += 1

            ef_total += act_total
            # línea en blanco entre actividades
            row_ptr += 1

        # Total por EF-4
        write_total_row(row_ptr, f"TOTAL {ef}", "", ef_total, "FFE9F5FF")
        row_ptr += 2

        total_general += ef_total

    # Total general
    write_total_row(row_ptr, "TOTAL GENERAL", "", total_general, "FFFFF2CC")

    # Anchos
    widths = {1: 38, 2: 42, 3: 22, 4: 18}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

# =============================
# Exportadores (SIEMPRE openpyxl)
# =============================
def _compute_maps_para_estructura(df_result: pd.DataFrame):
    """
    Devuelve mapas EXCLUSIVOS de Tipo1_sin_1101:
      - acc_debe_map / acc_haber_map: por CUENTA (clave_cta), sumas de debe_adj/haber_adj
      - rub_debe_map / rub_haber_map: por RUBRO, sumas de debe_adj/haber_adj
    Criterio:
      * tipo_ctb == "1"
      * Excluir expedientes cuyo 'mayor' == "1101" (se excluye TODO el exp_contable)
    """
    acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map = {}, {}, {}, {}
    if "tipo_ctb" not in df_result.columns:
        return acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map

    df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()

    mayor_series = df_result["mayor"].astype(str) if "mayor" in df_result.columns else pd.Series("", index=df_result.index)
    exp_con_1101 = set(df_result.loc[mayor_series.eq("1101"), "exp_contable"].unique())
    in_1101 = df_tipo1["exp_contable"].isin(exp_con_1101)
    df_tipo1_sin1101 = df_tipo1[~in_1101].copy()

    if df_tipo1_sin1101.empty:
        return acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map

    for col in ["debe_adj", "haber_adj"]:
        if col in df_tipo1_sin1101.columns:
            df_tipo1_sin1101[col] = pd.to_numeric(df_tipo1_sin1101[col], errors="coerce").fillna(0.0)

    # Por cuenta (clave_cta)
    acc = df_tipo1_sin1101.groupby("clave_cta")[["debe_adj", "haber_adj"]].sum(numeric_only=True).reset_index()
    acc_debe_map = dict(zip(acc["clave_cta"], acc["debe_adj"]))
    acc_haber_map = dict(zip(acc["clave_cta"], acc["haber_adj"]))

    # Por rubro (relleno de vacíos a "(Sin Rubro)")
    if "Rubros" in df_tipo1_sin1101.columns:
        tmp = df_tipo1_sin1101.copy()
        tmp["Rubros"] = tmp["Rubros"].astype(str)
        tmp["Rubros"] = tmp["Rubros"].replace({"": np.nan}).fillna(MISSING_RUBRO_LABEL)
        rub = tmp.groupby("Rubros")[["debe_adj", "haber_adj"]].sum(numeric_only=True).reset_index()
        rub_debe_map = dict(zip(rub["Rubros"], rub["debe_adj"]))
        rub_haber_map = dict(zip(rub["Rubros"], rub["haber_adj"]))

    return acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map

def build_excel_without_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Original
        df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # Resultado General
        df_result.to_excel(writer, index=False, sheet_name="Resultado General")

        # Particiones tipo_ctb (informativas)
        if "tipo_ctb" in df_result.columns:
            mayor_series = df_result["mayor"].astype(str) if "mayor" in df_result.columns else pd.Series("", index=df_result.index)
            exp_con_1101 = set(df_result.loc[mayor_series.eq("1101"), "exp_contable"].unique())

            df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()
            in_1101 = df_tipo1["exp_contable"].isin(exp_con_1101)
            df_tipo1_con1101 = df_tipo1[in_1101].copy()
            df_tipo1_sin1101 = df_tipo1[~in_1101].copy()
            df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
            df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
        else:
            pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                writer, index=False, sheet_name="Avisos"
            )

        # Hojas nuevas
        write_ht_ef_4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")

        # Mapas Debe/Haber para Estructura (EXCLUSIVOS Tipo1_sin_1101)
        acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map = _compute_maps_para_estructura(df_result)

        # EF-2: variaciones POR CUENTA
        ef2_acc_plus_map, ef2_acc_minus_map = _compute_ef2_variaciones_por_cuenta(equiv_bytes, df_equiv_ht)

        df_all = write_ht_ef4_estructura(
            writer, equiv_bytes, df_equiv_ht,
            sheet_name="HT EF-4 (Estructura)",
            acc_debe_map=acc_debe_map,
            acc_haber_map=acc_haber_map,
            rub_debe_map=rub_debe_map,
            rub_haber_map=rub_haber_map,
            ef2_acc_plus_map=ef2_acc_plus_map,
            ef2_acc_minus_map=ef2_acc_minus_map,
        )

        # Nueva hoja: Estructura EF-4 (Detalle)
        write_estructura_ef4_detalle(writer, equiv_bytes, df_equiv_ht, df_all, sheet_name="Estructura EF-4 (Detalle)")

    output.seek(0)
    return output

def build_excel_with_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Original
        df_original = pd.read_excel(BytesIO(main_bytes), dtype=str, engine="openpyxl")
        df_original.to_excel(writer, index=False, sheet_name="Original")

        # 2) Resultado General
        df_result.to_excel(writer, index=False, sheet_name="Resultado General")

        # 3) Particiones tipo_ctb (y conservar df_tipo1_sin1101 para sumar en G/H)
        mayor_series = df_result["mayor"].astype(str) if "mayor" in df_result.columns else pd.Series("", index=df_result.index)
        exp_con_1101 = set(df_result.loc[mayor_series.eq("1101"), "exp_contable"].unique())

        df_tipo1_sin1101 = pd.DataFrame()
        if "tipo_ctb" in df_result.columns:
            df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()
            in_1101 = df_tipo1["exp_contable"].isin(exp_con_1101)
            df_tipo1_con1101 = df_tipo1[in_1101].copy()
            df_tipo1_sin1101 = df_tipo1[~in_1101].copy()
            df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
            df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
        else:
            pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                writer, index=False, sheet_name="Avisos"
            )

        # 4) Copiar hoja HT EF-4 original desde Equivalencias y escribir G/H por Rubro (Tipo1_sin_1101)
        book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
        if COPIABLE_SHEET in book_equiv.sheetnames:
            src_ws = book_equiv[COPIABLE_SHEET]
            dst_ws = writer.book.create_sheet(COPIABLE_SHEET)
            copy_sheet_with_styles(src_ws, dst_ws)

            acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map = _compute_maps_para_estructura(df_result)
            rub_debe_map_norm = { _norm_text(k): v for k, v in (rub_debe_map or {}).items() }
            rub_haber_map_norm = { _norm_text(k): v for k, v in (rub_haber_map or {}).items() }

            if not df_tipo1_sin1101.empty and "Rubros" in df_tipo1_sin1101.columns:
                merged_ranges = dst_ws.merged_cells.ranges
                for i, row in enumerate(dst_ws.iter_rows(min_row=2), start=2):
                    rubro_val = row[1].value  # col B
                    rubro = str(rubro_val).strip() if rubro_val is not None else ""
                    if not rubro:
                        continue
                    rnorm = _norm_text(rubro)
                    debe_sum = float(rub_debe_map_norm.get(rnorm, 0.0))
                    haber_sum = float(rub_haber_map_norm.get(rnorm, 0.0))
                    if not is_inside_merged_area(i, 7, merged_ranges):
                        dst_ws.cell(row=i, column=7, value=debe_sum).number_format = '#,##0.00'
                    if not is_inside_merged_area(i, 8, merged_ranges):
                        dst_ws.cell(row=i, column=8, value=haber_sum).number_format = '#,##0.00'
        else:
            ws = writer.book.create_sheet("Aviso_HT_EF4")
            ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")

        # 5) Hojas nuevas
        write_ht_ef_4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")

        # Para Estructura (con Debe/Haber por cuenta y EF-2 en filas de cuenta) + Auditoría + Totales
        acc_debe_map, acc_haber_map, rub_debe_map, rub_haber_map = _compute_maps_para_estructura(df_result)
        ef2_acc_plus_map, ef2_acc_minus_map = _compute_ef2_variaciones_por_cuenta(equiv_bytes, df_equiv_ht)

        df_all = write_ht_ef4_estructura(
            writer, equiv_bytes, df_equiv_ht,
            sheet_name="HT EF-4 (Estructura)",
            acc_debe_map=acc_debe_map,
            acc_haber_map=acc_haber_map,
            rub_debe_map=rub_debe_map,
            rub_haber_map=rub_haber_map,
            ef2_acc_plus_map=ef2_acc_plus_map,
            ef2_acc_minus_map=ef2_acc_minus_map,
        )

        # Nueva hoja: Estructura EF-4 (Detalle)
        write_estructura_ef4_detalle(writer, equiv_bytes, df_equiv_ht, df_all, sheet_name="Estructura EF-4 (Detalle)")

    output.seek(0)
    return output

# =============================
# UI
# =============================
opt_col1, opt_col2 = st.columns([1, 1])
with opt_col1:
    copy_ht = st.checkbox(
        "Incluir copia de HT EF-4 (original, con estilos + sumas en G/H)",
        value=True,
        help="Si no marcas esta opción, igual se crearán las hojas HT EF-4 (Compilada), HT EF-4 (Estructura) y Estructura EF-4 (Detalle)."
    )
with opt_col2:
    st.caption("El 2º archivo (Equivalencias) debe contener 'Hoja de Trabajo' (con EF-4 y Actividad), 'EF-1 Apertura', 'EF-1 Final', opcional 'EF-2 Final' y preferible 'Estructura EF-4'.")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"], key="main")
with col2:
    equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo, EF-1 Apertura, EF-1 Final, EF-2 Final)", type=["xlsx"], key="equiv")

if uploaded_file and equiv_file:
    try:
        main_bytes = _read_file_bytes(uploaded_file)
        equiv_bytes = _read_file_bytes(equiv_file)

        # Carga y procesamiento
        df_main = load_main_df(main_bytes)
        df_equiv_ht = load_equiv_df(equiv_bytes)  # Mapeo Rubros/EF-4/Actividad desde "Hoja de Trabajo"
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
