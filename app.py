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

    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    missing = [c for c in REQUIRED_FOR_EXP if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas en el archivo principal: {', '.join(missing)}")

    parts = [df[c].astype(str).fillna("") for c in REQUIRED_FOR_EXP]
    df["exp_contable"] = parts[0] + "-" + parts[1] + "-" + parts[2] + "-" + parts[3]

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
    mask_1101 = df.get("mayor", "").astype(str).eq("1101")
    exp_con_1101 = set(df.loc[mask_1101, "exp_contable"].unique())
    in_1101 = df["exp_contable"].isin(exp_con_1101)

    debe = pd.to_numeric(df.get("debe", 0), errors="coerce").fillna(0.0)
    haber = pd.to_numeric(df.get("haber", 0), errors="coerce").fillna(0.0)

    df["debe_adj"] = np.where(in_1101, haber, debe)
    df["haber_adj"] = np.where(in_1101, debe, haber)

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
# (1) HT EF-4 (Compilada)
# =============================
def write_ht_ef_4_compilada(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Compilada)"):
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
            row_ptr += 2
            return

        df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=real_name, dtype=str, engine="openpyxl")
        df_sec.columns = [str(c).strip() for c in df_sec.columns]
        cta_col = _pick_col(df_sec, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
        df_out = df_sec.copy()
        if cta_col:
            df_out["Rubros"] = df_out[cta_col].astype(str).map(map_cta_to_rubro).fillna("")
        else:
            df_out["Rubros"] = ""

        # headers
        for j, col in enumerate(df_out.columns, start=1):
            ws.cell(row=row_ptr, column=j, value=col)
        r = row_ptr + 1
        for _, rr in df_out.iterrows():
            for j, col in enumerate(df_out.columns, start=1):
                ws.cell(row=r, column=j, value=str(rr[col]) if pd.notna(rr[col]) else "")
            r += 1
        row_ptr = r + 1  # blank line

    dump_section("EF-1 Apertura", ap_name)
    dump_section("EF-1 Final", fi_name)

# =============================
# (2) HT EF-4 (Estructura) + Variaciones y Totales por Rubro
# =============================
def write_ht_ef4_estructura(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Estructura)"):
    map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))
    ap_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*apert"])
    fi_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*final"])

    # --- Lectura de importes por cuenta (Apertura / Final) ---
    def read_importes(sheet_name: str) -> pd.DataFrame:
        if not sheet_name:
            return pd.DataFrame(columns=["Cuenta", "Importe"])
        df = pd.read_excel(BytesIO(equiv_bytes), sheet_name=sheet_name, dtype=str, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        cta_col = _pick_col(df, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
        imp_col = _pick_col(df, ["Importes", "Importe", "Monto", "Valor", "Importe S/.", "Importe S"], must_contain="import")
        if not cta_col:
            return pd.DataFrame(columns=["Cuenta", "Importe"])
        vals = pd.to_numeric(df[imp_col], errors="coerce").fillna(0.0) if imp_col else 0.0
        out = pd.DataFrame({"Cuenta": df[cta_col].astype(str).str.strip(), "Importe": vals})
        out = out.groupby("Cuenta", as_index=False)["Importe"].sum()
        return out

    ap_df = read_importes(ap_name)
    fi_df = read_importes(fi_name)

    # --- Consolidado base
    cuentas = sorted(set(ap_df["Cuenta"]).union(set(fi_df["Cuenta"])))
    rows_data = []
    for cta in cuentas:
        rub = map_cta_to_rubro.get(cta, "")
        ap_val = float(ap_df.loc[ap_df["Cuenta"] == cta, "Importe"].sum()) if not ap_df.empty else 0.0
        fi_val = float(fi_df.loc[fi_df["Cuenta"] == cta, "Importe"].sum()) if not fi_df.empty else 0.0
        # diff y variaciones
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
        rows_data.append({
            "Rubros": rub,
            "Cuenta Contable": cta,
            "EF-1 Final": fi_val,        # <- primero
            "EF-1 Apertura": ap_val,     # <- luego
            "Variación +": var_plus,
            "Variación -": var_minus,
        })

    df_all = pd.DataFrame(rows_data)
    if df_all.empty:
        pd.DataFrame({"info": ["No se pudieron consolidar Importes de EF-1 Apertura/Final."]}).to_excel(
            writer, index=False, sheet_name=sheet_name
        )
        return

    # === Orden estricto desde "Estructura del archivo"
    def struct_order_strict() -> list | None:
        struct_name = _find_sheet_name(equiv_bytes, [r"estructura.*archivo", r"estructura"])
        if not struct_name:
            return None
        try:
            df_struct = pd.read_excel(BytesIO(equiv_bytes), sheet_name=struct_name, dtype=str, engine="openpyxl")
            df_struct.columns = [str(c).strip() for c in df_struct.columns]
            rub_col = _pick_col(df_struct, ["Rubros", "Rubro", "Estructura", "DESCRIPCION", "Descripción", "Descripcion"])
            if rub_col is None:
                rub_col = _pick_col(df_struct, [], must_contain="estruct") or _pick_col(df_struct, [], must_contain="descr")
            ord_col = _pick_col(df_struct, ["Orden", "Order", "Ordenamiento", "N°", "No", "Nro", "Nro."], must_contain="orden")
            if rub_col is None:
                return None
            if ord_col:
                tmp = df_struct[[ord_col, rub_col]].copy()
                tmp[ord_col] = pd.to_numeric(tmp[ord_col], errors="coerce")
                tmp["_row"] = np.arange(len(tmp))
                tmp = tmp.sort_values([ord_col, "_row"], na_position="last", kind="mergesort")
                ordered_rubros = [str(x).strip() for x in tmp[rub_col].tolist()]
            else:
                ordered_rubros = [str(x).strip() for x in df_struct[rub_col].tolist()]
            seen, final = set(), []
            for r in ordered_rubros:
                if r and r not in seen:
                    seen.add(r)
                    final.append(r)
            return final
        except Exception:
            return None

    strict_order = struct_order_strict()
    if strict_order:
        rubros_order = [str(r).strip() for r in strict_order if str(r).strip() != ""]
    else:
        rubros_order = sorted(set(df_all["Rubros"].astype(str).str.strip().unique()))

    # Totales por rubro (para escribir en la fila del rubro)
    totals = (
        df_all.groupby("Rubros")[["EF-1 Final", "EF-1 Apertura", "Variación +", "Variación -"]]
        .sum(numeric_only=True)
        .to_dict()
    )

    # --- Construcción del layout final
    # Encabezado: B..G
    header = ["", "Rubro", "Cuenta Contable", "EF-1 Final", "EF-1 Apertura", "Variación +", "Variación -"]
    out_rows = [header]
    for rub in rubros_order:
        out_rows.append(["", rub, "", "", "", "", ""])  # fila Rubro
        block = df_all[df_all["Rubros"].astype(str).str.strip() == str(rub)]
        block = block.sort_values(["Cuenta Contable"]).drop_duplicates(subset=["Cuenta Contable"], keep="first")
        if block.empty:
            out_rows.append(["", "", "(sin cuentas)", 0.0, 0.0, 0.0, 0.0])
        else:
            for _, r in block.iterrows():
                out_rows.append([
                    "", "",
                    r["Cuenta Contable"],
                    float(r["EF-1 Final"]),
                    float(r["EF-1 Apertura"]),
                    float(r["Variación +"]),
                    float(r["Variación -"]),
                ])
        out_rows.append(["", "", "", "", "", "", ""])  # línea en blanco

    out_df = pd.DataFrame(out_rows[1:], columns=out_rows[0])
    out_df.to_excel(writer, index=False, sheet_name=sheet_name)

    # --------- FORMATO + Totales por Rubro en la fila de Rubro ---------
    ws = writer.book[sheet_name]
    max_row = ws.max_row

    # Anchos: B..G
    widths = {2: 42, 3: 22, 4: 18, 5: 18, 6: 16, 7: 16}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Encabezado
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFEFEFEF")
    center = Alignment(horizontal="center", vertical="center")
    for c in range(2, 8):  # B..G
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    # Números en D..G
    num_align = Alignment(horizontal="right")
    for r in range(2, max_row + 1):
        for c in [4, 5, 6, 7]:
            ws.cell(row=r, column=c).number_format = '#,##0.00'
            ws.cell(row=r, column=c).alignment = num_align

    # Fila Rubro (col B con texto y col C vacía) + totales
    rubro_fill = PatternFill("solid", fgColor="FFF7F7F7")
    rubro_font = Font(bold=True)
    for r in range(2, max_row + 1):
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        if (b is not None and str(b).strip() != "") and (c is None or str(c).strip() == ""):
            # Estilo
            for col in range(2, 8):
                ws.cell(row=r, column=col).fill = rubro_fill
            ws.cell(row=r, column=2).font = rubro_font
            # Totales
            rub = str(b).strip()
            ws.cell(row=r, column=4, value=float(totals.get("EF-1 Final", {}).get(rub, 0.0))).font = rubro_font
            ws.cell(row=r, column=5, value=float(totals.get("EF-1 Apertura", {}).get(rub, 0.0))).font = rubro_font
            ws.cell(row=r, column=6, value=float(totals.get("Variación +", {}).get(rub, 0.0))).font = rubro_font
            ws.cell(row=r, column=7, value=float(totals.get("Variación -", {}).get(rub, 0.0))).font = rubro_font

    # Bordes finos en B1:Gmax
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(1, max_row + 1):
        for c in range(2, 8):
            ws.cell(row=r, column=c).border = border

    # Autofiltro y freeze panes
    ws.auto_filter.ref = f"B1:G{max_row}"
    ws.freeze_panes = "B2"

# =============================
# Exportadores (SIEMPRE openpyxl)
# =============================
def build_excel_without_ht(main_bytes: bytes, df_result: pd.DataFrame, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame) -> BytesIO:
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
            in_1101 = df_tipo1["exp_contable"].isin(set(df_result.loc[df_result.get("mayor", "").astype(str).eq("1101"), "exp_contable"].unique()))
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
        write_ht_ef4_estructura(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Estructura)")

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

        # 3) Particiones tipo_ctb (guardar df_tipo1_sin1101 para sumas G/H)
        df_tipo1_sin1101 = pd.DataFrame()
        if "tipo_ctb" in df_result.columns:
            df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()
            in_1101 = df_tipo1["exp_contable"].isin(set(df_result.loc[df_result.get("mayor", "").astype(str).eq("1101"), "exp_contable"].unique()))
            df_tipo1_con1101 = df_tipo1[in_1101].copy()
            df_tipo1_sin1101 = df_tipo1[~in_1101].copy()
            df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
            df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
        else:
            pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                writer, index=False, sheet_name="Avisos"
            )

        # 4) Copiar hoja HT EF-4 original desde Equivalencias y escribir G/H por Rubro
        book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
        if COPIABLE_SHEET in book_equiv.sheetnames:
            src_ws = book_equiv[COPIABLE_SHEET]
            dst_ws = writer.book.create_sheet(COPIABLE_SHEET)
            copy_sheet_with_styles(src_ws, dst_ws)

            # --- Sumas por Rubro (G = Debe, H = Haber) desde df_tipo1_sin1101 ---
            if (not df_tipo1_sin1101.empty) and ("Rubros" in df_tipo1_sin1101.columns):
                for col in ["debe_adj", "haber_adj"]:
                    if col in df_tipo1_sin1101.columns:
                        df_tipo1_sin1101[col] = pd.to_numeric(df_tipo1_sin1101[col], errors="coerce").fillna(0.0)
                df_sum = (
                    df_tipo1_sin1101.groupby("Rubros")[["debe_adj", "haber_adj"]]
                    .sum(numeric_only=True)
                    .reset_index()
                )
                dict_debe = dict(zip(df_sum["Rubros"], df_sum["debe_adj"]))
                dict_haber = dict(zip(df_sum["Rubros"], df_sum["haber_adj"]))

                merged_ranges = dst_ws.merged_cells.ranges
                for i, row in enumerate(dst_ws.iter_rows(min_row=2), start=2):
                    rubro_val = row[1].value  # col B
                    rubro = str(rubro_val).strip() if rubro_val is not None else ""
                    if not rubro:
                        continue
                    debe_sum = float(dict_debe.get(rubro, 0.0))
                    haber_sum = float(dict_haber.get(rubro, 0.0))
                    if not is_inside_merged_area(i, 7, merged_ranges):
                        dst_ws.cell(row=i, column=7, value=debe_sum)
                        dst_ws.cell(row=i, column=7).number_format = '#,##0.00'
                    if not is_inside_merged_area(i, 8, merged_ranges):
                        dst_ws.cell(row=i, column=8, value=haber_sum)
                        dst_ws.cell(row=i, column=8).number_format = '#,##0.00'
        else:
            ws = writer.book.create_sheet("Aviso_HT_EF4")
            ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")

        # 5) Hojas nuevas
        write_ht_ef_4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")
        write_ht_ef4_estructura(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Estructura)")

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
        help="Si no marcas esta opción, igual se crearán las hojas HT EF-4 (Compilada) y HT EF-4 (Estructura)."
    )
with opt_col2:
    st.caption("El 2º archivo (Equivalencias) debe contener 'Hoja de Trabajo', 'EF-1 Apertura', 'EF-1 Final' y opcionalmente 'Estructura del archivo'.")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Sube tu archivo Excel principal", type=["xlsx"], key="main")
with col2:
    equiv_file = st.file_uploader("Sube tu archivo de Equivalencias (Hoja de Trabajo, EF-1 Apertura, EF-1 Final)", type=["xlsx"], key="equiv")

if uploaded_file and equiv_file:
    try:
        main_bytes = _read_file_bytes(uploaded_file)
        equiv_bytes = _read_file_bytes(equiv_file)

        df_main = load_main_df(main_bytes)
        df_equiv_ht = load_equiv_df(equiv_bytes)
        df_proc = compute_adjusted(df_main.copy())
        df_final = merge_equivalences(df_proc, df_equiv_ht)

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
