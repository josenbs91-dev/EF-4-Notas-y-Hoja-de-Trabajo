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
# (1) HT EF-4 (Compilada): EF-1 Apertura + EF-1 Final con columna "Rubros" añadida a la derecha
# =============================
def write_ht_ef4_compilada(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Compilada)"):
    # Map Cuentas -> Rubros (desde "Hoja de Trabajo")
    map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))

    ap_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*apert"])
    fi_name = _find_sheet_name(equiv_bytes, [r"ef\s*1.*final"])

    startrow = 0
    for label, real_name in [("EF-1 Apertura", ap_name), ("EF-1 Final", fi_name)]:
        # título de sección
        pd.DataFrame([["Sección:", label]]).to_excel(
            writer, sheet_name=sheet_name, startrow=startrow, index=False, header=False
        )
        startrow += 1

        if not real_name:
            pd.DataFrame([["(No se encontró la hoja solicitada)"]]).to_excel(
                writer, sheet_name=sheet_name, startrow=startrow, index=False, header=False
            )
            startrow += 2
            continue

        df_sec = pd.read_excel(BytesIO(equiv_bytes), sheet_name=real_name, dtype=str, engine="openpyxl")
        df_sec.columns = [str(c).strip() for c in df_sec.columns]
        cta_col = _pick_col(df_sec, ["Cuentas Contables", "Cuenta Contable", "Cuenta", "Cuentas"], must_contain="cuent")
        df_out = df_sec.copy()
        if cta_col:
            df_out["Rubros"] = df_out[cta_col].astype(str).map(map_cta_to_rubro).fillna("")
        else:
            df_out["Rubros"] = ""
        # escribir la tabla con Rubros agregada
        df_out.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
        startrow += len(df_out) + 2  # línea en blanco

# =============================
# (2) HT EF-4 (Estructura): Respeta ESTRICTAMENTE la hoja "Estructura del archivo"
#     Rubro → Cuentas únicas (Apertura+Final) + Importes (Apertura | Final)
#     Incluye rubros sin cuentas con fila "(sin cuentas)" y aplica formato
# =============================
def write_ht_ef4_estructura(writer, equiv_bytes: bytes, df_equiv_ht: pd.DataFrame, sheet_name: str = "HT EF-4 (Estructura)"):
    # Mapeo de Cuentas -> Rubros desde "Hoja de Trabajo"
    map_cta_to_rubro = dict(zip(df_equiv_ht["Cuentas Contables"], df_equiv_ht["Rubros"]))

    # Localizar hojas EF-1
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

    # --- Consolidado (cuentas únicas y sus importes)
    cuentas = sorted(set(ap_df["Cuenta"]).union(set(fi_df["Cuenta"])))
    data = []
    for cta in cuentas:
        rub = map_cta_to_rubro.get(cta, "")
        ap_val = float(ap_df.loc[ap_df["Cuenta"] == cta, "Importe"].sum()) if not ap_df.empty else 0.0
        fi_val = float(fi_df.loc[fi_df["Cuenta"] == cta, "Importe"].sum()) if not fi_df.empty else 0.0
        data.append({"Rubros": rub, "Cuenta Contable": cta, "EF-1 Apertura": ap_val, "EF-1 Final": fi_val})
    df_all = pd.DataFrame(data)

    if df_all.empty:
        pd.DataFrame({"info": ["No se pudieron consolidar Importes de EF-1 Apertura/Final."]}).to_excel(
            writer, index=False, sheet_name=sheet_name
        )
        return

    # --- Obtener ORDEN ESTRICTO desde 'Estructura del archivo'
    def struct_order_strict() -> list | None:
        struct_name = _find_sheet_name(equiv_bytes, [r"estructura.*archivo", r"estructura"])
        if not struct_name:
            return None
        try:
            df_struct = pd.read_excel(BytesIO(equiv_bytes), sheet_name=struct_name, dtype=str, engine="openpyxl")
            df_struct.columns = [str(c).strip() for c in df_struct.columns]
            rub_col = _pick_col(df_struct, ["Rubros", "Rubro"], must_contain="rubro")
            ord_col = _pick_col(df_struct, ["Orden", "Order", "Ordenamiento", "N°", "No", "Nro"], must_contain="orden")

            if rub_col is None:
                return None

            if ord_col:
                tmp = df_struct[[ord_col, rub_col]].copy()
                tmp[ord_col] = pd.to_numeric(tmp[ord_col], errors="coerce")
                tmp["_row"] = np.arange(len(tmp))
                tmp = tmp.sort_values([ord_col, "_row"], na_position="last", kind="mergesort")
                ordered_rubros = [str(x).strip() for x in tmp[rub_col].tolist()]
            else:
                # Sin 'Orden': usar fila tal cual
                ordered_rubros = [str(x).strip() for x in df_struct[rub_col].tolist()]

            # Quitar vacíos y duplicados preservando la primera aparición
            seen, final = set(), []
            for r in ordered_rubros:
                if r and r not in seen:
                    seen.add(r)
                    final.append(r)
            return final
        except Exception:
            return None

    strict_order = struct_order_strict()

    # Conjunto de rubros presentes en datos
    rubros_presentes = set(df_all["Rubros"].astype(str).str.strip().unique())

    if strict_order:
        # Usar SOLO los rubros que están en la estructura y también presentes en datos, en ese mismo orden.
        # Además, si en la estructura hay rubros sin cuentas en datos, IGUAL mostrarlos con "(sin cuentas)".
        rubros_order = []
        for r in strict_order:
            r_clean = str(r).strip()
            if r_clean:
                rubros_order.append(r_clean)
    else:
        # Fallback: solo rubros presentes, orden alfabético
        rubros_order = sorted(rubros_presentes)

    # --- Construir layout final (incluyendo rubros sin cuentas)
    rows = []
    rows.append(["", "Rubro", "Cuenta Contable", "EF-1 Apertura", "EF-1 Final"])  # encabezado
    for rub in rubros_order:
        rows.append(["", rub, "", "", ""])
        block = df_all[df_all["Rubros"].astype(str).str.strip() == str(rub)]
        block = block.sort_values(["Cuenta Contable"]).drop_duplicates(subset=["Cuenta Contable"], keep="first")

        if block.empty:
            # Mostrar rubro aunque no tenga cuentas
            rows.append(["", "", "(sin cuentas)", 0.0, 0.0])
        else:
            for _, r in block.iterrows():
                rows.append(["", "", r["Cuenta Contable"], float(r["EF-1 Apertura"]), float(r["EF-1 Final"])])

        rows.append(["", "", "", "", ""])  # línea en blanco

    out_df = pd.DataFrame(rows[1:], columns=rows[0])
    out_df.to_excel(writer, index=False, sheet_name=sheet_name)

    # --------- FORMATO con openpyxl ---------
    ws = writer.book[sheet_name]
    max_row = ws.max_row
    # Anchuras
    widths = {2: 42, 3: 22, 4: 18, 5: 18}  # B..E
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Encabezado (fila 1)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFEFEFEF")
    center = Alignment(horizontal="center", vertical="center")
    for c in range(2, 6):  # B..E
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center

    # Estilo números columnas D y E
    num_align = Alignment(horizontal="right")
    for r in range(2, max_row + 1):
        for c in [4, 5]:
            ws.cell(row=r, column=c).number_format = '#,##0.00'
            ws.cell(row=r, column=c).alignment = num_align

    # Negritas para filas de Rubro (B con texto y C vacío o '(sin cuentas)')
    rubro_fill = PatternFill("solid", fgColor="FFF7F7F7")
    rubro_font = Font(bold=True)
    for r in range(2, max_row + 1):
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        if (b is not None and str(b).strip() != "") and (c is None or str(c).strip() == ""):
            ws.cell(row=r, column=2).font = rubro_font
            for col in range(2, 6):
                ws.cell(row=r, column=col).fill = rubro_fill

    # Bordes finos en toda la tabla B1:Emax
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(1, max_row + 1):
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = border

    # Autofiltro y freeze panes
    ws.auto_filter.ref = f"B1:E{max_row}"
    ws.freeze_panes = "B2"

# =============================
# Exportadores (usamos SIEMPRE openpyxl para poder formatear)
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

        # === NUEVAS hojas ===
        write_ht_ef4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")
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

        # 3) Particiones tipo_ctb
        if "tipo_ctb" in df_result.columns:
            df_tipo1 = df_result[df_result["tipo_ctb"].astype(str) == "1"].copy()
            in_1101 = df_tipo1["exp_contable"].isin(set(df_result.loc[df_result.get("mayor", "").astype(str).eq("1101"), "exp_contable"].unique()))
            df_tipo1_con1101 = df_tipo1[in_1101].copy()
            df_tipo1_sin1101 = df_tipo1[~in_1101].copy()
            df_tipo1_con1101.to_excel(writer, index=False, sheet_name="Tipo1_con_1101")
            df_tipo1_sin1101.to_excel(writer, index=False, sheet_name="Tipo1_sin_1101")
        else:
            df_tipo1_sin1101 = pd.DataFrame()
            pd.DataFrame({"info": ["No se encontró la columna 'tipo_ctb' en el archivo principal."]}).to_excel(
                writer, index=False, sheet_name="Avisos"
            )

        # 4) Copiar hoja HT EF-4 original desde Equivalencias (si existe)
        book_equiv = openpyxl.load_workbook(BytesIO(equiv_bytes))
        if COPIABLE_SHEET in book_equiv.sheetnames:
            src_ws = book_equiv[COPIABLE_SHEET]
            dst_ws = writer.book.create_sheet(COPIABLE_SHEET)
            copy_sheet_with_styles(src_ws, dst_ws)
        else:
            ws = writer.book.create_sheet("Aviso_HT_EF4")
            ws.cell(row=1, column=1, value=f"No se encontró la hoja '{COPIABLE_SHEET}' en el archivo de Equivalencias.")

        # 5) === NUEVAS hojas ===
        write_ht_ef4_compilada(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Compilada)")
        write_ht_ef4_estructura(writer, equiv_bytes, df_equiv_ht, sheet_name="HT EF-4 (Estructura)")

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
        help="Si no marcas esta opción, igual se crearán las hojas HT EF-4 (Compilada) y HT EF-4 (Estructura)."
    )
with opt_col2:
    st.caption("El 2º archivo (Equivalencias) debe contener 'Hoja de Trabajo' y, si es posible, 'EF-1 Apertura' y 'EF-1 Final' (y opcionalmente 'Estructura del archivo').")

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
