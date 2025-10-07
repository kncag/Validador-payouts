import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import json
import requests
from datetime import datetime

st.set_page_config(page_title="Validador Excel", layout="centered")
st.title("📊 Validador y Analizador de Archivos Excel")

# ---------- Utilidades generales ----------
def normalize_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", "", str(val)).strip().lower()

def parse_number(val):
    try:
        if val is None:
            return np.nan
        s = str(val).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return np.nan
        s = re.sub(r"\s+", "", s)
        if "." in s:
            s = s.replace(",", "")
            if s.count(".") > 1:
                parts = s.split(".")
                s = "".join(parts[:-1]) + "." + parts[-1]
            return float(s)
        else:
            s = s.replace(",", "")
            return float(s)
    except:
        return np.nan

def safe_str_preserve(val):
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r"\.0+$", "", s)
    return s

def get_col_by_letter(letter, df):
    try:
        idx = ord(letter.upper()) - ord("A")
        return df.columns[idx]
    except:
        return None

def normalize_doc_for_match(s: str) -> str:
    s = "" if pd.isna(s) else str(s).strip()
    digits = re.sub(r"\D", "", s)
    return digits

def find_row_by_document_positional(orig_df, doc_val):
    """
    Busca la primera fila en orig_df cuya columna B o C (posicional) coincida con doc_val
    aplicando normalizaciones sucesivas. Devuelve la Series de la fila o None.
    """
    candidate_cols = []
    colB = get_col_by_letter("B", orig_df)
    colC = get_col_by_letter("C", orig_df)
    if colB is not None:
        candidate_cols.append(colB)
    if colC is not None and colC not in candidate_cols:
        candidate_cols.append(colC)
    if not candidate_cols:
        return None

    target_raw = "" if pd.isna(doc_val) else str(doc_val).strip()
    try:
        as_float = float(target_raw)
        as_int = str(int(as_float))
    except Exception:
        as_int = None
    target_digits = re.sub(r"\D", "", target_raw)

    for col in candidate_cols:
        series = orig_df[col].astype(str).apply(lambda x: str(x).strip())

        # exact raw match
        mask = series == target_raw
        if mask.any():
            return orig_df.loc[mask].iloc[0]

        # int-like match
        if as_int:
            mask_int = series == as_int
            if mask_int.any():
                return orig_df.loc[mask_int].iloc[0]

        # digits-only match
        if target_digits:
            s_digits = series.apply(lambda x: re.sub(r"\D", "", x))
            if (s_digits == target_digits).any():
                return orig_df.loc[s_digits == target_digits].iloc[0]
            for L in (8, 9, 11):
                if len(target_digits) <= L:
                    tz = target_digits.zfill(L)
                    if (s_digits == tz).any():
                        return orig_df.loc[s_digits == tz].iloc[0]
            s_nozeros = s_digits.apply(lambda x: x.lstrip("0"))
            if (s_nozeros == target_digits.lstrip("0")).any():
                return orig_df.loc[s_nozeros == target_digits.lstrip("0")].iloc[0]
            if len(target_digits) >= 4:
                tail_mask = s_digits.str.endswith(target_digits)
                if tail_mask.any():
                    return orig_df.loc[tail_mask].iloc[0]

    return None

# ---------- Constantes RECH ----------
ENDPOINT = "https://q6caqnpy09.execute-api.us-east-1.amazonaws.com/OPS/kpayout/v1/payout_process/reject_invoices_batch"

OUT_COLS = [
    "dni/cex",
    "nombre",
    "importe",
    "Referencia",
    "Estado",
    "Codigo de Rechazo",
    "Descripcion de Rechazo",
]

SUBSET_COLS = [
    "Referencia",
    "Estado",
    "Codigo de Rechazo",
    "Descripcion de Rechazo",
]

ESTADO = "rechazada"

# ---------- Funciones RECH ----------
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rechazos")
    return buf.getvalue()

def post_to_endpoint(excel_bytes: bytes) -> tuple[int, str]:
    files = {
        "edt": (
            "rechazos.xlsx",
            excel_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    }
    resp = requests.post(ENDPOINT, files=files)
    return resp.status_code, resp.text

def rech_post_handler(df: pd.DataFrame, ui_feedback_callable=None) -> tuple[bool, str]:
    if list(df.columns) != OUT_COLS:
        msg = f"Encabezados inválidos. Se requieren: {OUT_COLS}"
        if ui_feedback_callable:
            ui_feedback_callable("error", msg)
        return False, msg
    payload = df[SUBSET_COLS]
    try:
        excel_bytes = df_to_excel_bytes(payload)
    except Exception as e:
        msg = f"Error generando Excel: {e}"
        if ui_feedback_callable:
            ui_feedback_callable("error", msg)
        return False, msg
    try:
        status, resp_text = post_to_endpoint(excel_bytes)
    except Exception as e:
        msg = f"Error realizando POST: {e}"
        if ui_feedback_callable:
            ui_feedback_callable("error", msg)
        return False, msg
    msg = f"{status}: {resp_text}"
    if ui_feedback_callable:
        if 200 <= status < 300:
            ui_feedback_callable("success", msg)
        else:
            ui_feedback_callable("error", msg)
    return (200 <= status < 300), msg

# ---------- UI: uploader, Lista Negra y checkbox en la misma fila ----------
col_u, col_l, col_cb = st.columns([2, 3, 1])
with col_u:
    uploaded_files = st.file_uploader("📁 Suba uno o varios archivos Excel", type=["xlsx"], accept_multiple_files=True)
with col_l:
    lista_negra_input = st.text_input("🔎 Lista Negra (ingresa uno o más criterios separados por coma)")
with col_cb:
    include_ref = st.checkbox("Incluir I como Referencia", value=True)

THRESHOLD_FIXED = 30000

# Acumuladores
matches_report = []
duplicates_report = []
threshold_report = []
validation_report = []  # list of DataFrames (report_df)
originals = []  # lista de DataFrames originales leídos (mantener en paralelo)
error_log = []

# ---------- Procesamiento de archivos si hay uploads ----------
if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=0, dtype=str)
            df.columns = [str(col) for col in df.columns]

            originals.append(df.copy())

            # LISTA NEGRA (mismo criterio previo)
            if lista_negra_input:
                criteria = [normalize_text(x) for x in lista_negra_input.split(",") if x.strip() != ""]
                if criteria:
                    norm = df.applymap(lambda x: normalize_text(x))
                    mask = False
                    for c in criteria:
                        mask = mask | norm.isin([c])
                    matches = df[mask.any(axis=1)].copy()
                else:
                    matches = pd.DataFrame()
            else:
                matches = pd.DataFrame()
            if not matches.empty:
                matches["Archivo"] = file.name
                matches_report.append(matches)

            # DUPLICADOS (posicional)
            dup_letters = ["C", "D", "M", "R", "S"]
            if include_ref:
                dup_letters.insert(2, "I")
            dup_cols = []
            missing_dup = []
            for lt in dup_letters:
                c = get_col_by_letter(lt, df)
                if c is None:
                    missing_dup.append(lt)
                else:
                    dup_cols.append(c)
            if missing_dup:
                error_log.append(f"❌ Columnas para duplicados faltantes {missing_dup} en {file.name}")
            else:
                subset_df = df[dup_cols].fillna("").astype(str).applymap(lambda x: x.strip())
                duplicated_mask = subset_df.duplicated(keep=False)
                if duplicated_mask.any():
                    dups_report = df.loc[duplicated_mask].copy()
                    dups_report["Archivo"] = file.name
                    dups_report["Columnas comprobadas"] = ",".join(dup_letters)
                    duplicates_report.append(dups_report)

            # IMPORTES MAYORES A 30,000
            col_M = get_col_by_letter("M", df)
            if col_M is None:
                error_log.append(f"❌ Columna M no encontrada en {file.name}")
            else:
                try:
                    df["_M_num"] = df[col_M].apply(parse_number)
                    filtered = df[df["_M_num"] >= THRESHOLD_FIXED]
                    extract_letters = ["B", "C", "D", "L", "M"]
                    extract_cols = []
                    missing_extract = []
                    for lt in extract_letters:
                        c = get_col_by_letter(lt, df)
                        if c is None:
                            missing_extract.append(lt)
                        else:
                            extract_cols.append(c)
                    if missing_extract:
                        error_log.append(f"❌ Columnas faltantes para extracción {missing_extract} en {file.name}")
                    elif filtered is not None and not filtered.empty:
                        out = filtered[extract_cols].copy()
                        out["Archivo"] = file.name
                        threshold_report.append(out)
                except Exception as e:
                    error_log.append(f"❌ Error procesando importes en {file.name}: {e}")

            # DOCUMENTOS ERRADOS (posicional B -> C)
            col_B = get_col_by_letter("B", df)
            col_C = get_col_by_letter("C", df)
            if col_B is None or col_C is None:
                error_log.append(f"❌ Columnas B o C faltantes en {file.name}")
            else:
                try:
                    tipos = df[col_B].astype(str).apply(lambda x: safe_str_preserve(x).strip().upper())
                    numeros = df[col_C].astype(str).apply(lambda x: safe_str_preserve(x).strip())
                    def validate_pair(tipo, valor_raw):
                        if valor_raw == "" or valor_raw.lower() in {"nan", "none"}:
                            if tipo in {"DNI", "CEX", "RUC"}:
                                return f"{tipo} inválido - vacío"
                            return None
                        if tipo == "DNI":
                            if not valor_raw.isdigit() or len(valor_raw) != 8:
                                return "DNI inválido"
                            return None
                        elif tipo == "CEX":
                            valor_cex = valor_raw.zfill(9)
                            if not valor_cex.isdigit() or len(valor_cex) != 9:
                                return "CEX inválido"
                            return None
                        elif tipo == "RUC":
                            valor_ruc = valor_raw.zfill(11)
                            if not valor_ruc.isdigit() or len(valor_ruc) != 11:
                                return "RUC inválido"
                            return None
                        return None
                    errors_series = [validate_pair(t, n) for t, n in zip(tipos, numeros)]
                    report_df = pd.DataFrame({
                        "TipoDocumento": tipos.values,
                        "Documento": numeros.values,
                        "Error": errors_series
                    })
                    report_df = report_df[report_df["Error"].notna()].copy()
                    if not report_df.empty:
                        report_df["Archivo"] = file.name
                        validation_report.append(report_df)
                except Exception as e:
                    error_log.append(f"❌ Error en validación B/C en {file.name}: {e}")

        except Exception as e:
            error_log.append(f"❌ Error procesando {file.name}: {e}")

# ---------- Renderizado de secciones ----------
# LISTA NEGRA: preview + botones (R002 / CUENTA INVALIDA)
if matches_report:
    matches_df = pd.concat(matches_report, ignore_index=True)
    st.subheader("Lista Negra")
    st.dataframe(matches_df)

    out_rows_ln = []
    for _, r in matches_df.iterrows():
        try:
            dni = ""
            nombre = ""
            referencia = ""
            importe_val = ""
            if "DOCUMENTO" in r.index:
                dni = safe_str_preserve(r["DOCUMENTO"])
            elif "B" in r.index:
                dni = safe_str_preserve(r["B"])
            else:
                try:
                    dni = safe_str_preserve(r.iloc[1])
                except Exception:
                    dni = ""
            if "NOMBRE" in r.index:
                nombre = safe_str_preserve(r["NOMBRE"])
            elif "D" in r.index:
                nombre = safe_str_preserve(r["D"])
            else:
                try:
                    nombre = safe_str_preserve(r.iloc[3])
                except Exception:
                    nombre = ""
            if "REFERENCIA" in r.index:
                referencia = safe_str_preserve(r["REFERENCIA"])
            elif "I" in r.index:
                referencia = safe_str_preserve(r["I"])
            else:
                try:
                    referencia = safe_str_preserve(r.iloc[8])
                except Exception:
                    referencia = ""
            if "MONTO" in r.index:
                importe_val = parse_number(r["MONTO"])
            elif "M" in r.index:
                importe_val = parse_number(r["M"])
            else:
                try:
                    importe_val = parse_number(r.iloc[12])
                except Exception:
                    importe_val = ""
        except Exception:
            dni = safe_str_preserve(r.get("Documento",""))
            nombre = ""
            referencia = ""
            importe_val = ""
        out_rows_ln.append({
            "dni/cex": dni,
            "nombre": nombre,
            "importe": importe_val,
            "Referencia": referencia,
            "Estado": ESTADO,
            "Codigo de Rechazo": "R002",
            "Descripcion de Rechazo": "CUENTA INVALIDA",
        })
    df_out_ln = pd.DataFrame(out_rows_ln, columns=OUT_COLS)

    st.markdown("**Preview (exactamente lo que se enviará al endpoint - Lista Negra)**")
    st.dataframe(df_out_ln)

    btn_ln_1, btn_ln_2 = st.columns([1,1])
    with btn_ln_1:
        if st.button("RECH-POSTMAN - Lista Negra"):
            sent_ok, message = rech_post_handler(df_out_ln, ui_feedback_callable=lambda lvl, m: getattr(st, lvl)(m))
            if sent_ok:
                st.info("Envío completado correctamente.")
            else:
                st.error(f"Envío fallido: {message}")
    with btn_ln_2:
        excel_bytes = df_to_excel_bytes(df_out_ln)
        st.download_button("⬇️ Descargar", data=excel_bytes, file_name="lista_negra_rechazos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# DUPLICADOS: preview + botones (R019 / ID DE TRANSACCIÓN DUPLICADA)
if duplicates_report:
    dup_df = pd.concat(duplicates_report, ignore_index=True)
    st.subheader("Duplicados")
    st.dataframe(dup_df)

    # Construir df_out_dup mapeando desde las filas duplicadas
    out_rows_dup = []
    for _, r in dup_df.iterrows():
        try:
            dni = ""
            nombre = ""
            referencia = ""
            importe_val = ""
            # extraer por posición/encabezado similar a Lista Negra
            if "DOCUMENTO" in r.index:
                dni = safe_str_preserve(r["DOCUMENTO"])
            elif "B" in r.index:
                dni = safe_str_preserve(r["B"])
            else:
                try:
                    dni = safe_str_preserve(r.iloc[1])
                except Exception:
                    dni = ""
            if "NOMBRE" in r.index:
                nombre = safe_str_preserve(r["NOMBRE"])
            elif "D" in r.index:
                nombre = safe_str_preserve(r["D"])
            else:
                try:
                    nombre = safe_str_preserve(r.iloc[3])
                except Exception:
                    nombre = ""
            if "REFERENCIA" in r.index:
                referencia = safe_str_preserve(r["REFERENCIA"])
            elif "I" in r.index:
                referencia = safe_str_preserve(r["I"])
            else:
                try:
                    referencia = safe_str_preserve(r.iloc[8])
                except Exception:
                    referencia = ""
            if "MONTO" in r.index:
                importe_val = parse_number(r["MONTO"])
            elif "M" in r.index:
                importe_val = parse_number(r["M"])
            else:
                try:
                    importe_val = parse_number(r.iloc[12])
                except Exception:
                    importe_val = ""
        except Exception:
            dni = safe_str_preserve(r.get("Documento",""))
            nombre = ""
            referencia = ""
            importe_val = ""
        out_rows_dup.append({
            "dni/cex": dni,
            "nombre": nombre,
            "importe": importe_val,
            "Referencia": referencia,
            "Estado": ESTADO,
            "Codigo de Rechazo": "R019",
            "Descripcion de Rechazo": "ID DE TRANSACCIÓN DUPLICADA",
        })
    df_out_dup = pd.DataFrame(out_rows_dup, columns=OUT_COLS)

    st.markdown("**Preview (exactamente lo que se enviará al endpoint - Duplicados)**")
    st.dataframe(df_out_dup)

    btn_dup_1, btn_dup_2 = st.columns([1,1])
    with btn_dup_1:
        if st.button("RECH-POSTMAN - Duplicados"):
            sent_ok, message = rech_post_handler(df_out_dup, ui_feedback_callable=lambda lvl, m: getattr(st, lvl)(m))
            if sent_ok:
                st.info("Envío completado correctamente.")
            else:
                st.error(f"Envío fallido: {message}")
    with btn_dup_2:
        excel_bytes = df_to_excel_bytes(df_out_dup)
        st.download_button("⬇️ Descargar", data=excel_bytes, file_name="duplicados_rechazos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# IMPORTES MAYORES A 30,000
if threshold_report:
    th_df = pd.concat(threshold_report, ignore_index=True)
    st.subheader("Importes mayores a 30,000")
    st.dataframe(th_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        th_df.to_excel(writer, index=False, sheet_name="importes_mayores")
    st.download_button("⬇️ Descargar importes", data=buf.getvalue(), file_name="importes_mayores.xlsx")

# DOCUMENTOS ERRADOS (preview + botones; mapea desde originals por Documento)
if validation_report:
    val_df = pd.concat(validation_report, ignore_index=True)
    st.subheader("Documentos errados")
    st.dataframe(val_df)

    # Construir df_out mapeando desde los originales por columna B o C posicional
    out_rows = []
    unmapped_docs = []
    for _, err_row in val_df.iterrows():
        doc_val = err_row.get("Documento", "")
        mapped = False
        for orig_df in originals:
            candidate = find_row_by_document_positional(orig_df, doc_val)
            if candidate is not None:
                colB = get_col_by_letter("B", orig_df)
                colD = get_col_by_letter("D", orig_df)
                colI = get_col_by_letter("I", orig_df)
                colM = get_col_by_letter("M", orig_df)
                dni = safe_str_preserve(candidate[colB]) if colB else ""
                nombre = safe_str_preserve(candidate[colD]) if colD else ""
                referencia = safe_str_preserve(candidate[colI]) if colI else ""
                monto_num = parse_number(candidate[colM]) if colM else ""
                importe_val = monto_num if (monto_num is not None and not (isinstance(monto_num, float) and np.isnan(monto_num))) else ""
                out_rows.append({
                    "dni/cex": dni,
                    "nombre": nombre,
                    "importe": importe_val,
                    "Referencia": referencia,
                    "Estado": ESTADO,
                    "Codigo de Rechazo": "R001",
                    "Descripcion de Rechazo": "DOCUMENTO ERRADO",
                })
                mapped = True
                break
        if not mapped:
            unmapped_docs.append(doc_val)
            out_rows.append({
                "dni/cex": doc_val,
                "nombre": "",
                "importe": "",
                "Referencia": "",
                "Estado": ESTADO,
                "Codigo de Rechazo": "R001",
                "Descripcion de Rechazo": "DOCUMENTO ERRADO",
            })

    df_out = pd.DataFrame(out_rows, columns=OUT_COLS)

    if unmapped_docs:
        st.warning(f"No se pudo mapear Referencia/nombre/importe para {len(unmapped_docs)} documento(s). Ejemplos: {unmapped_docs[:5]}")

    st.markdown("**Preview (exactamente lo que se enviará al endpoint)**")
    st.dataframe(df_out)

    # Botones: enviar (RECH-POSTMAN) y descargar (descarga exactamente preview)
    btn1, btn2 = st.columns([1, 1])
    with btn1:
        if st.button("RECH-POSTMAN"):
            sent_ok, message = rech_post_handler(df_out, ui_feedback_callable=lambda lvl, m: getattr(st, lvl)(m))
            if sent_ok:
                st.info("Envío completado correctamente.")
            else:
                st.error(f"Envío fallido: {message}")
    with btn2:
        excel_bytes = df_to_excel_bytes(df_out)
        st.download_button("⬇️ Descargar", data=excel_bytes, file_name="documentos_errados_rechazos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ERROR DE ARCHIVO
if error_log:
    st.subheader("Error de archivo")
    for err in error_log:
        st.error(err)
