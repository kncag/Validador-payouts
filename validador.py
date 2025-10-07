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
        # Regla: '.' decimal, ',' miles
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

CODE_DESC = {
    "R001": "DOCUMENTO ERRADO",
    "R002": "CUENTA INVALIDA",
    "R007": "RECHAZO POR CCI",
}

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
    """
    Valida df y lo envía al ENDPOINT.
    - df debe tener columnas EXACTAS en el orden OUT_COLS.
    - ui_feedback_callable: opcional, función que recibe (level, message) e.g. lambda lvl,m: getattr(st, lvl)(m)
    Returns (sent_ok, message).
    """
    if list(df.columns) != OUT_COLS:
        msg = f"Encabezados inválidos. Se requieren: {OUT_COLS}"
        if ui_feedback_callable: ui_feedback_callable("error", msg)
        return False, msg

    payload = df[SUBSET_COLS]
    try:
        excel_bytes = df_to_excel_bytes(payload)
    except Exception as e:
        msg = f"Error generando Excel: {e}"
        if ui_feedback_callable: ui_feedback_callable("error", msg)
        return False, msg

    try:
        status, resp_text = post_to_endpoint(excel_bytes)
    except Exception as e:
        msg = f"Error realizando POST: {e}"
        if ui_feedback_callable: ui_feedback_callable("error", msg)
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

# Contenedores de resultados (mostraremos subtítulos solo si hay datos)
matches_report = []
duplicates_report = []
threshold_report = []
validation_report = []
error_log = []

# Procesamiento de archivos si hay uploads
if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=0, dtype=str)
            df.columns = [str(col) for col in df.columns]

            # Lista Negra: múltiples criterios
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

            # Duplicados: coincidencia completa en C, D, (I opcional), M, R, S
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

            # Importes mayores a 30,000 (umbral fijo)
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

            # Documentos errados: validación B -> C (DNI 8, CEX 9, RUC 11)
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

# ---------- Renderizado de secciones: solo mostrar subtítulos si hay datos ----------

# Lista Negra (primera)
if matches_report:
    matches_df = pd.concat(matches_report, ignore_index=True)
    st.subheader("Lista Negra")
    st.dataframe(matches_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        matches_df.to_excel(writer, index=False, sheet_name="lista_negra")
    st.download_button("⬇️ Descargar Lista Negra", data=buf.getvalue(), file_name="lista_negra.xlsx")

# Duplicados
if duplicates_report:
    dup_df = pd.concat(duplicates_report, ignore_index=True)
    st.subheader("Duplicados")
    st.dataframe(dup_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        dup_df.to_excel(writer, index=False, sheet_name="duplicados")
    st.download_button("⬇️ Descargar Duplicados", data=buf.getvalue(), file_name="duplicados.xlsx")

# Importes mayores a 30,000
if threshold_report:
    th_df = pd.concat(threshold_report, ignore_index=True)
    st.subheader("Importes mayores a 30,000")
    st.dataframe(th_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        th_df.to_excel(writer, index=False, sheet_name="importes_mayores")
    st.download_button("⬇️ Descargar importes", data=buf.getvalue(), file_name="importes_mayores.xlsx")

# Documentos errados + RECH UI (preview, endpoint, descarga complementaria)
if validation_report:
    val_df = pd.concat(validation_report, ignore_index=True)
    st.subheader("Documentos errados")

    # Transformación para endpoint (simple mapping)
    def transform_for_endpoint(df_in):
        df = df_in.copy()
        for col in ["TipoDocumento", "Documento", "Error", "Archivo"]:
            if col not in df.columns:
                df[col] = None
        payload = []
        for _, r in df.iterrows():
            payload.append({
                "tipo": str(r["TipoDocumento"]) if pd.notna(r["TipoDocumento"]) else "",
                "documento": str(r["Documento"]) if pd.notna(r["Documento"]) else "",
                "motivo": str(r["Error"]) if pd.notna(r["Error"]) else "",
                "origen": str(r["Archivo"]) if pd.notna(r["Archivo"]) else ""
            })
        return payload

    payload = transform_for_endpoint(val_df)

    # Botones y preview in fila
    btn_col1, btn_col2, btn_col3 = st.columns([2, 2, 2])

    with btn_col1:
        preview_df = pd.DataFrame(payload[:20])
        st.markdown("**Preview (primeros 20 items)**")
        st.dataframe(preview_df)

    with btn_col2:
        # Construir df_out con OUT_COLS si es posible; se usa Documento y Error como fuentes
        df_out = pd.DataFrame(columns=OUT_COLS)
        for _, r in val_df.iterrows():
            row = {
                "dni/cex": r.get("Documento", ""),
                "nombre": "",
                "importe": "",
                "Referencia": "",
                "Estado": ESTADO,
                "Codigo de Rechazo": "R001",
                "Descripcion de Rechazo": r.get("Error", "")
            }
            df_out = pd.concat([df_out, pd.DataFrame([row])], ignore_index=True)

        if st.button("📤 Enviar al endpoint (RECH-POSTMAN)"):
            sent_ok, message = rech_post_handler(df_out, ui_feedback_callable=lambda lvl, m: getattr(st, lvl)(m))
            if sent_ok:
                st.info("Envío completado correctamente.")
            else:
                st.error(f"Envío fallido: {message}")

    with btn_col3:
        complement_df = val_df.copy()
        complement_df["payload_json"] = complement_df.apply(
            lambda r: json.dumps({
                "tipo": r.get("TipoDocumento",""),
                "documento": r.get("Documento",""),
                "motivo": r.get("Error",""),
                "origen": r.get("Archivo","")
            }, ensure_ascii=False), axis=1)
        buf = io.StringIO()
        complement_df.to_csv(buf, index=False)
        csv_data = buf.getvalue().encode("utf-8")
        st.download_button("⬇️ Descargar complementaria (CSV)", data=csv_data, file_name="documentos_errados_complementaria.csv", mime="text/csv")

    # Mostrar la tabla original de errores debajo
    st.dataframe(val_df)

# Error de archivo (solo si hay mensajes)
if error_log:
    st.subheader("Error de archivo")
    for err in error_log:
        st.error(err)
