# validador_typed.py
from __future__ import annotations
import re
import io
from dataclasses import dataclass
from typing import List, Optional, Tuple, Dict, Any

import pandas as pd
import numpy as np
import requests
import streamlit as st

# ---------- Configuración ----------
ENDPOINT = "https://q6caqnpy09.execute-api.us-east-1.amazonaws.com/OPS/kpayout/v1/payout_process/reject_invoices_batch"
THRESHOLD_FIXED = 30000

OUT_COLUMNS = [
    "dni/cex",
    "nombre",
    "importe",
    "Referencia",
    "Estado",
    "Codigo de Rechazo",
    "Descripcion de Rechazo",
]

SUBSET_COLUMNS = ["Referencia", "Estado", "Codigo de Rechazo", "Descripcion de Rechazo"]
DEFAULT_ESTADO = "rechazada"


# ---------- Dataclasses para estructuras claras ----------
@dataclass
class FileProcessingResult:
    file_name: str
    df_original: pd.DataFrame


@dataclass
class ValidationRecord:
    tipo_documento: str
    documento: str
    error: str
    origen: str


@dataclass
class RejectionRow:
    dni_cex: str
    nombre: str
    importe: Any
    referencia: str
    estado: str
    codigo_rechazo: str
    descripcion_rechazo: str

    def as_dict(self) -> Dict[str, Any]:
        return {
            "dni/cex": self.dni_cex,
            "nombre": self.nombre,
            "importe": self.importe,
            "Referencia": self.referencia,
            "Estado": self.estado,
            "Codigo de Rechazo": self.codigo_rechazo,
            "Descripcion de Rechazo": self.descripcion_rechazo,
        }


# ---------- Utilidades con tipos ----------
def normalize_text(val: Any) -> str:
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", "", str(val)).strip().lower()


def safe_str(val: Any) -> str:
    if pd.isna(val):
        return ""
    s = str(val)
    return re.sub(r"\.0+$", "", s).strip()


def parse_number(val: Any) -> float:
    try:
        if val is None:
            return float("nan")
        s = str(val).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return float("nan")
        s = re.sub(r"\s+", "", s)
        if "." in s:
            s = s.replace(",", "")
            if s.count(".") > 1:
                parts = s.split(".")
                s = "".join(parts[:-1]) + "." + parts[-1]
            return float(s)
        return float(s.replace(",", ""))
    except Exception:
        return float("nan")


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Rechazos") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def get_col_by_letter(letter: str, df: pd.DataFrame) -> Optional[str]:
    try:
        idx = ord(letter.upper()) - ord("A")
        return df.columns[idx]
    except Exception:
        return None


def find_row_by_document_positional(orig_df: pd.DataFrame, doc_val: Any) -> Optional[pd.Series]:
    if pd.isna(doc_val):
        return None
    target_raw = str(doc_val).strip()

    candidate_cols: List[str] = []
    for letter in ("B", "C"):
        col = get_col_by_letter(letter, orig_df)
        if col is not None and col not in candidate_cols:
            candidate_cols.append(col)
    if not candidate_cols:
        return None

    try:
        as_float = float(target_raw)
        as_int = str(int(as_float))
    except Exception:
        as_int = None
    target_digits = re.sub(r"\D", "", target_raw)

    for col in candidate_cols:
        series = orig_df[col].astype(str).apply(lambda x: str(x).strip())
        # exact match
        mask = series == target_raw
        if mask.any():
            return orig_df.loc[mask].iloc[0]
        # int-like match
        if as_int and (series == as_int).any():
            return orig_df.loc[series == as_int].iloc[0]
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


# ---------- Comunicación con endpoint ----------
def post_rejections_excel(excel_bytes: bytes) -> Tuple[int, str]:
    files = {
        "edt": ("rechazos.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    }
    resp = requests.post(ENDPOINT, files=files)
    return resp.status_code, resp.text


def send_rejections_handler(df_out: pd.DataFrame, ui_callable=None) -> Tuple[bool, str]:
    if list(df_out.columns) != OUT_COLUMNS:
        msg = f"Encabezados inválidos. Se requieren: {OUT_COLUMNS}"
        if ui_callable:
            ui_callable("error", msg)
        return False, msg
    payload = df_out[SUBSET_COLUMNS]
    try:
        excel_bytes = df_to_xlsx_bytes(payload)
    except Exception as e:
        msg = f"Error generando Excel: {e}"
        if ui_callable:
            ui_callable("error", msg)
        return False, msg
    try:
        status, text = post_rejections_excel(excel_bytes)
    except Exception as e:
        msg = f"Error realizando POST: {e}"
        if ui_callable:
            ui_callable("error", msg)
        return False, msg
    msg = f"{status}: {text}"
    if ui_callable:
        if 200 <= status < 300:
            ui_callable("success", msg)
        else:
            ui_callable("error", msg)
    return (200 <= status < 300), msg


# ---------- UI helper modular ----------
def render_section_actions(df_preview: pd.DataFrame, section_name: str, code: str, description: str, primary_label: str = "RECH-POSTMAN"):
    st.markdown(f"**Preview (exactamente lo que se enviará al endpoint - {section_name})**")
    st.dataframe(df_preview)
    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button(f"{primary_label} - {section_name}"):
            ok, message = send_rejections_handler(df_preview, ui_callable=lambda lvl, m: getattr(st, lvl)(m))
            if ok:
                st.info("Envío completado correctamente.")
            else:
                st.error(f"Envío fallido: {message}")
    with col_b:
        xlsx_bytes = df_to_xlsx_bytes(df_preview)
        st.download_button("⬇️ Descargar", data=xlsx_bytes, file_name=f"{section_name.lower().replace(' ','_')}_rechazos.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- Flujo principal (modular) ----------
def process_uploaded_files(files: List[Any], lista_negra_input: str, include_ref: bool) -> Tuple[
    List[pd.DataFrame], List[pd.DataFrame], List[pd.DataFrame], List[pd.DataFrame], List[pd.DataFrame], List[str]
]:
    """
    Devuelve: matches_report, duplicates_report, threshold_report, validation_report, originals, error_log
    """
    matches_report: List[pd.DataFrame] = []
    duplicates_report: List[pd.DataFrame] = []
    threshold_report: List[pd.DataFrame] = []
    validation_report: List[pd.DataFrame] = []
    originals: List[pd.DataFrame] = []
    error_log: List[str] = []

    for file in files:
        try:
            df = pd.read_excel(file, header=0, dtype=str)
            df.columns = [str(c) for c in df.columns]
            originals.append(df.copy())

            # Lista negra
            if lista_negra_input:
                criteria = [normalize_text(x) for x in lista_negra_input.split(",") if x.strip()]
                if criteria:
                    norm = df.applymap(normalize_text)
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

            # Duplicados
            dup_letters = ["C", "D", "M", "R", "S"]
            if include_ref:
                dup_letters.insert(2, "I")
            dup_cols: List[str] = []
            missing_dup: List[str] = []
            for lt in dup_letters:
                c = get_col_by_letter(lt, df)
                if c is None:
                    missing_dup.append(lt)
                else:
                    dup_cols.append(c)
            if missing_dup:
                error_log.append(f"Columnas duplicados faltantes {missing_dup} en {file.name}")
            else:
                subset = df[dup_cols].fillna("").astype(str).applymap(lambda x: x.strip())
                dup_mask = subset.duplicated(keep=False)
                if dup_mask.any():
                    dups = df.loc[dup_mask].copy()
                    dups["Archivo"] = file.name
                    dups["Columnas comprobadas"] = ",".join(dup_letters)
                    duplicates_report.append(dups)

            # Importes mayores
            col_M = get_col_by_letter("M", df)
            if col_M is None:
                error_log.append(f"Columna M no encontrada en {file.name}")
            else:
                df["_M_num"] = df[col_M].apply(parse_number)
                filtered = df[df["_M_num"] >= THRESHOLD_FIXED]
                extract_letters = ["B", "C", "D", "L", "M"]
                extract_cols: List[str] = []
                missing_ex: List[str] = []
                for lt in extract_letters:
                    c = get_col_by_letter(lt, df)
                    if c is None:
                        missing_ex.append(lt)
                    else:
                        extract_cols.append(c)
                if missing_ex:
                    error_log.append(f"Columnas faltantes para extracción {missing_ex} en {file.name}")
                elif not filtered.empty:
                    out = filtered[extract_cols].copy()
                    out["Archivo"] = file.name
                    threshold_report.append(out)

            # Documentos errados (B->C posicional)
            col_B = get_col_by_letter("B", df)
            col_C = get_col_by_letter("C", df)
            if col_B is None or col_C is None:
                error_log.append(f"Columnas B o C faltantes en {file.name}")
            else:
                tipos = df[col_B].astype(str).apply(lambda x: safe_str(x).upper())
                numeros = df[col_C].astype(str).apply(lambda x: safe_str(x))
                def validate_pair(tipo: str, valor_raw: str) -> Optional[str]:
                    if valor_raw == "" or valor_raw.lower() in {"nan", "none"}:
                        if tipo in {"DNI", "CEX", "RUC"}:
                            return f"{tipo} inválido - vacío"
                        return None
                    if tipo == "DNI":
                        if not valor_raw.isdigit() or len(valor_raw) != 8:
                            return "DNI inválido"
                        return None
                    elif tipo == "CEX":
                        vr = valor_raw.zfill(9)
                        if not vr.isdigit() or len(vr) != 9:
                            return "CEX inválido"
                        return None
                    elif tipo == "RUC":
                        vr = valor_raw.zfill(11)
                        if not vr.isdigit() or len(vr) != 11:
                            return "RUC inválido"
                        return None
                    return None
                errs = [validate_pair(t, n) for t, n in zip(tipos, numeros)]
                report_df = pd.DataFrame({"TipoDocumento": tipos.values, "Documento": numeros.values, "Error": errs})
                report_df = report_df[report_df["Error"].notna()].copy()
                if not report_df.empty:
                    report_df["Archivo"] = file.name
                    validation_report.append(report_df)

        except Exception as exc:
            error_log.append(f"Error procesando {file.name}: {exc}")

    return matches_report, duplicates_report, threshold_report, validation_report, originals, error_log


# ---------- UI principal modular ----------
def main():
    st.set_page_config(page_title="Validador Excel", layout="centered")
    st.title("📊 Validador y Analizador de Archivos Excel")

    # uploader en su propia fila con badge
    st.subheader("Subir archivos")
    st.markdown("Arrastra o selecciona archivos Excel (.xlsx).")
    uploaded_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True, key="main_uploader")
    uploaded_count = len(uploaded_files) if uploaded_files else 0
    st.metric(label="Archivos cargados", value=str(uploaded_count))

    # inputs secundarios en fila separada
    col1, col2, _ = st.columns([3, 2, 1])
    with col1:
        lista_negra_input = st.text_input("🔎 Lista Negra (criterios separados por coma)")
    with col2:
        include_ref = st.checkbox("Incluir I como Referencia", value=True)

    # procesamiento
    matches_report, duplicates_report, threshold_report, validation_report, originals, error_log = process_uploaded_files(
        uploaded_files or [], lista_negra_input, include_ref
    )

    # Lista Negra
    if matches_report:
        matches_df = pd.concat(matches_report, ignore_index=True)
        st.subheader("Lista Negra")
        st.dataframe(matches_df)
        rows: List[RejectionRow] = []
        for _, r in matches_df.iterrows():
            dni = safe_str(r.get("DOCUMENTO") or r.get("B") or "")
            nombre = safe_str(r.get("NOMBRE") or r.get("D") or "")
            referencia = safe_str(r.get("REFERENCIA") or r.get("I") or "")
            importe_val = parse_number(r.get("MONTO") or r.get("M") or "")
            rows.append(RejectionRow(
                dni_cex=dni,
                nombre=nombre,
                importe=importe_val if not np.isnan(importe_val) else "",
                referencia=referencia,
                estado=DEFAULT_ESTADO,
                codigo_rechazo="R002",
                descripcion_rechazo="CUENTA INVALIDA"
            ))
        df_out_ln = pd.DataFrame([r.as_dict() for r in rows], columns=OUT_COLUMNS)
        render_section_actions(df_out_ln, section_name="Lista Negra", code="R002", description="CUENTA INVALIDA")

    # Duplicados
    if duplicates_report:
        dup_df = pd.concat(duplicates_report, ignore_index=True)
        st.subheader("Duplicados")
        st.dataframe(dup_df)
        rows_dup: List[RejectionRow] = []
        for _, r in dup_df.iterrows():
            dni = safe_str(r.get("DOCUMENTO") or r.get("B") or "")
            nombre = safe_str(r.get("NOMBRE") or r.get("D") or "")
            referencia = safe_str(r.get("REFERENCIA") or r.get("I") or "")
            importe_val = parse_number(r.get("MONTO") or r.get("M") or "")
            rows_dup.append(RejectionRow(
                dni_cex=dni,
                nombre=nombre,
                importe=importe_val if not np.isnan(importe_val) else "",
                referencia=referencia,
                estado=DEFAULT_ESTADO,
                codigo_rechazo="R019",
                descripcion_rechazo="ID DE TRANSACCIÓN DUPLICADA"
            ))
        df_out_dup = pd.DataFrame([r.as_dict() for r in rows_dup], columns=OUT_COLUMNS)
        render_section_actions(df_out_dup, section_name="Duplicados", code="R019", description="ID DE TRANSACCIÓN DUPLICADA")

    # Importes mayores
    if threshold_report:
        th_df = pd.concat(threshold_report, ignore_index=True)
        st.subheader("Importes mayores a 30,000")
        st.dataframe(th_df)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            th_df.to_excel(writer, index=False, sheet_name="importes_mayores")
        st.download_button("⬇️ Descargar importes", data=buf.getvalue(), file_name="importes_mayores.xlsx")

    # Documentos errados
    if validation_report:
        val_df = pd.concat(validation_report, ignore_index=True)
        st.subheader("Documentos errados")
        st.dataframe(val_df)

        rows_err: List[RejectionRow] = []
        unmapped: List[str] = []
        for _, err_row in val_df.iterrows():
            doc_val = err_row.get("Documento", "")
            mapped = False
            for orig in originals:
                candidate = find_row_by_document_positional(orig, doc_val)
                if candidate is not None:
                    dni = safe_str(candidate.get(get_col_by_letter("B", orig))) if get_col_by_letter("B", orig) else ""
                    nombre = safe_str(candidate.get(get_col_by_letter("D", orig))) if get_col_by_letter("D", orig) else ""
                    referencia = safe_str(candidate.get(get_col_by_letter("I", orig))) if get_col_by_letter("I", orig) else ""
                    monto_val = parse_number(candidate.get(get_col_by_letter("M", orig))) if get_col_by_letter("M", orig) else float("nan")
                    rows_err.append(RejectionRow(
                        dni_cex=dni,
                        nombre=nombre,
                        importe=monto_val if not np.isnan(monto_val) else "",
                        referencia=referencia,
                        estado=DEFAULT_ESTADO,
                        codigo_rechazo="R001",
                        descripcion_rechazo="DOCUMENTO ERRADO"
                    ))
                    mapped = True
                    break
            if not mapped:
                unmapped.append(str(doc_val))
                rows_err.append(RejectionRow(
                    dni_cex=str(doc_val),
                    nombre="",
                    importe="",
                    referencia="",
                    estado=DEFAULT_ESTADO,
                    codigo_rechazo="R001",
                    descripcion_rechazo="DOCUMENTO ERRADO"
                ))
        df_out_err = pd.DataFrame([r.as_dict() for r in rows_err], columns=OUT_COLUMNS)
        if unmapped:
            st.warning(f"No se pudo mapear Referencia/nombre/importe para {len(unmapped)} documento(s). Ejemplos: {unmapped[:5]}")
        render_section_actions(df_out_err, section_name="Documentos errados", code="R001", description="DOCUMENTO ERRADO")

    # Errores
    if error_log:
        st.subheader("Error de archivo")
        for e in error_log:
            st.error(e)


if __name__ == "__main__":
    main()
