import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(page_title="Validador Excel", layout="centered")
st.title("📊 Validador y Analizador de Archivos Excel")

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

# Layout: misma fila con uploader, Lista Negra y checkbox
col_u, col_l, col_cb = st.columns([2, 3, 1])
with col_u:
    uploaded_files = st.file_uploader("📁 Suba uno o varios archivos Excel", type=["xlsx"], accept_multiple_files=True)
with col_l:
    lista_negra_input = st.text_input("🔎 Lista Negra (ingresa uno o más criterios separados por coma)")
with col_cb:
    include_ref = st.checkbox("Incluir I como Referencia", value=True)

# Contenedores / acumuladores
matches_report = []
duplicates_report = []
threshold_report = []
validation_report = []
error_log = []

# Umbral fijo (sin input) según tu pedido
THRESHOLD_FIXED = 30000

if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=0, dtype=str)
            df.columns = [str(col) for col in df.columns]

            def get_col_by_letter(letter):
                try:
                    idx = ord(letter.upper()) - ord("A")
                    return df.columns[idx]
                except:
                    return None

            # LISTA NEGRA: múltiples criterios (coma-separados), coincidencia exacta sin espacios
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

            # DUPLICADOS: coincidencia completa en C, D, (I opcional), M, R, S
            dup_letters = ["C", "D", "M", "R", "S"]
            if include_ref:
                dup_letters.insert(2, "I")
            dup_cols = []
            missing_dup = []
            for lt in dup_letters:
                c = get_col_by_letter(lt)
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

            # IMPORTES MAYORES A 30,000 (umbral fijo)
            col_M = get_col_by_letter("M")
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
                        c = get_col_by_letter(lt)
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

            # DOCUMENTOS ERRADOS: validación B -> C (DNI 8, CEX 9, RUC 11)
            col_B = get_col_by_letter("B")
            col_C = get_col_by_letter("C")
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

# Mostrar secciones solo si hay datos; subtítulos renombrados

if matches_report:
    matches_df = pd.concat(matches_report, ignore_index=True)
    st.subheader("Lista Negra")
    st.dataframe(matches_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        matches_df.to_excel(writer, index=False, sheet_name="lista_negra")
    st.download_button("⬇️ Descargar Lista Negra", data=buf.getvalue(), file_name="lista_negra.xlsx")

if duplicates_report:
    dup_df = pd.concat(duplicates_report, ignore_index=True)
    st.subheader("Duplicados")
    st.dataframe(dup_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        dup_df.to_excel(writer, index=False, sheet_name="duplicados")
    st.download_button("⬇️ Descargar Duplicados", data=buf.getvalue(), file_name="duplicados.xlsx")

if threshold_report:
    th_df = pd.concat(threshold_report, ignore_index=True)
    st.subheader("Importes mayores a 30,000")
    st.dataframe(th_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        th_df.to_excel(writer, index=False, sheet_name="importes_mayores")
    st.download_button("⬇️ Descargar importes", data=buf.getvalue(), file_name="importes_mayores.xlsx")

if validation_report:
    val_df = pd.concat(validation_report, ignore_index=True)
    st.subheader("Documentos errados")
    st.dataframe(val_df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        val_df.to_excel(writer, index=False, sheet_name="documentos_errados")
    st.download_button("⬇️ Descargar documentos errados", data=buf.getvalue(), file_name="documentos_errados.xlsx")

if error_log:
    st.subheader("Error de archivo")
    for err in error_log:
        st.error(err)
