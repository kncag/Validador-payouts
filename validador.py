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

# Layout centrado para uploader y controles
col_left, col_center, col_right = st.columns([1, 2, 1])
with col_center:
    uploaded_files = st.file_uploader("📁 Suba uno o varios archivos Excel", type=["xlsx"], accept_multiple_files=True)

# Preparar contenedores de salida para que siempre existan en la UI
duplicates_container = st.container()
threshold_container = st.container()
validation_container = st.container()
matches_container = st.container()
errors_container = st.container()

if uploaded_files:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        search_term = st.text_input("🔍 Ingrese texto o número a buscar (coincidencia exacta, sin espacios)")
        threshold = st.number_input("⚙️ Umbral para columna M (ej. 30000)", min_value=0, value=30000)

    error_log = []
    duplicates_report = []
    threshold_report = []
    validation_report = []
    matches_report = []

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

            # Subtítulo 1: búsqueda en todo el archivo
            norm = df.applymap(lambda x: normalize_text(x))
            if search_term:
                target = normalize_text(search_term)
                found_mask = norm.isin([target])
                match_rows = df[found_mask.any(axis=1)].copy()
            else:
                match_rows = pd.DataFrame()  # vacío si no se buscó

            if not match_rows.empty:
                match_rows["Archivo"] = file.name
                matches_report.append(match_rows)

            # Subtítulo 2: duplicados completos en C, D, I, M, R, S
            dup_letters = ["C", "D", "I", "M", "R", "S"]
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

            # Subtítulo 3: threshold en M y extracción B,C,D,L,M (parseando M según regla . decimal, , miles)
            col_M = get_col_by_letter("M")
            if col_M is None:
                error_log.append(f"❌ Columna M no encontrada en {file.name}")
            else:
                try:
                    df["_M_num"] = df[col_M].apply(parse_number)
                    filtered = df[df["_M_num"] >= threshold]
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
                    error_log.append(f"❌ Error procesando threshold en {file.name}: {e}")

            # Subtítulo 4: validación B -> C (DNI 8, CEX 9, RUC 11)
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

    # Sección Errores (si no hay, igualmente se muestra)
    with errors_container:
        st.subheader("🚨 Errores detectados")
        if error_log:
            for err in error_log:
                st.error(err)
        else:
            st.info("No se detectaron errores de procesamiento.")

    # Sección Coincidencias (Subtítulo 1)
    with matches_container:
        st.subheader("📌 Coincidencias de búsqueda")
        if matches_report:
            matches_df = pd.concat(matches_report, ignore_index=True)
            st.dataframe(matches_df)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                matches_df.to_excel(writer, index=False, sheet_name="coincidencias")
            st.download_button("⬇️ Descargar coincidencias", data=buf.getvalue(), file_name="coincidencias.xlsx")
        else:
            st.info("No se realizaron búsquedas o no se encontraron coincidencias.")

    # Sección Duplicados (Subtítulo 2)
    with duplicates_container:
        st.subheader("📋 Duplicados detectados (C, D, I, M, R, S)")
        if duplicates_report:
            dup_df = pd.concat(duplicates_report, ignore_index=True)
            st.dataframe(dup_df)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                dup_df.to_excel(writer, index=False, sheet_name="duplicados")
            st.download_button("⬇️ Descargar duplicados", data=buf.getvalue(), file_name="duplicados.xlsx")
        else:
            st.info("No se encontraron duplicados completos en C, D, I, M, R, S.")

    # Sección Threshold (Subtítulo 3)
    with threshold_container:
        st.subheader("📈 Filas con M >= threshold")
        if threshold_report:
            th_df = pd.concat(threshold_report, ignore_index=True)
            st.dataframe(th_df)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                th_df.to_excel(writer, index=False, sheet_name="filtrados_threshold")
            st.download_button("⬇️ Descargar filtrados por threshold", data=buf.getvalue(), file_name="filtrados_threshold.xlsx")
        else:
            st.info("No se encontraron filas que cumplan el threshold en las columnas M procesadas.")

    # Sección Validaciones B/C (Subtítulo 4)
    with validation_container:
        st.subheader("🧪 Validaciones de formato B/C (solo errores)")
        if validation_report:
            val_df = pd.concat(validation_report, ignore_index=True)
            st.dataframe(val_df)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                val_df.to_excel(writer, index=False, sheet_name="errores_validacion")
            st.download_button("⬇️ Descargar errores de validación", data=buf.getvalue(), file_name="errores_validacion.xlsx")
        else:
            st.info("No se detectaron errores de validación B/C.")
else:
    # No hay archivos subidos: mostrar todas las secciones vacías
    with errors_container:
        st.subheader("🚨 Errores detectados")
        st.info("No se detectaron errores de procesamiento.")
    with matches_container:
        st.subheader("📌 Coincidencias de búsqueda")
        st.info("Sube archivos y realiza una búsqueda para ver coincidencias.")
    with duplicates_container:
        st.subheader("📋 Duplicados detectados (C, D, I, M, R, S)")
        st.info("Sube archivos para detectar duplicados completos en C, D, I, M, R, S.")
    with threshold_container:
        st.subheader("📈 Filas con M >= threshold")
        st.info("Sube archivos para evaluar el threshold en columna M.")
    with validation_container:
        st.subheader("🧪 Validaciones de formato B/C (solo errores)")
        st.info("Sube archivos para ejecutar las validaciones B/C.")
