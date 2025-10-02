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
    """Interpreta números siguiendo la regla: '.' es separador decimal y ',' separador de miles.
    Ejemplos válidos: '1,234.56' -> 1234.56; '1234.56' -> 1234.56; '1.234' -> 1.234 (punto decimal);
    '1,234' -> 1234 (coma como separador de miles); '1234' -> 1234.
    También maneja espacios y valores vacíos."""
    try:
        if val is None:
            return np.nan
        s = str(val).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return np.nan
        # Eliminar espacios intermedios
        s = re.sub(r"\s+", "", s)
        # Si hay al menos un punto, lo tratamos como decimal.
        # Eliminamos las comas que actúan como separador de miles.
        if "." in s:
            s = s.replace(",", "")  # coma = miles -> eliminar
            # ahora s tiene punto(s). Si hay múltiples puntos, solo el último se considera decimal:
            if s.count(".") > 1:
                # unir los posibles miles con puntos y dejar último como decimal: 1.234.567.89 -> 1234567.89
                parts = s.split(".")
                s = "".join(parts[:-1]) + "." + parts[-1]
            return float(s)
        else:
            # No hay punto. Las comas se consideran separador de miles -> eliminar y parsear entero
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

# Centrar uploader y controles usando columnas
col_left, col_center, col_right = st.columns([1, 2, 1])
with col_center:
    uploaded_files = st.file_uploader("📁 Suba uno o varios archivos Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        search_term = st.text_input("🔍 Ingrese texto o número a buscar (coincidencia exacta, sin espacios)")
        threshold = st.number_input("⚙️ Umbral para columna M (ej. 30000)", min_value=0, value=30000)

    error_log = []
    duplicates_report = []
    threshold_report = []
    validation_report = []

    for file in uploaded_files:
        try:
            # Leer todo como string para preservar ceros iniciales
            df = pd.read_excel(file, header=0, dtype=str)
            df.columns = [str(col) for col in df.columns]

            # Helper: obtener nombre de columna por letra (indexación por posición)
            def get_col_by_letter(letter):
                try:
                    idx = ord(letter.upper()) - ord("A")
                    return df.columns[idx]
                except:
                    return None

            # Subtítulo 1: búsqueda en todo el archivo
            if search_term:
                norm = df.applymap(lambda x: normalize_text(x))
                target = normalize_text(search_term)
                found_mask = norm.isin([target])
                match_rows = df[found_mask.any(axis=1)]
                match_rows_display = match_rows.copy()
                match_rows_display["Archivo"] = file.name
                st.subheader(f"📌 Coincidencias en archivo: {file.name}")
                if not match_rows_display.empty:
                    st.dataframe(match_rows_display)
                else:
                    st.info(f"No se encontraron coincidencias en {file.name}.")

            # Subtítulo 2: duplicados en M, I, C
            for letter in ["M", "I", "C"]:
                colname = get_col_by_letter(letter)
                if colname is None:
                    error_log.append(f"❌ Columna {letter} no encontrada en {file.name}")
                    continue
                try:
                    dups = df[df.duplicated(subset=[colname], keep=False) & df[colname].notna()]
                    if not dups.empty:
                        dups_report = dups.copy()
                        dups_report["Archivo"] = file.name
                        dups_report["Columna duplicada"] = letter
                        duplicates_report.append(dups_report)
                except Exception as e:
                    error_log.append(f"❌ Error detectando duplicados en columna {letter} de {file.name}: {e}")

            # Subtítulo 3: threshold en M y extracción B,C,D,L,M
            col_M = get_col_by_letter("M")
            if col_M is None:
                error_log.append(f"❌ Columna M no encontrada en {file.name}")
            else:
                try:
                    # Aplicar parse_number que interpreta '.' como decimal y ',' como miles
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
                    if filtered is not None and not filtered.empty and not missing_extract:
                        out = filtered[extract_cols].copy()
                        out["Archivo"] = file.name
                        threshold_report.append(out)
                    elif missing_extract:
                        error_log.append(f"❌ Columnas faltantes para extracción {missing_extract} en {file.name}")
                except Exception as e:
                    error_log.append(f"❌ Error procesando threshold en {file.name}: {e}")

            # Subtítulo 4: validación B -> C (DNI 8, CEX 9, RUC 11)
            col_B = get_col_by_letter("B")
            col_C = get_col_by_letter("C")
            if col_B is None or col_C is None:
                error_log.append(f"❌ Columnas B o C faltantes en {file.name}")
            else:
                try:
                    df["_tipo_doc"] = df[col_B].astype(str).apply(lambda x: safe_str_preserve(x).strip().upper())
                    df["_num_doc"] = df[col_C].astype(str).apply(lambda x: safe_str_preserve(x).strip())

                    def validate_row_std(tipo, valor_raw):
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

                    df["_Error_validacion"] = df.apply(lambda r: validate_row_std(r["_tipo_doc"], r["_num_doc"]), axis=1)
                    errores = df[df["_Error_validacion"].notna()].copy()
                    if not errores.empty:
                        report = pd.DataFrame({
                            "TipoDocumento": errores["_tipo_doc"].values,
                            "Documento": errores["_num_doc"].values,
                            "Error": errores["_Error_validacion"].values,
                            "Archivo": file.name
                        })
                        validation_report.append(report)
                except Exception as e:
                    error_log.append(f"❌ Error en validación B/C en {file.name}: {e}")

        except Exception as e:
            error_log.append(f"❌ Error procesando {file.name}: {e}")

    # Mostrar errores de procesamiento
    if error_log:
        st.subheader("🚨 Errores detectados")
        for err in error_log:
            st.error(err)

    # Duplicados
    if duplicates_report:
        st.subheader("📋 Duplicados detectados")
        dup_df = pd.concat(duplicates_report, ignore_index=True)
        st.dataframe(dup_df)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            dup_df.to_excel(writer, index=False, sheet_name="duplicados")
        st.download_button("⬇️ Descargar duplicados", data=buf.getvalue(), file_name="duplicados.xlsx")

    # Threshold
    if threshold_report:
        st.subheader("📈 Filas con M >= threshold")
        th_df = pd.concat(threshold_report, ignore_index=True)
        st.dataframe(th_df)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            th_df.to_excel(writer, index=False, sheet_name="filtrados_threshold")
        st.download_button("⬇️ Descargar filtrados por threshold", data=buf.getvalue(), file_name="filtrados_threshold.xlsx")

    # Validaciones B/C (sección garantizada)
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
