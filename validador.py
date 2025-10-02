import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(page_title="Validador Excel", layout="wide")
st.title("📊 Validador y Analizador de Archivos Excel")

def normalize_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", "", str(val)).strip().lower()

def parse_number(val):
    try:
        s = str(val).strip()
        if s == "":
            return np.nan
        s = s.replace(".", "").replace(",", ".")
        return float(s)
    except:
        return np.nan

uploaded_files = st.file_uploader("📁 Suba uno o varios archivos Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    search_term = st.text_input("🔍 Ingrese texto o número a buscar (coincidencia exacta, sin espacios)")
    threshold = st.number_input("⚙️ Umbral para columna M (ej. 30000)", min_value=0, value=30000)

    all_data = []
    error_log = []
    duplicates_report = []
    threshold_report = []
    validation_report = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=0)
            df_original = df.copy()
            df.columns = [str(col) for col in df.columns]

            # Subtítulo 1: búsqueda de coincidencias exactas en todo el archivo
            if search_term:
                norm = df.applymap(lambda x: normalize_text(x))
                target = normalize_text(search_term)
                found_mask = norm.isin([target])
                match_rows = df[found_mask.any(axis=1)]
                if not match_rows.empty:
                    match_rows_display = match_rows.copy()
                    match_rows_display["Archivo"] = file.name
                    st.subheader(f"📌 Coincidencias en archivo: {file.name}")
                    st.dataframe(match_rows_display)
                else:
                    st.info(f"No se encontraron coincidencias en {file.name}.")

            # Preparar lista de columnas por letra (por posición)
            def get_col_by_letter(letter):
                try:
                    idx = ord(letter.upper()) - ord("A")
                    return df.columns[idx]
                except:
                    return None

            # Subtítulo 2: detectar duplicados en columnas M, I, C
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

            # Subtítulo 3: threshold en columna M y extracción B,C,D,L,M
            col_M = get_col_by_letter("M")
            if col_M is None:
                error_log.append(f"❌ Columna M no encontrada en {file.name}")
            else:
                try:
                    df["_M_num"] = df[col_M].apply(parse_number)
                    filtered = df[df["_M_num"] >= threshold]
                    if not filtered.empty:
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
                        else:
                            out = filtered[extract_cols].copy()
                            out["Archivo"] = file.name
                            threshold_report.append(out)
                except Exception as e:
                    error_log.append(f"❌ Error procesando threshold en {file.name}: {e}")

            # Subtítulo 4: validación según B -> C (DNI 8, CEX 9, RUC 11), solo marcar errores
            col_B = get_col_by_letter("B")
            col_C = get_col_by_letter("C")
            if col_B is None or col_C is None:
                error_log.append(f"❌ Columnas B o C faltantes en {file.name}")
            else:
                try:
                    def validate_row(row):
                        tipo = str(row[col_B]).strip().upper()
                        valor_raw = str(row[col_C]).strip()
                        # vacíos se consideran inválidos
                        if valor_raw == "" or pd.isna(row[col_C]):
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

                    df["Error"] = df.apply(validate_row, axis=1)
                    errores = df[df["Error"].notna()].copy()
                    if not errores.empty:
                        errores["Archivo"] = file.name
                        errores["Tipo validación"] = "B/C"
                        validation_report.append(errores)
                except Exception as e:
                    error_log.append(f"❌ Error en validación B/C en {file.name}: {e}")

        except Exception as e:
            error_log.append(f"❌ Error procesando {file.name}: {e}")

    # Mostrar errores de procesamiento
    if error_log:
        st.subheader("🚨 Errores detectados")
        for err in error_log:
            st.error(err)

    # Mostrar y permitir descarga de duplicados
    if duplicates_report:
        st.subheader("📋 Duplicados detectados")
        dup_df = pd.concat(duplicates_report, ignore_index=True)
        st.dataframe(dup_df)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            dup_df.to_excel(writer, index=False, sheet_name="duplicados")
        st.download_button("⬇️ Descargar duplicados", data=buf.getvalue(), file_name="duplicados.xlsx")

    # Mostrar y permitir descarga de threshold
    if threshold_report:
        st.subheader("📈 Filas con M >= threshold")
        th_df = pd.concat(threshold_report, ignore_index=True)
        st.dataframe(th_df)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            th_df.to_excel(writer, index=False, sheet_name="filtrados_threshold")
        st.download_button("⬇️ Descargar filtrados por threshold", data=buf.getvalue(), file_name="filtrados_threshold.xlsx")

    # Mostrar y permitir descarga de validaciones B/C
    if validation_report:
        st.subheader("🧪 Validaciones de formato B/C (solo errores)")
        val_df = pd.concat(validation_report, ignore_index=True)
        display_cols = [c for c in val_df.columns if c in [col_B, col_C, "Error", "Archivo", "Tipo validación"]]
        st.dataframe(val_df[display_cols])
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            val_df.to_excel(writer, index=False, sheet_name="errores_validacion")
        st.download_button("⬇️ Descargar errores de validación", data=buf.getvalue(), file_name="errores_validacion.xlsx")
