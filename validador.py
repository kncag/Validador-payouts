import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(page_title="Validador Excel", layout="wide")
st.title("📊 Validador y Analizador de Archivos Excel")

# Función para normalizar texto
def normalize_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", "", str(val)).strip().lower()

# Función para convertir valores numéricos con miles y decimales regionales
def parse_number(val):
    try:
        val = str(val).replace(".", "").replace(",", ".")
        return float(val)
    except:
        return np.nan

# Cargar archivos
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

            # Normalizar todo para búsqueda
            if search_term:
                found = df.applymap(lambda x: normalize_text(x)).isin([normalize_text(search_term)])
                match_rows = df[found.any(axis=1)]
                st.subheader(f"📌 Coincidencias en archivo: {file.name}")
                st.dataframe(match_rows)

            # Duplicados en columnas M, I, C (por posición si no hay nombre)
            cols_to_check = []
            for col_letter in ["M", "I", "C"]:
                try:
                    idx = ord(col_letter) - ord("A")
                    colname = df.columns[idx]
                    cols_to_check.append(colname)
                except:
                    error_log.append(f"❌ Columna {col_letter} no encontrada en {file.name}")
                    continue

            for col in cols_to_check:
                dups = df[df.duplicated(subset=[col], keep=False)]
                if not dups.empty:
                    dups["Archivo"] = file.name
                    dups["Columna duplicada"] = col
                    duplicates_report.append(dups)

            # Threshold en columna M
            try:
                col_M = df.columns[ord("M") - ord("A")]
                df["M_num"] = df[col_M].apply(parse_number)
                filtered = df[df["M_num"] >= threshold]
                if not filtered.empty:
                    extract_cols = [df.columns[ord(c) - ord("A")] for c in ["B", "C", "D", "L", "M"]]
                    filtered_out = filtered[extract_cols]
                    filtered_out["Archivo"] = file.name
                    threshold_report.append(filtered_out)
            except:
                error_log.append(f"❌ Columna M no encontrada o inválida en {file.name}")

            # Validación de formato según tipo de documento
            try:
                col_B = df.columns[ord("B") - ord("A")]
                col_C = df.columns[ord("C") - ord("A")]
                df["C_str"] = df[col_C].astype(str).str.zfill(11)

                def validate_row(row):
                    tipo = str(row[col_B]).strip().upper()
                    valor = str(row["C_str"])
                    if tipo == "DNI" and len(valor) != 8:
                        return "DNI inválido"
                    elif tipo == "CEX" and len(valor) != 9:
                        return "CEX inválido"
                    elif tipo == "RUC" and len(valor) != 11:
                        return "RUC inválido"
                    return None

                df["Error"] = df.apply(validate_row, axis=1)
                errores = df[df["Error"].notna()]
                if not errores.empty:
                    errores["Archivo"] = file.name
                    validation_report.append(errores)
            except:
                error_log.append(f"❌ Error en validación B/C en {file.name}")

        except Exception as e:
            error_log.append(f"❌ Error procesando {file.name}: {str(e)}")

    # Mostrar errores
    if error_log:
        st.subheader("🚨 Errores detectados")
        for err in error_log:
            st.error(err)

    # Mostrar duplicados
    if duplicates_report:
        st.subheader("📋 Duplicados detectados")
        dup_df = pd.concat(duplicates_report, ignore_index=True)
        st.dataframe(dup_df)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            dup_df.to_excel(writer, index=False)
        st.download_button("⬇️ Descargar duplicados", data=output.getvalue(), file_name="duplicados.xlsx")

    # Mostrar threshold
    if threshold_report:
        st.subheader("📈 Filas con M >= threshold")
        th_df = pd.concat(threshold_report, ignore_index=True)
        st.dataframe(th_df)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            th_df.to_excel(writer, index=False)
        st.download_button("⬇️ Descargar filtrados por threshold", data=output.getvalue(), file_name="filtrados_threshold.xlsx")

    # Mostrar validaciones
    if validation_report:
        st.subheader("🧪 Validaciones de formato B/C")
        val_df = pd.concat(validation_report, ignore_index=True)
        st.dataframe(val_df[[col_B, col_C, "Error", "Archivo"]])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            val_df.to_excel(writer, index=False)
        st.download_button("⬇️ Descargar errores de validación", data=output.getvalue(), file_name="errores_validacion.xlsx")
