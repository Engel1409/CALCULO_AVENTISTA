
import streamlit as st
import pandas as pd
import re
import io
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(page_title="Validación de Documentos", layout="wide")
st.title("📊 Validación de Documentos y Resumen de Archivos")

# Subir archivos
archivos = st.file_uploader("Sube tus archivos Excel", type=["xlsx"], accept_multiple_files=True)

# Botón para procesar
if st.button("Procesar archivos") and archivos:
    no_validos = []
    resumen = []

    def validar_documento(row):
        tipo = str(row.get("Tipo de Documento", "")).strip().upper()
        num = str(row.get("Número de Documento", "")).strip()
        if tipo == "DNI":
            return "DNI válido" if num.isdigit() and len(num) == 8 else "DNI inválido"
        return "No es DNI"

    for archivo in archivos:
        nombre_archivo = archivo.name
        st.write(f"Procesando: {nombre_archivo}...")
        df = pd.read_excel(archivo, dtype={"Número de Documento": str})
        df.columns = df.columns.str.strip()
        df = df.dropna(how="all")
        df["fila_en_excel"] = df.index + 2

        if df.empty:
            resumen.append({"Archivo": nombre_archivo, "Poliza": "no declara"})
            continue

        for col in ["Tipo de Documento", "Número de Documento", "Capital Asegurado", "Prima"]:
            if col not in df.columns:
                df[col] = pd.NA

        df["validación documento"] = df.apply(validar_documento, axis=1)

        df_no_validos = df[df["validación documento"] == "No es DNI"].copy()
        df_no_validos["archivo_origen"] = nombre_archivo

        columnas_finales = [
            "Tipo de Documento", "Número de Documento", "Nombre Completo",
            "validación documento", "archivo_origen", "fila_en_excel"
        ]
        for col in columnas_finales:
            if col not in df_no_validos.columns:
                df_no_validos[col] = pd.NA
        df_no_validos = df_no_validos[columnas_finales]

        df_no_validos = df_no_validos[
            df_no_validos["Número de Documento"].notna() &
            df_no_validos["Número de Documento"].astype(str).str.strip().ne("") &
            df_no_validos["Nombre Completo"].notna() &
            df_no_validos["Nombre Completo"].astype(str).str.strip().ne("")
        ]

        if not df_no_validos.empty:
            no_validos.append(df_no_validos)

        ultima_es_subtotal = df.iloc[-1].astype(str).str.contains('TOTAL', case=False, na=False).any() if len(df) >= 1 else False

        if ultima_es_subtotal and len(df) > 1:
            ultima_fila = df.iloc[-1]
            df_sin_ultima = df.iloc[:-1].copy()
            sub_capital = ultima_fila.get("Capital Asegurado", "no declara")
            sub_prima = ultima_fila.get("Prima", "no declara")
        else:
            df_sin_ultima = df.copy()
            sub_capital = "no declara"
            sub_prima = "no declara"

        total_capital_num = df_sin_ultima["Capital Asegurado"].sum(min_count=1) if pd.api.types.is_numeric_dtype(df_sin_ultima["Capital Asegurado"]) else pd.NA
        s = (df_sin_ultima["Prima"].astype(str)
                            .str.replace('\u00A0', '', regex=False)
                            .str.replace('\u202F', '', regex=False)
                            .str.replace(' ', '', regex=False)
                            .str.replace('S/', '', regex=False)
                            .str.replace('s/', '', regex=False)
                            .str.replace('.', '', regex=False)
                            .str.replace(',', '.', regex=False))
        total_prima_num = pd.to_numeric(s, errors="coerce").sum(min_count=1)

        match = re.search(r'\d{10,}', nombre_archivo)
        poliza = match.group(0) if match else "no declara"

        resumen.append({
            "Archivo": nombre_archivo,
            "Poliza": poliza,
            "Cantidad_registros": len(df_sin_ultima),
            "Total_capital": total_capital_num if pd.notna(total_capital_num) else "no declara",
            "Total_prima": total_prima_num if pd.notna(total_prima_num) else "no declara",
            "Total_origen_col_H": sub_capital,
            "Total_origen_col_J": sub_prima
        })

    df_no_validos_final = pd.concat(no_validos, ignore_index=True) if no_validos else pd.DataFrame()
    df_resumen = pd.DataFrame(resumen)

    # ✅ Vista previa
    st.subheader("Vista previa de datos")
    st.write("**No válidos:**")
    st.dataframe(df_no_validos_final if not df_no_validos_final.empty else pd.DataFrame({"mensaje": ["no declara"]}))
    st.write("**Totales por archivo:**")
    st.dataframe(df_resumen)

    # Exportar a Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if df_no_validos_final.empty:
            pd.DataFrame({"mensaje": ["no declara"]}).to_excel(writer, sheet_name="No válidos", index=False)
        else:
            df_no_validos_final.to_excel(writer, sheet_name="No válidos", index=False)
        df_resumen.to_excel(writer, sheet_name="Totales por archivo", index=False)

    st.success("✅ Proceso completado.")
    st.download_button(
        label="📥 Descargar resultado",
        data=output.getvalue(),
        file_name="Resumen_Validacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
