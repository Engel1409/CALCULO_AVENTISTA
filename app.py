import streamlit as st
import pandas as pd
import re
import io
import warnings
from datetime import datetime
from openpyxl.styles import PatternFill, Font
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(page_title="Validación de Documentos", layout="wide")
st.title("📊 VALIDACION ADVENTISTAS📊")

# ------------------ Selector de zona (afecta NETA) ------------------
zona = st.selectbox("Selecciona la zona", options=["Sur", "Norte"], index=0)
NETA = 0.00038 if zona == "Sur" else 0.00036
V_D_E = 0.03
V_IGV = 0.18

# ------------------ Lista de usuarios ------------------
usuarios = ["Engel", "Carlos", "Rosa", "Claudia", "Administrador"]
usuario_seleccionado = st.selectbox("Selecciona tu usuario:", usuarios)

# Fecha del reporte
fecha_reporte = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
st.write(f"📅 **Fecha del reporte:** {fecha_reporte}")

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

        df = pd.read_excel(archivo, dtype={"Número de Documento": str})
        df.columns = df.columns.str.strip()
        df = df.dropna(how="all")
        df["fila_en_excel"] = df.index + 2

        if df.empty:
            resumen.append({"Archivo": nombre_archivo, "Poliza": "no declara"})
            continue

        # Asegurar columnas obligatorias
        for col in ["Tipo de Documento", "Número de Documento", "Capital Asegurado", "Prima"]:
            if col not in df.columns:
                df[col] = pd.NA
        if "Nombre Completo" not in df.columns:
            df["Nombre Completo"] = pd.NA

        # Validación DNI
        df["validación documento"] = df.apply(validar_documento, axis=1)

        # Filtrar no válidos
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

        # Detectar si última fila es TOTAL
        ultima_es_subtotal = df.iloc[-1].astype(str).str.contains("TOTAL", case=False, na=False).any()

        if ultima_es_subtotal and len(df) > 1:
            ultima_fila = df.iloc[-1]
            df_sin_ultima = df.iloc[:-1].copy()
            sub_capital = ultima_fila.get("Capital Asegurado", "no declara")
            sub_prima = ultima_fila.get("Prima", "no declara")
        else:
            df_sin_ultima = df.copy()
            sub_capital = "no declara"
            sub_prima = "no declara"

        # Totales existentes
        total_capital_num = df_sin_ultima["Capital Asegurado"].sum(min_count=1)

        s = (df_sin_ultima["Prima"].astype(str)
             .str.replace('\u00A0', '', regex=False)
             .str.replace('\u202F', '', regex=False)
             .str.replace(' ', '', regex=False)
             .str.replace('S/', '', regex=False)
             .str.replace('s/', '', regex=False)
             .str.replace('.', '', regex=False)
             .str.replace(',', '.', regex=False))

        total_prima_num = pd.to_numeric(s, errors="coerce").sum(min_count=1)

        # ---------- Cálculos ----------
        capital_num = pd.to_numeric(df_sin_ultima["Capital Asegurado"], errors="coerce")

        prima_neta_reg = capital_num * NETA
        d_e_reg = prima_neta_reg * V_D_E
        igv_reg = (prima_neta_reg + d_e_reg) * V_IGV
        total_reg = prima_neta_reg + d_e_reg + igv_reg

        def red2(x):
            return float(round(x, 2)) if pd.notna(x) else "no declara"

        suma_prima_neta = red2(prima_neta_reg.sum(min_count=1))
        suma_d_e = red2(d_e_reg.sum(min_count=1))
        suma_igv = red2(igv_reg.sum(min_count=1))
        suma_total = red2(total_reg.sum(min_count=1))

        # Extraer póliza
        match = re.search(r'\d{10,}', nombre_archivo)
        poliza = match.group(0) if match else "no declara"

        # ---------- RESUMEN ----------
        resumen.append({
            "Archivo": nombre_archivo,
            "Poliza": poliza,
            "Usuario": usuario_seleccionado,
            "Zona": zona,
            "Fecha_reporte": fecha_reporte,
            "Cantidad_registros": len(df_sin_ultima),
            "Total_capital": total_capital_num,
            "Total_origen_col_H": sub_capital,
            "Total_origen_col_J": sub_prima,
            "Suma_prima_neta": suma_prima_neta,
            "Suma_D_E": suma_d_e,
            "Suma_IGV": suma_igv,
            "Suma_TOTAL": suma_total
        })

    df_no_validos_final = pd.concat(no_validos, ignore_index=True) if no_validos else pd.DataFrame()
    df_resumen = pd.DataFrame(resumen)

    # Orden columnas
    orden_cols = [
        "Archivo", "Poliza", "Usuario", "Zona", "Fecha_reporte",
        "Cantidad_registros", "Total_capital",
        "Total_origen_col_H", "Total_origen_col_J",
        "Suma_prima_neta", "Suma_D_E", "Suma_IGV", "Suma_TOTAL"
    ]
    df_resumen = df_resumen[orden_cols]

    # Vista previa
    st.subheader("Vista previa de datos")
    st.write("**Totales por archivo:**")
    st.dataframe(df_resumen)
    st.write("**No válidos:**")
    st.dataframe(df_no_validos_final)

    # Exportar a Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # PRIMERO: Totales por archivo
        df_resumen.to_excel(writer, sheet_name="Totales por archivo", index=False)

        # SEGUNDO: No válidos
        df_no_validos_final.to_excel(writer, sheet_name="No válidos", index=False)

        # ----------- COLOR A LAS CABECERAS -----------
        wb = writer.book
        fill = PatternFill(start_color="D53032", end_color="D53032", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        hojas = ["Totales por archivo", "No válidos"]

        for hoja in hojas:
            ws = wb[hoja]
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font_white

    st.success("✅ Proceso completado.")
    st.download_button(
        label="📥 Descargar reporte final",
        data=output.getvalue(),
        file_name="Resumen_Validacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
