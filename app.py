import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

DB_FILE = "data.db"

# 1. BASE DE DATOS ULTRA LIMPIA
def inicializar_bd():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS despachos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT, hora_despacho TEXT, hora_llegada TEXT,
            conductor TEXT, orden_trabajo TEXT, cliente TEXT,
            cantidad_etiquetas INTEGER, cantidad_cajas INTEGER,
            factura_guia TEXT, recepcionado_por TEXT
        );
    """)
    conn.commit()
    conn.close()

inicializar_bd()

# 2. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="Control de Despachos", layout="centered")

# --- 🖼️ AQUÍ SE COLOCA TU LOGO ---
# Opción A: Si tu logo está en la misma carpeta que este archivo app.py:
# st.image("logo.png", width=150)
#
# Opción B: Si tienes el link de internet de tu imagen (reemplaza esta URL de ejemplo):
st.image("https://flaticon.com", width=100)

st.title("🚚 Registro de Despachos")

# 3. FORMULARIO DE INGRESO
with st.form(key="formulario", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    fecha = col1.date_input("Fecha")
    hora_despacho = col2.time_input("Hora Despacho")
    hora_llegada = col3.time_input("Hora Llegada")
        
    conductor = st.selectbox("Conductor / Transporte", ["ELIZABETH", "ELIZABETH - COLCHAGUA", "CHILEXPRESS", "STARKEN", "FEDEX", "BLUE EXPRESS", "OTRO EXTERNO"])
    orden_trabajo = st.text_input("Orden de Trabajo (OT)")
    cliente = st.selectbox("Cliente (Remitente)", ["VIÑEDOS DE AGUIRRE", "VIÑA CASADONOSO", "VIÑA AROMO", "VIÑA CASA SOLIS", "VIÑA CONCHA Y TORO", "VIÑA LUIS FELIPE EDWARDS", "TRANSPORTES COLCHAGUA"])
    facturas_guias_texto = st.text_area("Facturas o Guías (Una por línea)")
        
    col4, col5 = st.columns(2)
    cant_etiquetas = col4.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    cant_cajas = col5.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    recepcionado_por = st.selectbox("Recepcionado Por", ["CAMILA VILLASANA", "OLGA GONZALEZ", "JORGE DIAZ", "JAVIERA RAMIREZ", "ENCARGADO COLCHAGUA"])
    
    boton_guardar = st.form_submit_button(label="💾 Registrar Despacho")

# 4. LÓGICA DE GUARDADO
if boton_guardar:
    lista_documentos = [linea.strip() for linea in facturas_guias_texto.split("\n") if linea.strip()]
    if not lista_documentos:
        st.error("⚠️ Por favor, ingresa al menos una factura o guía.")
    else:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        for doc in lista_documentos:
            cursor.execute("""
                INSERT INTO despachos (fecha, hora_despacho, hora_llegada, conductor, orden_trabajo, cliente, cantidad_etiquetas, cantidad_cajas, factura_guia, recepcionado_por)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (fecha.strftime("%Y-%m-%d"), hora_despacho.strftime("%H:%M"), hora_llegada.strftime("%H:%M"), conductor, orden_trabajo, cliente, int(cant_etiquetas), int(cant_cajas), doc, recepcionado_por))
        conn.commit()
        conn.close()
        st.success("✅ ¡Despacho registrado!")
        st.rerun()

# 5. DESCARGA DIRECTA A EXCEL
if os.path.exists(DB_FILE):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT fecha, hora_despacho, hora_llegada, conductor, orden_trabajo, cliente, cantidad_etiquetas, cantidad_cajas, factura_guia, recepcionado_por FROM despachos ORDER BY id DESC;", conn)
    conn.close()

    if not df.empty:
        st.markdown("---")
        
        df.columns = ["Fecha", "Hora Despacho", "Hora Llegada", "Conductor", "OT", "Cliente", "Etiquetas", "Cajas", "Factura / Guía", "Recepcionado Por"]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Despachos', index=False)
        
        st.download_button(
            label="📥 Descargar Reporte Excel",
            data=output.getvalue(),
            file_name=f"Reporte_Despachos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
