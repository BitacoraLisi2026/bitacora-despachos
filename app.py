import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

# Archivo de la base de datos fija en la nube
DB_FILE = "data.db"

# ==============================================================
# 📋 CONFIGURA AQUÍ TUS DATOS REALES
# ==============================================================
LISTA_CLIENTES = [
    "Selecciona un cliente...", 
    "VIÑEDOS DE AGUIRRE", "VIÑA CASADONOSO", "VIÑA AROMO",
    "VIÑA CASA SOLIS", "VIÑA CONCHA Y TORO", 
    "VIÑA LUIS FELIPE EDWARDS", "TRANSPORTES COLCHAGUA"
]

LISTA_RECEIPCION = [
    "Selecciona quién recibe...", 
    "CAMILA VILLASANA", "OLGA GONZALEZ", "JORGE DIAZ",
    "JAVIERA RAMIREZ", "ENCARGADO COLCHAGUA"
]

LISTA_CONDUCTORES = [
    "Selecciona el conductor o transporte...",
    "ELIZABETH",              
    "ELIZABETH - COLCHAGUA",  
    "CHILEXPRESS", "STARKEN", "FEDEX", "BLUE EXPRESS", "OTRO EXTERNO"
]
# ==============================================================

st.set_page_config(page_title="Bitácora de Despachos", page_icon="🚚", layout="centered")

# --- CONEXIÓN INICIAL A BASE DE DATOS SQL NATIVA ---
def inicializar_bd():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS despachos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            hora_despacho TEXT,
            hora_llegada TEXT,
            conductor TEXT,
            orden_trabajo TEXT,
            cliente TEXT,
            cantidad_etiquetas INTEGER,
            cantidad_cajas INTEGER,
            factura_guia TEXT,
            recepcionado_por TEXT
        );
    """)
    conn.commit()
    conn.close()

inicializar_bd()

# --- AJUSTE DE HORA LOCAL (CHILE) ---
try:
    import zoneinfo
    zona_horaria = zoneinfo.ZoneInfo("America/Santiago") 
    hora_local_dt = datetime.now(zona_horaria)
except Exception:
    hora_local_dt = datetime.now()

fecha_actual = hora_local_dt.date()
hora_actual = hora_local_dt.time()

st.markdown("<h2 style='font-size: 30px; margin-bottom: 0px;'>🚚 Bitácora de Despachos (Nube)</h2>", unsafe_allow_html=True)
st.write("Ingresa los datos para actualizar la base de datos en tiempo real.")

# Formulario básico limpio de kilómetros
with st.form(key="formulario_bitacora", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    with col1: fecha = st.date_input("Fecha", value=fecha_actual)
    with col2: hora_despacho = st.time_input("Hora de Despacho", value=hora_actual)
    with col3: hora_llegada = st.time_input("Hora de Llegada", value=hora_actual)
        
    st.markdown("---")
    col4, col5 = st.columns(2)
    with col4:
        conductor = st.selectbox("Conductor / Empresa de Transporte", options=LISTA_CONDUCTORES)
        orden_trabajo = st.text_input("Orden de Trabajo")
    with col5:
        cliente = st.selectbox("Cliente (Remitente)", options=LISTA_CLIENTES)
        facturas_guias_texto = st.text_area("Facturas o Guías (Una por línea)")
        
    st.markdown("---")
    col6, col7 = st.columns(2)
    with col6: cant_etiquetas = st.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    with col7: cant_cajas = st.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    recepcionado_por = st.selectbox("Recepcionado Por", options=LISTA_RECEIPCION)
    
    boton_guardar = st.form_submit_button(label="💾 Registrar Despacho")

# Lógica de Guardado en SQL Nativo
if boton_guardar:
    lista_documentos = [linea.strip() for linea in facturas_guias_texto.split("\n") if linea.strip()]
    
    if conductor == "Selecciona el conductor o transporte...":
        st.error("⚠️ Por favor, selecciona un Conductor o Empresa de transporte válida.")
    elif cliente == "Selecciona un cliente...":
        st.error("⚠️ Por favor, selecciona un Cliente válido.")
    elif not lista_documentos:
        st.error("⚠️ Por favor, ingresa al menos una Factura o Guía.")
    elif recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, selecciona el receptor.")
    else:
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            
            for doc in lista_documentos:
                cursor.execute(
                    """
                    INSERT INTO despachos (
                        fecha, hora_despacho, hora_llegada, conductor, 
                        orden_trabajo, cliente, cantidad_etiquetas, 
                        cantidad_cajas, factura_guia, recepcionado_por
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """,
                    (
                        fecha.strftime("%Y-%m-%d"),
                        hora_despacho.strftime("%H:%M"),
                        hora_llegada.strftime("%H:%M"),
                        conductor,
                        orden_trabajo,
                        cliente,
                        int(cant_etiquetas),
                        int(cant_cajas),
                        doc,
                        recepcionado_por
                    )
                )
            conn.commit()
            conn.close()
                
            st.success(f"✅ ¡Despacho registrado en la nube! Se agregaron {len(lista_documentos)} filas.")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al guardar en la base de datos: {e}")

# Sección de Visualización y Descarga desde SQL Nativo
try:
    if os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        df_mostrar = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC LIMIT 50;", conn)
        conn.close()
        
        if not df_mostrar.empty:
            st.markdown("---")
            st.subheader("📋 Últimos despachos registrados (Base de Datos)")
            
            # Limpiar la visualización para el usuario
            df_visualizar = df_mostrar.drop(columns=["id"], errors="ignore")
            st.dataframe(df_visualizar.head(15), use_container_width=True)
            
            # Crear el archivo Excel en memoria para la descarga
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_visualizar.to_excel(writer, index=False, sheet_name='Despachos')
            
            st.download_button(
                label="📥 Descargar Base de Datos Completa (Excel)",
                data=buffer.getvalue(),
                file_name="bitacora_despachos_nube.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
except Exception:
    pass
