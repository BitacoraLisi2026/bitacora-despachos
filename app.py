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

PASSWORD_ADMIN = "admin123"
# ==============================================================

st.set_page_config(page_title="Gestión de Despachos", page_icon="🚚", layout="wide") 

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

# ==============================================================
# 📊 MÓDULO 1: DASHBOARD EJECUTIVO (INFORME GERENCIAL)
# ==============================================================
st.markdown("<h1 style='font-size: 32px; color: #1E3A8A; margin-bottom: 5px;'>📊 Panel de Gestión e Informe de Despachos</h1>", unsafe_allow_html=True)
st.write("Resumen ejecutivo del estado de las gestiones y distribución logística en tiempo real.")

try:
    if os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        df_informe = pd.read_sql_query("SELECT * FROM despachos;", conn)
        conn.close()
        
        if not df_informe.empty:
            # 1. Fila de Tarjetas de Métricas Ejecutivas
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric(label="📄 Total Documentos Procesados", value=f"{len(df_informe)} guías/facturas")
            with m2:
                st.metric(label="📦 Total Cajas Distribuidas", value=f"{int(df_informe['cantidad_cajas'].sum()):,}")
            with m3:
                st.metric(label="🏷️ Total Etiquetas Controladas", value=f"{int(df_informe['cantidad_etiquetas'].sum()):,}")
            with m4:
                clientes_activos = df_informe['cliente'].nunique()
                st.metric(label="🏢 Clientes Atendidos", value=clientes_activos)
                
            st.markdown("---")
            
            # 2. Fila de Gráficos Estadísticos
            g1, g2 = st.columns(2)
            
            with g1:
                st.markdown("### 🏢 Volumen de Despachos por Cliente")
                despachos_cliente = df_informe['cliente'].value_counts().reset_index()
                despachos_cliente.columns = ['Cliente', 'Documentos']
                st.bar_chart(data=despachos_cliente, x='Cliente', y='Documentos', color="#1E3A8A")
                
            with g2:
                st.markdown("### 🚚 Operaciones por Conductor / Transporte")
                despachos_conductor = df_informe['conductor'].value_counts().reset_index()
                despachos_conductor.columns = ['Conductor', 'Despachos']
                st.bar_chart(data=despachos_conductor, x='Conductor', y='Despachos', color="#10B981")
                
            st.markdown("<br>", unsafe_allow_html=True)
        else:
            st.info("ℹ️ Esperando los primeros registros para construir el informe estadístico gerencial.")
except Exception as e:
    st.warning(f"Cargando panel estadístico... ({e})")

st.markdown("---")

# ==============================================================
# 🚚 MÓDULO 2: INGRESO DE DATOS (FORMULARIO ABIERTO)
# ==============================================================
st.markdown("<h2 style='font-size: 24px; margin-bottom: 0px;'>🚚 Registro de Nuevos Despachos</h2>", unsafe_allow_html=True)
st.write("Área operativa para la digitación en terreno desde dispositivos móviles o PC.")

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
                    (fecha.strftime("%Y-%m-%d"), hora_despacho.strftime("%H:%M"), hora_llegada.strftime("%H:%M"),
                     conductor, orden_trabajo, cliente, int(cant_etiquetas), int(cant_cajas), doc, recepcionado_por)
                )
            conn.commit()
            conn.close()
            st.success(f"✅ ¡Despacho registrado en la nube!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al guardar: {e}")

# ==============================================================
# ⚙️ MÓDULO 3: HISTORIAL, CONTROL DE DATOS Y DESCARGA
# ==============================================================
df_completo = pd.DataFrame()

try:
    if os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        df_completo = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC;", conn)
        conn.close()
except Exception as e:
    st.error(f"Error al leer la base de datos: {e}")

if not df_completo.empty:
    st.markdown("---")
    st.subheader("📋 Últimos despachos registrados")
    
    with st.expander("🛠️ Panel de Administración (Modificar o borrar registros)"):
        clave = st.text_input("Ingresa la clave de administrador para desbloquear la edición:", type="password")
        
        if clave == PASSWORD_ADMIN:
            st.info("💡 Haz doble clic en una celda para editarla. Para borrar filas completas, selecciónalas y presiona 'Supr'.")
            df_editado = st.data_editor(df_completo, num_rows="dynamic", use_container_width=True, key="editor_bd")
            boton_guardar_cambios = st.button("🔒 Guardar Cambios en la Nube")
            
            if boton_guardar_cambios:
                try:
                    conn = sqlite3.connect(DB_FILE)
                    df_editado.to_sql("despachos", conn, if_exists="replace", index=False)
                    conn.commit()
                    conn.close()
                    st.success("✅ ¡Base de datos actualizada con éxito en la nube!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al actualizar los cambios: {e}")
        elif clave != "":
            st.error("🔑 Clave incorrecta.")
    
    # Vista normal estática para usuarios generales
    df_visualizar = df_completo.drop(columns=["id"], errors="ignore")
    st.dataframe(df_visualizar.head(10), use_container_width=True)
    
    # Conversión directa e infalible en 2 líneas sin declarar objetos anidados
    buffer = io.BytesIO()
    df_visualizar.to_excel(buffer, index=False, engine="openpyxl")
    
    st.download_button(
        label="📥 Descargar Base de Datos Completa (Excel)",
        data=buffer.getvalue(),
        file_name="bitacora_despachos_nube.xlsx",
