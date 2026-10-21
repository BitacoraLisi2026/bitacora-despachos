import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

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

# --- CONEXIÓN DIRECTA A GOOGLE SHEETS ---
# Esto maneja la base de datos en tiempo real en la nube
conn = st.connection("gsheets", type=GSheetsConnection)

# --- AJUSTE DE HORA LOCAL (CHILE) ---
try:
    import zoneinfo
    zona_horaria = zoneinfo.ZoneInfo("America/Santiago") 
    hora_local_dt = datetime.now(zona_horaria)
except Exception:
    hora_local_dt = datetime.now()

fecha_actual = hora_local_dt.date()
hora_actual = hora_local_dt.time()

st.markdown("<h2 style='font-size: 30px; margin-bottom: 0px;'>🚚 Bitácora de Despachos</h2>", unsafe_allow_html=True)
st.write("Ingresa los datos para actualizar la planilla automáticamente en la nube.")

# Formulario
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

# Lógica de Guardado en la Nube
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
        nuevos_registros = []
        for doc in lista_documentos:
            nuevos_registros.append({
                "FECHA": fecha.strftime("%Y-%m-%d"),
                "HORA DE DESPACHO": hora_despacho.strftime("%H:%M"),
                "HORA DE LLEGADA": hora_llegada.strftime("%H:%M"),
                "CONDUCTOR": conductor, 
                "ORDEN DE TRABAJO": orden_trabajo,
                "CLIENTE": cliente,
                "CANTIDAD DE ETIQUETAS": int(cant_etiquetas),
                "CANTIDAD DE CAJAS": int(cant_cajas),
                "FACTURA O GUIA": doc,
                "RECEIPCIONADO POR": recepcionado_por
            })
        
        df_nuevos = pd.DataFrame(nuevos_registros)
        
        try:
            # 1. Leer los datos actuales que ya existen en Google Sheets
            try:
                df_existente = conn.read()
                # Limpiar filas completamente vacías si existen
                df_existente = df_existente.dropna(how="all")
            except Exception:
                df_existente = pd.DataFrame()
            
            # 2. Combinar lo viejo con lo nuevo
            if not df_existente.empty:
                df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
            else:
                df_final = df_nuevos
                
            # 3. Guardar directamente en la hoja de cálculo de Google
            conn.update(data=df_final)
            st.success(f"✅ ¡Guardado en la nube! Se registraron {len(lista_documentos)} documentos.")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al conectar con Google Sheets: {e}")

# Visualización en tiempo real desde la nube
try:
    st.markdown("---")
    st.subheader("📋 Últimos despachos registrados (Nube)")
    df_mostrar = conn.read()
    if not df_mostrar.empty:
        st.dataframe(df_mostrar.dropna(how="all").tail(10), use_container_width=True)
    else:
        st.info("Aún no hay registros en la base de datos de Google.")
except Exception:
    pass
