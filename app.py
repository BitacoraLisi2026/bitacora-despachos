import streamlit as st
import pandas as pd
from datetime import datetime

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

# --- CONEXIÓN DIRECTA PARA LEER GOOGLE SHEETS ---
try:
    conn = st.connection("gsheets", type="sheets")
except Exception:
    conn = None

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

# Formulario básico
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

# Lógica de Guardado usando la API de envío directo
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
            # Obtener el ID de la hoja desde los secrets de manera limpia
            url_planilla = st.secrets["connections"]["gsheets"]["spreadsheet"]
            sheet_id = url_planilla.split("/d/")[1].split("/")[0]
            
            # Dirección del script de Google alternativo para escritura libre
            url_envio = f"https://google.com" 
            
            # Para asegurar la escritura sin Cuenta de Servicio pesada, simulamos la inserción mediante Pandas URL si falla
            # Intentamos la inserción nativa limpia primero
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
                df_existente = conn.read()
            except Exception:
                df_existente = pd.DataFrame()
                
            if not df_existente.empty:
                df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
            else:
                df_final = df_nuevos
                
            # Intento de empuje directo usando el conector de Sheets básico
            conn.update(data=df_final)
            st.success(f"✅ ¡Guardado en la nube! Se registraron {len(lista_documentos)} documentos.")
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error al guardar datos: {e}")

# Muestra la tabla de Google Sheets abajo si está conectada
if conn:
    try:
        st.markdown("---")
        st.subheader("📋 Últimos despachos registrados (Nube)")
        df_mostrar = conn.read()
        if not df_mostrar.empty:
            st.dataframe(df_mostrar.dropna(how="all").tail(10), use_container_width=True)
    except Exception:
        pass
