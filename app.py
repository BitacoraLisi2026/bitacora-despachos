import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# Nombre del archivo de Excel
EXCEL_FILE = "bitacora_despachos.xlsx"

# ==============================================================
# 📋 CONFIGURA AQUÍ TUS DATOS REALES (Escríbelos dentro de las comillas)
# ==============================================================
LISTA_CLIENTES = [
    "Selecciona un cliente...", 
    "VIÑEDOS DE AGUIRRE",
    "VIÑA CASADONOSO",
    "VIÑA AROMO",
    "VIÑA CASA SOLIS",
    "VIÑA CONCHA Y TORO",
    "VIÑA LUIS FELIPE EDWARDS",
    "TRANSPORTES COLCHAGUA"
]

LISTA_RECEIPCION = [
    "Selecciona quién recibe...", 
    "CAMILA VILLASANA",
    "OLGA GONZALEZ",
    "JORGE DIAZ",
    "JAVIERA RAMIREZ",
    "ENCARGADO COLCHAGUA"
]

# Lista Unificada de Conductores y Transportes Externos
LISTA_CONDUCTORES = [
    "Selecciona el conductor o transporte...",
    "ELIZABETH",               # <- FLOTA PROPIA (ACTIVA KILÓMETROS)
    "ELIZABETH - COLCHAGUA",  # <- FLOTA PROPIA (ACTIVA KILÓMETROS)
    "CHILEXPRESS",            # <- EXTERNO (OMITE KILÓMETROS)
    "STARKEN",                # <- EXTERNO (OMITE KILÓMETROS)
    "FEDEX",                  # <- EXTERNO (OMITE KILÓMETROS)
    "BLUE EXPRESS",           # <- EXTERNO (OMITE KILÓMETROS)
    "OTRO EXTERNO"
]
# ==============================================================

# Configuración elegante de la página
st.set_page_config(page_title="Bitácora de Despachos", page_icon="🚚", layout="centered")

# --- AJUSTE DE HORA LOCAL (ZONA HORARIA DE CHILE) ---
try:
    import zoneinfo
    zona_horaria = zoneinfo.ZoneInfo("America/Santiago") 
    hora_local_dt = datetime.now(zona_horaria)
except Exception:
    hora_local_dt = datetime.now()

fecha_actual = hora_local_dt.date()
hora_actual = hora_local_dt.time()
# --------------------------------------------------------------

# 🎨 Bloque del Logo: Busca tu imagen sin bloquear la app
try:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=180)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=180)
except Exception:
    pass

st.markdown("<h2 style='font-size: 30px; margin-bottom: 0px;'>🚚 Bitácora de Despachos</h2>", unsafe_allow_html=True)
st.write("Vamos que se puede, ingresa todos los datos para tu planilla.")

# Formulario elegante de entrada de datos
with st.form(key="formulario_bitacora", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        fecha = st.date_input("Fecha", value=fecha_actual)
    with col2:
        hora_despacho = st.time_input("Hora de Despacho", value=hora_actual)
    with col3:
        hora_llegada = st.time_input("Hora de Llegada", value=hora_actual)
        
    st.markdown("---")
    
    col4, col5 = st.columns(2)
    with col4:
        # Ahora el conductor es una lista desplegable unificada
        conductor = st.selectbox("Conductor / Empresa de Transporte", options=LISTA_CONDUCTORES)
        orden_trabajo = st.text_input("Orden de Trabajo")
        
    with col5:
        cliente = st.selectbox("Cliente (Remitente)", options=LISTA_CLIENTES)
        facturas_guias_texto = st.text_area("Facturas o Guías (Escribe una por línea si son varias)")

    # Detección inteligente del tipo de viaje
    es_interno = (conductor == "Elizabeth - Colchagua")
    # Es externo si seleccionó cualquier transporte que NO sea el de Elizabeth ni el mensaje inicial
    es_externo = (conductor != "Selecciona el conductor o transporte..." and not es_interno)
        
    st.markdown("---")
    col6, col7 = st.columns(2)
    with col6:
        cant_etiquetas = st.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    with col7:
        cant_cajas = st.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    recepcionado_por = st.selectbox("Recepcionado Por", options=LISTA_RECEIPCION)
    
    st.markdown("---")
    st.subheader("📊 Control de Kilometraje")
    
    if es_externo:
        st.info(f"ℹ️ Transporte externo ({conductor}) seleccionado: El kilometraje se registrará automáticamente en 0.0")
        km_inicial = 0.0
        km_final = 0.0
    else:
        col8, col9 = st.columns(2)
        with col8:
            km_inicial = st.number_input("Kilómetros Iniciales", min_value=0.0, step=1.0, format="%.1f")
        with col9:
            km_final = st.number_input("Kilómetros Finales", min_value=0.0, step=1.0, format="%.1f")
    
    boton_guardar = st.form_submit_button(label="💾 Registrar Despacho")

# Lógica robusta para guardar en Excel
if boton_guardar:
    lista_documentos = [linea.strip() for linea in facturas_guias_texto.split("\n") if linea.strip()]
    
    if conductor == "Selecciona el conductor o transporte...":
        st.error("⚠️ Por favor, selecciona un Conductor o Empresa de transporte válida.")
    elif cliente == "Selecciona un cliente...":
        st.error("⚠️ Por favor, selecciona un Cliente válido de la lista.")
    elif len(lista_documentos) == 0:
        st.error("⚠️ Por favor, ingresa al menos una Factura o Guía.")
    elif recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, selecciona la persona que Recepcionó de la lista.")
    elif es_interno and km_final < km_inicial:
        st.error("⚠️ Error: Los Kilómetros Finales no pueden ser menores que los Kilómetros Iniciales.")
    else:
        km_recorridos = km_final - km_inicial if es_interno else 0.0
        nuevos_registros = []
        
        for doc in lista_documentos:
            registro = {
                "FECHA": fecha.strftime("%Y-%m-%d"),
                "HORA DE DESPACHO": hora_despacho.strftime("%H:%M"),
                "HORA DE LLEGADA": hora_llegada.strftime("%H:%M"),
                "CONDUCTOR": conductor, 
                "ORDEN DE TRABAJO": orden_trabajo,
                "CLIENTE": cliente,
                "CANTIDAD DE ETIQUETAS": cant_etiquetas,
                "CANTIDAD DE CAJAS": cant_cajas,
                "FACTURA O GUIA": doc,
                "RECEPCIONADO POR": recepcionado_por,
                "KM INICIAL": km_inicial,
                "KM FINAL": km_final,
                "KM RECORRIDOS": km_recorridos,
                "TIPO TRANSPORTE": "INTERNO" if es_interno else "EXTERNO"
            }
            nuevos_registros.append(registro)
        
        df_nuevos = pd.DataFrame(nuevos_registros)
        
        try:
            if os.path.exists(EXCEL_FILE):
                df_existente = pd.read_excel(EXCEL_FILE)
                df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
            else:
                df_final = df_nuevos
                
            df_final.to_excel(EXCEL_FILE, index=False)
            st.success(f"✅ ¡Se registraron exitosamente {len(lista_documentos)} documentos!")
        except Exception as e:
            st.error(f"❌ Error al guardar en Excel: {e}. Inténtalo de nuevo.")

# 📥 Bloque de descarga (Siempre visible abajo)
try:
    if os.path.exists(EXCEL_FILE):
        st.markdown("---")
        st.subheader("📋 Últimos despachos registrados")
        df_mostrar = pd.read_excel(EXCEL_FILE)
        st.dataframe(df_mostrar.tail(10))
        
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Descargar Bitácora en Excel",
                data=f,
                file_name="bitacora_despachos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
except Exception:
    pass
