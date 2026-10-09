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
    "DESPACHO EXTERNO",           # <- Opción clave que activa el modo externo
    "Cliente Real 1", 
    "Cliente Real 2", 
    "Cliente Real 3"
]

LISTA_RECEIPCION = [
    "Selecciona quién recibe...", 
    "Persona Real 1", 
    "Persona Real 2"
]

# Agrega o quita aquí tus empresas de transporte externo habituales
LISTA_TRANSPORTES_EXTERNOS = [
    "Selecciona la empresa de transporte...",
    "Chilexpress",
    "Starken",
    "FedEx",
    "Blue Express",
    "Transportes Colchagua",
    "Otro Externo"
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
st.write("Introduce los datos del despacho para registrarlos en el archivo de Excel.")

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
    with col5:
        cliente = st.selectbox("Cliente (Remitente)", options=LISTA_CLIENTES)
        facturas_guias_texto = st.text_area("Facturas o Guías (Escribe una por línea si son varias)")

    # Detección automática si es un despacho externo
    es_viaje_externo = (cliente == "DESPACHO EXTERNO")

    with col4:
        if es_viaje_externo:
            # Si seleccionan Despacho Externo, se abre la lista de tus transportistas
            conductor = st.selectbox("Selecciona la Empresa Externa", options=LISTA_TRANSPORTES_EXTERNOS)
            destino = st.text_input("¿Para dónde lo lleva? (Destino)", value="")
            nro_flete = st.text_input("Nº de Flete / Orden de Seguimiento", value="")
        else:
            # Si eligen un cliente normal, va Elizabeth Utrera por defecto
            conductor = st.text_input("Conductora", value="Elizabeth Utrera")
            destino = "Entrega Directa"
            nro_flete = "N/A"
            
        orden_trabajo = st.text_input("Orden de Trabajo")
        
    st.markdown("---")
    col6, col7 = st.columns(2)
    with col6:
        cant_etiquetas = st.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    with col7:
        cant_cajas = st.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    recepcionado_por = st.selectbox("Recepcionado Por", options=LISTA_RECEIPCION)
    
    st.markdown("---")
    st.subheader("📊 Control de Kilometraje")
    
    if es_viaje_externo:
        st.info("ℹ️ Despacho Externo seleccionado: El kilometraje se registrará automáticamente en 0.0")
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
    
    if cliente == "Selecciona un cliente...":
        st.error("⚠️ Por favor, selecciona un Cliente o el Tipo de Despacho válido.")
    elif es_viaje_externo and conductor == "Selecciona la empresa de transporte...":
        st.error("⚠️ Por favor, selecciona la Empresa Externa que realizará el traslado.")
    elif len(lista_documentos) == 0:
        st.error("⚠️ Por favor, ingresa al menos una Factura o Guía.")
    elif recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, selecciona la persona que Recepcionó de la lista.")
    elif not es_viaje_externo and km_final < km_inicial:
        st.error("⚠️ Error: Los Kilómetros Finales no pueden ser menores que los Kilómetros Iniciales.")
    elif es_viaje_externo and not destino.strip():
        st.error("⚠️ Por favor, indica para dónde lo lleva (Destino).")
    elif es_viaje_externo and not nro_flete.strip():
        st.error("⚠️ Por favor, indica el Número de Flete o Seguimiento.")
    else:
        km_recorridos = km_final - km_inicial if not es_viaje_externo else 0.0
        nuevos_registros = []
        
        for doc in lista_documentos:
            registro = {
                "FECHA": fecha.strftime("%Y-%m-%d"),
                "HORA DE DESPACHO": hora_despacho.strftime("%H:%M"),
                "HORA DE LLEGADA": hora_llegada.strftime("%H:%M"),
                "CONDUCTOR / EMPRESA": conductor, 
                "ORDEN DE TRABAJO": orden_trabajo,
                "CLIENTE": cliente if not es_viaje_externo else "CLIENTE EXTERNO TERCERIZADO",
                "CANTIDAD DE ETIQUETAS": cant_etiquetas,
                "CANTIDAD DE CAJAS": cant_cajas,
                "FACTURA O GUIA": doc,
                "RECEPCIONADO POR": recepcionado_por,
                "KM INICIAL": km_inicial,
                "KM FINAL": km_final,
                "KM RECORRIDOS": km_recorridos,
                "TIPO TRANSPORTE": "EXTERNO" if es_viaje_externo else "INTERNO",
                "DESTINO": destino,
                "NRO FLETE / SEGUIMIENTO": nro_flete
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
            st.success(f"✅ ¡Se registraron exitosamente {len(lista_documentos)} documentos para este viaje!")
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
