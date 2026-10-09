import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Nombre del archivo de Excel
EXCEL_FILE = "bitacora_despachos.xlsx"

# --- CONFIGURA AQUÍ TUS LISTAS DESPLEGABLES ---
LISTA_CLIENTES = [
    "Selecciona un cliente...", 
    "CASADONOSO", 
    "VIÑEDOS DE AGUIRRE", 
    "VIÑA LOS AROMOS",
    "VIÑA SOLIS"
]

LISTA_RECEPCION = [
    "Selecciona quién recibe...", 
    "Camila Donoso", 
    "María López", 
    "Carlos Rodríguez", 
    "Ana Martínez"
]
# ----------------------------------------------

# Configuración de la página
st.set_page_config(page_title="Bitácora de Despachos", page_icon="🚚", layout="centered")
st.title("🚚 Registro de Bitácora de Despachos")
st.write("Introduce los datos del despacho para registrarlos en el archivo de Excel.")

# Formulario de entrada de datos
with st.form(key="formulario_bitacora", clear_on_submit=True):
    # Campos de fecha y hora
    col1, col2, col3 = st.columns(3)
    with col1:
        fecha = st.date_input("Fecha", value=datetime.today())
    with col2:
        hora_despacho = st.time_input("Hora de Despacho")
    with col3:
        hora_llegada = st.time_input("Hora de Llegada")
        
    # Campos de texto y desplegables principales
    col4, col5 = st.columns(2)
    with col4:
        # Campo Conductor fijo con el nombre solicitado
        conductor = st.text_input("Conductora", value="Elizabeth Utrera", disabled=False)
        orden_trabajo = st.text_input("Orden de Trabajo")
    with col5:
        # Menú desplegable para Cliente
        cliente = st.selectbox("Cliente", options=LISTA_CLIENTES)
        factura_guia = st.text_input("Factura o Guía")
        
    # Campos numéricos de carga
    col6, col7 = st.columns(2)
    with col6:
        cant_etiquetas = st.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    with col7:
        cant_cajas = st.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    # Menú desplegable para Recepción
    recepcionado_por = st.selectbox("Recepcionado Por", options=LISTA_RECEPCION)
    
    st.markdown("---")
    st.subheader("📊 Control de Kilometraje")
    # Campos numéricos de kilometraje
    col8, col9 = st.columns(2)
    with col8:
        km_inicial = st.number_input("Kilómetros Iniciales", min_value=0.0, step=1.0, format="%.1f")
    with col9:
        km_final = st.number_input("Kilómetros Finales", min_value=0.0, step=1.0, format="%.1f")
    
    # Botón de envío
    boton_guardar = st.form_submit_button(label="💾 Registrar Despacho")

# Lógica para guardar en Excel al hacer clic
if boton_guardar:
    # Validaciones obligatorias para los desplegables y kilómetros
    if cliente == "Selecciona un cliente...":
        st.error("⚠️ Por favor, selecciona un Cliente válido de la lista.")
    elif recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, selecciona la persona que Recepcionó de la lista.")
    elif km_final < km_inicial:
        st.error("⚠️ Error: Los Kilómetros Finales no pueden ser menores que los Kilómetros Iniciales.")
    else:
        # Calcular automáticamente los kilómetros recorridos en el viaje
        km_recorridos = km_final - km_inicial
        
        # Formatear los datos en un diccionario
        nuevo_registro = {
            "FECHA": fecha.strftime("%Y-%m-%d"),
            "HORA DE DESPACHO": hora_despacho.strftime("%H:%M"),
            "HORA DE LLEGADA": hora_llegada.strftime("%H:%M"),
            "CONDUCTOR": conductor,
            "ORDEN DE TRABAJO": orden_trabajo,
            "CLIENTE": cliente,
            "CANTIDAD DE ETIQUETAS": cant_etiquetas,
            "CANTIDAD DE CAJAS": cant_cajas,
            "FACTURA O GUIA": factura_guia,
            "RECEPCIONADO POR": recepcionado_por,
            "KM INICIAL": km_inicial,
            "KM FINAL": km_final,
            "KM RECORRIDOS": km_recorridos
        }
        
        # Crear DataFrame con el nuevo registro
        df_nuevo = pd.DataFrame([nuevo_registro])
        
        # Si el archivo ya existe, cargar datos previos y añadir el nuevo
        if os.path.exists(EXCEL_FILE):
            df_existente = pd.read_excel(EXCEL_FILE)
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            df_final = df_nuevo
            
        # Guardar de vuelta al archivo de Excel
        df_final.to_excel(EXCEL_FILE, index=False)
        st.success(f"✅ ¡Registro guardado con éxito! Kilómetros recorridos: {km_recorridos:.1f} km")

# Mostrar los últimos registros en la app
if os.path.exists(EXCEL_FILE):
    st.subheader("📋 Últimos despachos registrados")
    df_mostrar = pd.read_excel(EXCEL_FILE)
    st.dataframe(df_mostrar.tail(5))

# Botón web para descargar el archivo Excel directamente
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            label="📥 Descargar Bitácora en Excel",
            data=f,
            file_name="bitacora_despachos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
