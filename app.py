import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import io

# Configuración de la página y el navegador
st.set_page_config(
    page_title="Control de Entregas", 
    page_icon="📦", 
    layout="centered"
)

# --- CONFIGURACIÓN DEL LOGO ---
NOMBRES_LOGO = ["logo.png", "logo.jpg", "logo.jpeg"]
logo_encontrado = None

for nombre in NOMBRES_LOGO:
    if os.path.exists(nombre):
        logo_encontrado = nombre
        break

# Nombre del archivo de base de datos local
DB_FILE = "registro_entregas.csv"

# Inicializar la lista temporal de guías en la sesión si no existe
if "lista_guias_temporal" not in st.session_state:
    st.session_state.lista_guias_temporal = []

# Función para cargar los datos existentes sin perder el formato
def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    else:
        return pd.DataFrame(columns=[
            "Fecha Salida", 
            "Hora de Entrega",
            "Cliente", 
            "Guías / Facturas", 
            "Quién Recepciona", 
            "Tipo Transporte", 
            "Transporte Utilizado",
            "Fecha Registro"
        ])

# Función para guardar un nuevo registro
def guardar_registro(fecha_salida, cliente, guias_lista, recepciona, tipo_transporte, transporte):
    df = cargar_datos()
    ahora = datetime.now()
    
    # Juntamos todas las guías de la lista en un solo texto separado por comas para la fila de Excel
    guias_consolidadas = ", ".join(guias_lista)
    
    nueva_entrega = {
        "Fecha Salida": fecha_salida.strftime("%Y-%m-%d"),
        "Hora de Entrega": ahora.strftime("%H:%M:%S"),
        "Cliente": cliente,
        "Guías / Facturas": guias_consolidadas,
        "Quién Recepciona": recepciona,
        "Tipo Transporte": tipo_transporte,
        "Transporte Utilizado": transporte,
        "Fecha Registro": ahora.strftime("%Y-%m-%d %H:%M:%S")
    }
    df = pd.concat([df, pd.DataFrame([nueva_entrega])], ignore_index=True)
    df.to_csv(DB_FILE, index=False)

# Cargar los datos actuales al iniciar la app
df_actual = cargar_datos()

# --- INTERFAZ DE USUARIO ---
if logo_encontrado:
    col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
    with col_logo_2:
        st.image(logo_encontrado, use_container_width=True)

st.title("📦 Control de Registro de Entregas")
st.write("Complete la información detallada de la salida del producto:")

# --- BLOQUE DE AGREGAR GUÍAS (HACIA ABAJO) ---
st.subheader("📝 Agregar Guías o Facturas")
col_input, col_btn = st.columns([3, 1])

with col_input:
    nueva_guia = st.text_input("Ingrese número de Guía / Factura", key="input_guia_individual", placeholder="Ej: 1024")

with col_btn:
    st.write("") # Espacio para alinear el botón
    st.write("") 
    if st.button("➕ Añadir", use_container_width=True):
        if nueva_guia.strip():
            st.session_state.lista_guias_temporal.append(nueva_guia.strip())
            st.rerun()

# Mostrar la lista de guías agregadas hacia abajo
if st.session_state.lista_guias_temporal:
    st.write("**Guías añadidas para este despacho:**")
    for i, g in enumerate(st.session_state.lista_guias_temporal):
        st.markdown(f" * Guía N°: **{g}**")
    
    if st.button("🗑️ Limpiar lista de guías"):
        st.session_state.lista_guias_temporal = []
        st.rerun()
else:
    st.info("Aún no has agregado guías para esta entrega. Usa el campo de arriba.")

st.divider()

# --- FORMULARIO PRINCIPAL DE ENVÍO ---
with st.form(key="formulario_entrega_final"):
    fecha_salida = st.date_input("Fecha de Salida *", value=date.today())
    
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Cliente *")
    
    with col2:
        recepciona = st.text_input("Quién Recepciona *")
        tipo_transporte = st.radio(
            "Tipo de Transporte",
            ["Interno", "Externo"],
            horizontal=True
        )
        
    transporte = st.text_input("¿Qué transporte utilizó? *", placeholder="Ej: Camión Propio N°2, Starken, Chilexpress...")
    
    st.markdown("<small>* Campos obligatorios</small>", unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label="🚀 Registrar Entrega Completa")

# Lógica al presionar el botón de guardar definitivo
if boton_guardar:
    if not st.session_state.lista_guias_temporal:
        st.error("⚠️ Debe añadir al menos una guía en la sección de arriba antes de registrar la entrega.")
    elif cliente and recepciona and transporte:
        # Guardamos pasando la lista acumulada
        guardar_registro(fecha_salida, cliente, st.session_state.lista_guias_temporal, recepciona, tipo_transporte, transporte)
        
        # Limpiamos la lista temporal para la siguiente entrega
        st.session_state.lista_guias_temporal = []
        st.success("✅ ¡Entrega registrada con éxito en una sola fila!")
        st.rerun()
    else:
        st.error("⚠️ Por favor, complete todos los campos obligatorios (*).")

# --- SECCIÓN DE REPORTES ---
st.divider()
st.subheader("📋 Historial y Reporte en Excel")

if not df_actual.empty:
    st.dataframe(df_actual, use_container_width=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_actual.to_excel(writer, index=False, sheet_name='Historial de Entregas')
    
    st.download_button(
        label="📥 Descargar Reporte en Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"reporte_entregas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aún no se han registrado entregas en el sistema.")
