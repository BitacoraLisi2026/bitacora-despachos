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

# Función para guardar registros independientes por cada guía
def guardar_registro(fecha_salida, cliente, guias_texto, recepciona, tipo_transporte, transporte):
    df = cargar_datos()
    ahora = datetime.now()
    
    # Separamos el texto por comas y limpiamos los espacios en blanco de cada guía
    lista_guias = [g.strip() for g in guias_texto.split(",") if g.strip()]
    
    # Creamos una lista de diccionarios (una fila para cada guía)
    nuevas_filas = []
    for guia in lista_guias:
        nueva_entrega = {
            "Fecha Salida": fecha_salida.strftime("%Y-%m-%d"),
            "Hora de Entrega": ahora.strftime("%H:%M:%S"),
            "Cliente": cliente,
            "Guías / Facturas": guia,  # <- Aquí va una sola guía por fila
            "Quién Recepciona": recepciona,
            "Tipo Transporte": tipo_transporte,
            "Transporte Utilizado": transporte,
            "Fecha Registro": ahora.strftime("%Y-%m-%d %H:%M:%S")
        }
        nuevas_filas.append(nueva_entrega)
    
    # Concatenamos todas las nuevas filas creadas hacia abajo
    df = pd.concat([df, pd.DataFrame(nuevas_filas)], ignore_index=True)
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

# Formulario de entrada de datos original
with st.form(key="formulario_entrega", clear_on_submit=True):
    
    fecha_salida = st.date_input("Fecha de Salida *", value=date.today())
    
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Cliente *")
        
        # El usuario ingresa las guías separadas por comas de forma simple
        guias_facturas = st.text_input(
            "Guías / Facturas *", 
            placeholder="Ej: 1024, 1025, 1026"
        )
    
    with col2:
        recepciona = st.text_input("Quién Recepciona *")
        tipo_transporte = st.radio(
            "Tipo de Transporte",
            ["Interno", "Externo"],
            horizontal=True
        )
        
    transporte = st.text_input("¿Qué transporte utilizó? *", placeholder="Ej: Camión Propio N°2, Starken, Chilexpress...")
    
    st.markdown("<small>* Campos obligatorios</small>", unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label="Registrar Entrega")

# Lógica al presionar el botón de guardar
if boton_guardar:
    if cliente and guias_facturas and recepciona and transporte:
        guardar_registro(fecha_salida, cliente, guias_facturas, recepciona, tipo_transporte, transporte)
        st.success("✅ ¡Entrega registrada! Las guías se han desglosado en filas independientes.")
        st.rerun()
    else:
        st.error("⚠️ Por favor, complete todos los campos obligatorios antes de enviar.")

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
