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

# Nombre del archivo de base de datos local
DB_FILE = "registro_entregas.csv"

# Función para cargar los datos existentes sin perder el formato
def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    else:
        return pd.DataFrame(columns=[
            "Fecha Salida", 
            "Cliente", 
            "Guías / Facturas", 
            "Quién Recepciona", 
            "Hora de entrega
            "Tipo Transporte", 
            "Transporte Utilizado",
            "Fecha Registro"
        ])

# Función para guardar un nuevo registro
def guardar_registro(fecha_salida, cliente, guias_facturas, recepciona, tipo_transporte, transporte):
    df = cargar_datos()
    nueva_entrega = {
        "Fecha Salida": fecha_salida.strftime("%Y-%m-%d"),
        "Cliente": cliente,
        "Guías / Facturas": guias_facturas,
        "Quién Recepciona": recepciona,
        "Hora de entrega": entrega,
        "Tipo Transporte": tipo_transporte,
        "Transporte Utilizado": transporte,
        "Fecha Registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    df = pd.concat([df, pd.DataFrame([nueva_entrega])], ignore_index=True)
    df.to_csv(DB_FILE, index=False)

# Cargar los datos actuales al iniciar la app
df_actual = cargar_datos()

# --- INTERFAZ DE USUARIO ---
st.title("📦 Control de Registro de Entregas")
st.write("Complete la información detallada de la salida del producto:")

# Formulario de entrada de datos
with st.form(key="formulario_entrega", clear_on_submit=True):
    
    # Campo de fecha
    fecha_salida = st.date_input("Fecha de Salida *", value=date.today())
    
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Cliente *")
        guias_facturas = st.text_input("Número de Guías / Facturas *")
    
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
        st.success("✅ ¡Entrega registrada y guardada con éxito!")
        st.rerun()
    else:
        st.error("⚠️ Por favor, complete todos los campos obligatorios antes de enviar.")

# --- SECCIÓN DE REPORTES ---
st.divider()
st.subheader("📋 Historial y Reporte en Excel")

if not df_actual.empty:
    # Mostrar tabla con los datos actuales ordenados por el último registro
    st.dataframe(df_actual, use_container_width=True)
    
    # Generar el archivo Excel en memoria utilizando OpenPyXL
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_actual.to_excel(writer, index=False, sheet_name='Historial de Entregas')
    
    # Botón de descarga directa
    st.download_button(
        label="📥 Descargar Reporte en Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"reporte_entregas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aún no se han registrado entregas en el sistema.")

