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
            "Conductor", 
            "Transporte Utilizado",
            "Fecha Registro"
        ])

# Cargar los datos actuales al iniciar la app
df_actual = cargar_datos()

# --- OBTENER LISTADO DE CLIENTES (ACTUALIZADO) ---
# Tu lista oficial de clientes y transportes
clientes_base = [
    "VIÑEDOS DE AGUIRRE",
    "VIÑA CASADONOSO",
    "VIÑA AROMO",
    "VIÑA CASA SOLIS",
    "VIÑA CONCHA Y TORO",
    "VIÑA LUIS FELIPE EDWARDS",
    "TRANSPORTES COLCHAGUA"
]

if not df_actual.empty:
    # Extrae los clientes que ya han sido guardados antes en el archivo CSV para no perderlos
    clientes_historicos = df_actual["Cliente"].dropna().unique().tolist()
    # Junta ambas listas quitando duplicados y los ordena alfabéticamente
    lista_clientes_completa = sorted(list(set(clientes_base + clientes_historicos)))
else:
    lista_clientes_completa = sorted(clientes_base)

# Añadimos la opción para ingresar uno completamente nuevo al final
lista_clientes_completa.append("➕ Agregar Cliente Nuevo...")

# Función para guardar registros independientes por cada línea (guía)
def guardar_registro(fecha_salida, cliente, guias_bloque_texto, recepciona, conductor, transporte):
    df = cargar_datos()
    ahora = datetime.now()
    
    # Separamos el bloque de texto por saltos de línea y limpiamos espacios vacíos
    lista_guias = [g.strip() for g in guias_bloque_texto.split("\n") if g.strip()]
    
    # Creamos una fila independiente en el Excel por cada línea escrita
    nuevas_filas = []
    for copia_guia in lista_guias:
        nueva_entrega = {
            "Fecha Salida": fecha_salida.strftime("%Y-%m-%d"),
            "Hora de Entrega": ahora.strftime("%H:%M:%S"),
            "Cliente": cliente,
            "Guías / Facturas": copia_guia,
            "Quién Recepciona": recepciona,
            "Conductor": conductor,
            "Transporte Utilizado": transporte,
            "Fecha Registro": ahora.strftime("%Y-%m-%d %H:%M:%S")
        }
        nuevas_filas.append(nueva_entrega)
    
    # Concatenamos las filas al registro general
    df = pd.concat([df, pd.DataFrame(nuevas_filas)], ignore_index=True)
    df.to_csv(DB_FILE, index=False)

# --- INTERFAZ DE USUARIO ---
if logo_encontrado:
    col_logo_1, col_logo_2, col_logo_3 = st.columns(3)
    with col_logo_2:
        st.image(logo_encontrado, use_container_width=True)

st.title("📦 Control de Registro de Entregas")
st.write("Complete la información detallada de la salida del producto:")

# --- DATOS FUERA DEL FORMULARIO PARA RESPUESTA INMEDIATA ---
fecha_salida = st.date_input("Fecha de Salida *", value=date.today())

col_cond_1, col_cond_2 = st.columns(2)
with col_cond_1:
    conductor = st.radio(
        "Conductor *",
        ["Elizabeth", "Empresa Externa"],
        index=0,
        horizontal=True
    )

# Lógica dinámica inmediata para el transporte
transporte_final = ""
with col_cond_2:
    if conductor == "Elizabeth":
        st.info("ℹ️ Asignado: **Vehículo Interno**")
        transporte_final = "Vehículo Interno"
    else:
        empresa_externa_sel = st.selectbox(
            "Seleccione Empresa Externa *",
            [
                "Starken", 
                "Chilexpress", 
                "Blue Express", 
                "CorreosChile", 
                "Varmontt", 
                "Transportes Cruz del Sur", 
                "Pullman Cargo", 
                "TNT / FedEx", 
                "Otra Empresa"
            ]
        )
        if empresa_externa_sel == "Otra Empresa":
            transporte_final = st.text_input("Escriba el nombre de la empresa *", placeholder="Ej: Transportes Juanito...")
        else:
            transporte_final = empresa_externa_sel

st.divider()

# --- SELECCIÓN DINÁMICA DEL CLIENTE ---
col_cli_1, col_cli_2 = st.columns(2)
with col_cli_1:
    cliente_seleccionado = st.selectbox("Seleccione el Cliente *", lista_clientes_completa)

cliente_final = ""
with col_cli_2:
    if cliente_seleccionado == "➕ Agregar Cliente Nuevo...":
        cliente_final = st.text_input("Escriba el nombre del nuevo cliente *", placeholder="Ej: Nueva Viña S.A.")
    else:
        cliente_final = cliente_seleccionado
        st.info(f"📋 Seleccionado: **{cliente_final}**")

# --- FORMULARIO RESTANTE PARA LAS GUÍAS ---
with st.form(key="formulario_datos_entrega", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        guias_facturas = st.text_area(
            "Guías / Facturas (Escriba una por línea usando ENTER) *", 
            placeholder="1024\n1025\n1026",
            height=120
        )
    with col2:
        recepciona = st.text_input("Quién Recepciona *")
    
    st.markdown("<small>* Campos obligatorios</small>", unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label="🚀 Registrar Entrega Completa")

# Lógica al presionar el botón de guardar definitivo
if boton_guardar:
    if cliente_final and guias_facturas and recepciona and conductor and transporte_final:
        guardar_registro(fecha_salida, cliente_final, guias_facturas, recepciona, conductor, transporte_final)
        st.success("✅ ¡Entrega registrada! Las filas se generaron correctamente hacia abajo.")
        st.rerun()
    else:
        st.error("⚠️ Por favor, asegúrese de completar todos los campos obligatorios (*) antes de guardar.")

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
