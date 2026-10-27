import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import io

# Importaciones para el diseño elegante de Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# --- OBTENER LISTADO DE CLIENTES ---
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
    clientes_historicos = df_actual["Cliente"].dropna().unique().tolist()
    lista_clientes_completa = sorted(list(set(clientes_base + clientes_historicos)))
else:
    lista_clientes_completa = sorted(clientes_base)

lista_clientes_completa.insert(0, "Selecciona un cliente...")
lista_clientes_completa.append("➕ Agregar Cliente Nuevo...")


# --- OBTENER LISTADO DE RECEPTORES ---
receptores_base = [
    "CAMILA VILLASANA",
    "OLGA GONZALEZ",
    "JORGE DIAZ",
    "JAVIERA RAMIREZ",
    "ENCARGADO COLCHAGUA"
]

if not df_actual.empty:
    receptores_historicos = df_actual["Quién Recepciona"].dropna().unique().tolist()
    lista_receptores_completa = sorted(list(set(receptores_base + receptores_historicos)))
else:
    lista_receptores_completa = sorted(receptores_base)

lista_receptores_completa.insert(0, "Selecciona quién recibe...")
lista_receptores_completa.append("➕ Agregar Receptor Nuevo...")


# Función para guardar registros independientes por cada línea (guía)
def guardar_registro(fecha_salida, cliente, guias_bloque_texto, recepciona, conductor, transporte):
    df = cargar_datos()
    ahora = datetime.now()
    
    lista_guias = [g.strip() for g in guias_bloque_texto.split("\n") if g.strip()]
    
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
                "Selecciona una empresa...",
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
    elif cliente_seleccionado == "Selecciona un cliente...":
        cliente_final = "Selecciona un cliente..."
    else:
        cliente_final = cliente_seleccionado
        st.info(f"📋 Seleccionado: **{cliente_final}**")

st.divider()

# --- SELECCIÓN DINÁMICA DE QUIÊN RECEPCIONA ---
col_rec_1, col_rec_2 = st.columns(2)
with col_rec_1:
    receptor_seleccionado = st.selectbox("Seleccione Quién Recibe *", lista_receptores_completa)

receptor_final = ""
with col_rec_2:
    if receptor_seleccionado == "➕ Agregar Receptor Nuevo...":
        receptor_final = st.text_input("Escriba el nombre del nuevo receptor *", placeholder="Ej: Juan Pérez")
    elif receptor_seleccionado == "Selecciona quién recibe...":
        receptor_final = "Selecciona quién recibe..."
    else:
        receptor_final = receptor_seleccionado
        st.info(f"👤 Asignado: **{receptor_final}**")

st.divider()

# --- FORMULARIO RESTANTE PARA LAS GUÍAS ---
with st.form(key="formulario_datos_entrega", clear_on_submit=True):
    guias_facturas = st.text_area(
        "Guías / Facturas (Escriba una por línea usando ENTER) *", 
        placeholder="1024\n1025\n1026",
        height=120
    )
    
    st.markdown("<small>* Campos obligatorios</small>", unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label="🚀 Registrar Entrega Completa")

# Lógica al presionar el botón de guardar definitivo con todas las validaciones estrictas
if boton_guardar:
    if cliente_final == "Selecciona un cliente...":
        st.error("⚠️ Debe seleccionar un cliente válido de la lista.")
    elif transporte_final == "Selecciona una empresa...":
        st.error("⚠️ Debe seleccionar una empresa de transporte válida para la opción Externa.")
    elif receptor_final == "Selecciona quién recibe...":
        st.error("⚠️ Debe seleccionar quién recibe el producto de la lista.")
    elif cliente_final and guias_facturas and receptor_final and conductor and transporte_final:
        guardar_registro(fecha_salida, cliente_final, guias_facturas, receptor_final, conductor, transporte_final)
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
        df_actual.to_excel(writer, index=False, sheet_name='Bitácora de Despachos')
        
        workbook = writer.book
        worksheet = workbook['Bitácora de Despachos']
        
        fuente_titulo = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        fuente_datos = Font(name='Calibri', size=11, bold=False, color='000000')
        
        relleno_titulo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        relleno_par = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        relleno_impar = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        
        alineacion_centro = Alignment(horizontal='center', vertical='center')
        alineacion_izquierda = Alignment(horizontal='left', vertical='center')
        
        borde_fino = Side(border_style="thin", color="D9D9D9")
        cuadrícula = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
        
        worksheet.row_dimensions.height = 25
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = fuente_titulo
            cell.fill = relleno_titulo
            cell.alignment = alineacion_centro
            cell.border = cuadrícula
            
        for row_idx in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[row_idx].height = 18
            relleno_actual = relleno_par if row_idx % 2 == 0 else relleno_impar
            
            for col_idx in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = fuente_datos
                cell.fill = relleno_actual
                cell.border = cuadrícula
                
                # Modificado a tupla fija para evitar el error de desaparición de texto
                columnas_centradas = (1, 2, 4, 6, 8)
                if col_idx in columnas_centradas:
                    cell.alignment = alineacion_centro

