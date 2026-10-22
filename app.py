import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

# Librerías de diseño estético para Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "data.db"

# ==============================================================
# 📋 CONFIGURACIÓN DE DATOS MAESTROS
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

PASSWORD_ADMIN = "admin123"

# ==============================================================
# 🗄️ INICIALIZACIÓN DE BASE DE DATOS
# ==============================================================
def inicializar_bd():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS despachos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            hora_despacho TEXT,
            hora_llegada TEXT,
            conductor TEXT,
            orden_trabajo TEXT,
            cliente TEXT,
            cantidad_etiquetas INTEGER,
            cantidad_cajas INTEGER,
            factura_guia TEXT,
            recepcionado_por TEXT
        );
    """)
    conn.commit()
    conn.close()

inicializar_bd()

# --- Zona Horaria Local (Chile) ---
try:
    import zoneinfo
    zona_horaria = zoneinfo.ZoneInfo("America/Santiago") 
    hora_local_dt = datetime.now(zona_horaria)
except Exception:
    hora_local_dt = datetime.now()

fecha_actual = hora_local_dt.date()
hora_actual = hora_local_dt.time()

# Configuración de página de Streamlit
st.set_page_config(page_title="Control de Despachos", page_icon="🚚", layout="centered") 

# ==============================================================
# 🚚 MÓDULO 1: INGRESO DE DATOS (FORMULARIO OPERATIVO)
# ==============================================================
st.markdown("<h1 style='font-size: 28px; color: #1E3A8A; margin-bottom: 5px;'>🚚 Registro de Despachos Logísticos</h1>", unsafe_allow_html=True)
st.write("Complete el formulario para alimentar la bitácora y generar las planillas de descarga.")

with st.form(key="formulario_bitacora", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    with col1: fecha = st.date_input("Fecha", value=fecha_actual)
    with col2: hora_despacho = st.time_input("Hora Despacho", value=hora_actual)
    with col3: hora_llegada = st.time_input("Hora Llegada", value=hora_actual)
        
    st.markdown("---")
    col4, col5 = st.columns(2)
    with col4:
        conductor = st.selectbox("Conductor / Transporte", options=LISTA_CONDUCTORES)
        orden_trabajo = st.text_input("Orden de Trabajo (OT)")
    with col5:
        cliente = st.selectbox("Cliente (Remitente)", options=LISTA_CLIENTES)
        facturas_guias_texto = st.text_area("Facturas o Guías (Una por línea)")
        
    st.markdown("---")
    col6, col7 = st.columns(2)
    with col6: cant_etiquetas = st.number_input("Cantidad de Etiquetas", min_value=0, step=1)
    with col7: cant_cajas = st.number_input("Cantidad de Cajas", min_value=0, step=1)
        
    recepcionado_por = st.selectbox("Recepcionado Por", options=LISTA_RECEIPCION)
    
    boton_guardar = st.form_submit_button(label="💾 Registrar en Bitácora")

if boton_guardar:
    lista_documentos = [linea.strip() for linea in facturas_guias_texto.split("\n") if linea.strip()]
    if conductor == "Selecciona el conductor o transporte..." or cliente == "Selecciona un cliente..." or not lista_documentos or recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, rellena todos los campos obligatorios del formulario.")
    else:
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            for doc in lista_documentos:
                cursor.execute(
                    """
                    INSERT INTO despachos (
                        fecha, hora_despacho, hora_llegada, conductor, 
                        orden_trabajo, cliente, cantidad_etiquetas, 
                        cantidad_cajas, factura_guia, recepcionado_por
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """,
                    (fecha.strftime("%Y-%m-%d"), hora_despacho.strftime("%H:%M"), hora_llegada.strftime("%H:%M"),
                     conductor, orden_trabajo, cliente, int(cant_etiquetas), int(cant_cajas), doc, recepcionado_por)
                )
            conn.commit()
            conn.close()
            st.success("✅ Despacho guardado correctamente.")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al guardar: {e}")

# ==============================================================
# ⚙️ MÓDULO 2: DESCARGA DE REPORTE EXCEL CORPORATIVO
# ==============================================================
df_completo = pd.DataFrame()
if os.path.exists(DB_FILE):
    try:
        conn = sqlite3.connect(DB_FILE)
        df_completo = pd.read_sql_query("SELECT * FROM despachos ORDER BY fecha DESC, id DESC;", conn)
        conn.close()
    except Exception:
        pass

if not df_completo.empty:
    st.markdown("---")
    st.markdown("<h3 style='font-size: 20px; color: #2C3E50;'>📥 Descarga de Archivos</h3>", unsafe_allow_html=True)
    
    # 🎨 CONSTRUCCIÓN DEL EXCEL ELEGANTE CON OPENPYXL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Limpieza de columnas internas para el reporte corporativo
        df_excel = df_completo.drop(columns=["id"], errors="ignore")
        df_excel.columns = [col.replace('_', ' ').title() for col in df_excel.columns]
        
        # Exportar datos base dejando espacio arriba para el membrete corporativo
        df_excel.to_excel(writer, sheet_name='Bitácora Despachos', index=False, startrow=4)
        
        workbook = writer.book
        worksheet = workbook['Bitácora Despachos']
        
        # --- DISEÑO DE ESTILOS ---
        font_titulo = Font(name="Segoe UI", size=16, bold=True, color="1E3A8A")
        font_subtitulo = Font(name="Segoe UI", size=10, italic=True, color="555555")
        font_cabecera = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_cuerpo = Font(name="Segoe UI", size=11, color="000000")
        
        fill_cabecera = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        fill_cebra = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
        
        border_fino = Side(border_style="thin", color="D3D3D3")
        border_celda = Border(left=border_fino, right=border_fino, top=border_fino, bottom=border_fino)
        
        # --- APLICACIÓN DE ENCABEZADOS DE DISEÑO (Banners Superiores) ---
        worksheet['A1'] = "REPORTE CORPORATIVO DE DESPACHOS"
        worksheet['A1'].font = font_titulo
        
        worksheet['A2'] = f"Generado el: {fecha_actual.strftime('%d/%m/%Y')} | Registros totales: {len(df_excel)}"
        worksheet['A2'].font = font_subtitulo
        
        # Estilizar la fila de cabecera de la tabla (Fila 5)
        for col_num in range(1, len(df_excel.columns) + 1):
            cell = worksheet.cell(row=5, column=col_num)
            cell.fill = fill_cabecera
            cell.font = font_cabecera
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_celda
        
        # Estilizar el cuerpo de los datos (Desde la fila 6)
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=6, max_row=worksheet.max_row, min_col=1, max_col=len(df_excel.columns)), start=6):
            aplicar_cebra = (row_idx % 2 == 0)
            
            for cell in row:
                cell.font = font_cuerpo
                cell.border = border_celda
                if aplicar_cebra:
                    cell.fill = fill_cebra
                
                # Obtener el nombre de la columna para la alineación
                nombre_columna = df_excel.columns[cell.column - 1]
                
                # Alineación y formatos según el tipo de columna o dato
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif nombre_columna in ["Fecha", "Hora Despacho", "Hora Llegada"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
        # Configurar alturas de fila para dar "aire" visual
        worksheet.row_dimensions[5].height = 28 # Cabecera más alta
        for r in range(6, worksheet.max_row + 1):
            worksheet.row_dimensions[r].height = 20 # Filas de datos cómodas
            
        # Autoajustar el ancho de las columnas de forma inteligente
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col.column)
            for cell in col:
                if cell.row >= 5 and cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # Botón Destacado de Descarga
    st.download_button(
        label="📥 Descargar Reporte Excel Elegante",
