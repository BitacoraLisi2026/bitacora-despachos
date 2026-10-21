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

PASSWORD_ADMIN = "admin123"
# ==============================================================

st.set_page_config(page_title="Gestión de Despachos", page_icon="🚚", layout="wide") 

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

# --- AJUSTE DE HORA LOCAL (CHILE) ---
try:
    import zoneinfo
    zona_horaria = zoneinfo.ZoneInfo("America/Santiago") 
    hora_local_dt = datetime.now(zona_horaria)
except Exception:
    hora_local_dt = datetime.now()

fecha_actual = hora_local_dt.date()
hora_actual = hora_local_dt.time()

# ==============================================================
# 📊 MÓDULO 1: DASHBOARD EJECUTIVO EN LA APP
# ==============================================================
st.markdown("<h1 style='font-size: 32px; color: #1E3A8A; margin-bottom: 5px;'>📊 Panel de Gestión e Informe de Despachos</h1>", unsafe_allow_html=True)
st.write("Resumen ejecutivo del estado de las gestiones y distribución logística en tiempo real.")

df_informe = pd.DataFrame()
try:
    if os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        df_informe = pd.read_sql_query("SELECT * FROM despachos;", conn)
        conn.close()
        
        if not df_informe.empty:
            m1, m2, m3, m4 = st.columns(4)
            with m1: st.metric(label="📄 Total Documentos Procesados", value=f"{len(df_informe)} guías")
            with m2: st.metric(label="📦 Total Cajas Distribuidas", value=f"{int(df_informe['cantidad_cajas'].sum()):,}")
            with m3: st.metric(label="🏷️ Total Etiquetas Controladas", value=f"{int(df_informe['cantidad_etiquetas'].sum()):,}")
            with m4: st.metric(label="🏢 Clientes Atendidos", value=df_informe['cliente'].nunique())
except Exception:
    pass

st.markdown("---")

# ==============================================================
# 🚚 MÓDULO 2: INGRESO DE DATOS (FORMULARIO OPERATIVO)
# ==============================================================
st.markdown("<h2 style='font-size: 24px; margin-bottom: 0px;'>🚚 Registro de Nuevos Despachos</h2>", unsafe_allow_html=True)
st.write("Área operativa para la digitación en terreno desde dispositivos móviles o PC.")

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
            st.success("✅ ¡Despacho registrado en la nube!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al guardar: {e}")

# ==============================================================
# ⚙️ MÓDULO 3: VISTA DE REGISTROS Y EXPORTACIÓN CORPORATIVA EXCEL
# ==============================================================
df_completo = pd.DataFrame()
if os.path.exists(DB_FILE):
    try:
        conn = sqlite3.connect(DB_FILE)
        df_completo = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC;", conn)
        conn.close()
    except Exception:
        pass

if not df_completo.empty:
    st.markdown("---")
    st.subheader("📋 Últimos despachos registrados")
    
    with st.expander("🛠️ Panel de Administración (Modificar o borrar registros)"):
        clave = st.text_input("Ingresa la clave de administrador para desbloquear la edición:", type="password")
        if clave == PASSWORD_ADMIN:
            df_editado = st.data_editor(df_completo, num_rows="dynamic", use_container_width=True, key="editor_bd")
            if st.button("🔒 Guardar Cambios en la Nube"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    df_editado.to_sql("despachos", conn, if_exists="replace", index=False)
                    conn.commit()
                    conn.close()
                    st.success("✅ ¡Base de datos actualizada con éxito!")
                    st.rerun()
                except Exception as e: st.error(f"❌ Error: {e}")
        elif clave != "": st.error("🔑 Clave incorrecta.")
    
    df_visualizar = df_completo.drop(columns=["id"], errors="ignore")
    st.dataframe(df_visualizar.head(10), use_container_width=True)

    # 📥 GENERADOR EXCLUSIVO DE REPORTE EXCEL ELEGANTE (OPENPYXL)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Pestaña 2: Detalle de Datos base
        df_visualizar.to_excel(writer, sheet_name='Detalle Despachos', index=False)
        
        # Pestaña 1: Resumen Ejecutivo (Dashboard en Excel)
        workbook = writer.book
        ws_dash = workbook.create_sheet(title="Resumen Ejecutivo", index=0)
        ws_dash.views.sheetView[0].showGridLines = True
        ws_det = writer.sheets['Detalle Despachos']
        ws_det.views.sheetView[0].showGridLines = True

        # Estilos Corporativos (Azul Marino y Gris Ejecutivo)
        font_titulo = Font(name='Segoe UI', size=16, bold=True, color='1E3A8A')
        font_subtitulo = Font(name='Segoe UI', size=11, italic=True, color='4B5563')
        font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        font_kpi_num = Font(name='Segoe UI', size=18, bold=True, color='1E3A8A')
        font_kpi_lbl = Font(name='Segoe UI', size=9, bold=False, color='374151')
        font_normal = Font(name='Segoe UI', size=11)
        
        fill_azul = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
        fill_kpi = PatternFill(start_color='F3F4F6', end_color='F3F4F6', fill_type='solid')
        
        align_center = Alignment(horizontal='center', vertical='center')
        align_left = Alignment(horizontal='left', vertical='center')
        align_right = Alignment(horizontal='right', vertical='center')
        
        borde_delgado = Border(
            left=Side(style='thin', color='D1D5DB'), right=Side(style='thin', color='D1D5DB'),
            top=Side(style='thin', color='D1D5DB'), bottom=Side(style='thin', color='D1D5DB')
        )

        # ---- DISEÑO PESTAÑA: RESUMEN EJECUTIVO ----
        ws_dash['B2'] = "INFORME GERENCIAL DE GESTIÓN LOGÍSTICA"
        ws_dash['B2'].font = font_titulo
        ws_dash['B3'] = f"Reporte automatizado en tiempo real • Generado el {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ws_dash['B3'].font = font_subtitulo

