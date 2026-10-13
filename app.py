import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Librerías profesionales para la construcción de PDFs (ReportLab)
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

# ==============================================================
# 🔒 DATOS FIJOS DE TU EMPRESA (Escríbelos aquí dentro de las comillas)
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

# Lista unificada de tus conductores y transportes
LISTA_CONDUCTORES = [
    "Selecciona el conductor o transporte...",
    "ELIZABETH",              # <- FLOTA PROPIA (ACTIVA KILÓMETROS)
    "ELIZABETH - COLCHAGUA",  # <- FLOTA PROPIA (ACTIVA KILÓMETROS)
    "CHILEXPRESS",            # <- EXTERNO (OMITE KILÓMETROS)
    "STARKEN",                # <- EXTERNO (OMITE KILÓMETROS)
    "FEDEX",                  # <- EXTERNO (OMITE KILÓMETROS)
    "BLUE EXPRESS",           # <- EXTERNO (OMITE KILÓMETROS)
    "OTRO EXTERNO"
]
# ==============================================================

# Configuración elegante de la interfaz de la página
st.set_page_config(page_title="Generador de Informes de Despacho", page_icon="📄", layout="centered")

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

# 🎨 Bloque del Logo Corporativo
if os.path.exists("logo.png"):
    st.image("logo.png", width=180)
elif os.path.exists("logo.jpg"):
    st.image("logo.jpg", width=180)

st.markdown("<h2 style='font-size: 30px; margin-bottom: 0px;'>📄 Generador de Informes de Despacho PDF</h2>", unsafe_allow_html=True)
st.write("Completa el formulario para generar y descargar inmediatamente tu reporte listo para presentar.")

# Formulario de entrada de datos
with st.form(key="formulario_pdf", clear_on_submit=False):
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
        conductor = st.selectbox("Conductor / Empresa de Transporte", options=LISTA_CONDUCTORES)
        orden_trabajo = st.text_input("Orden de Trabajo")
        
    with col5:
        cliente = st.selectbox("Cliente (Destino)", options=LISTA_CLIENTES)
        facturas_guias_texto = st.text_area("Facturas o Guías (Escribe una por línea si son varias)")

    # Detección automática del tipo de viaje
    es_interno = (conductor == "Elizabeth")
    es_interno = (conductor == "Elizabeth - Colchagua")
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
        st.info(f"ℹ️ Transporte externo ({conductor}) seleccionado: Control de kilómetros omitido para el informe.")
        km_inicial = 0.0
        km_final = 0.0
    else:
        col8, col9 = st.columns(2)
        with col8:
            km_inicial = st.number_input("Kilómetros Iniciales", min_value=0.0, step=1.0, format="%.1f")
        with col9:
            km_final = st.number_input("Kilómetros Finales", min_value=0.0, step=1.0, format="%.1f")
    
    boton_preparar = st.form_submit_button(label="⚙️ Procesar Datos para PDF")

# Lógica de Validación y Creación Dinámica del PDF
if boton_preparar:
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
        
        # --- MOTOR DE GENERACIÓN DEL PDF EN MEMORIA ---
        buffer_pdf = io.BytesIO()
        doc = SimpleDocTemplate(buffer_pdf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36, title="Informe de Despacho")
        story = []
        styles = getSampleStyleSheet()
        
        # Estilos corporativos elegantes
        style_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontSize=20, leading=24, textColor=colors.HexColor('#1A365D'), alignment=1)
        style_sub = ParagraphStyle('SubPDF', parent=styles['Normal'], fontSize=10, textColor=colors.gray, alignment=1)
        style_tabla_header = ParagraphStyle('HeaderTabla', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
        style_tabla_celda = ParagraphStyle('CeldaTabla', parent=styles['Normal'], fontSize=9, alignment=1)
        
        # Encabezado del documento
        story.append(Paragraph("INFORME DE CONTROL LOGÍSTICO Y DESPACHOS", style_titulo))
        story.append(Spacer(1, 4))
        story.append(Paragraph(f"Documento generado el: {datetime.now().strftime('%d/%m/%Y a las %H:%M')}", style_sub))
        story.append(Spacer(1, 20))
        
        # Datos del formulario estructurados en formato de tabla ejecutiva
        datos_tabla = [
            [Paragraph("CONCEPTO / CAMPO", style_tabla_header), Paragraph("DETALLE DEL DESPACHO REGISTRADO", style_tabla_header)],
            [Paragraph("<b>FECHA DEL VIAJE</b>", style_tabla_celda), Paragraph(fecha.strftime("%d/%m/%Y"), style_tabla_celda)],
            [Paragraph("<b>HORARIO (DESPACHO / LLEGADA)</b>", style_tabla_celda), Paragraph(f"{hora_despacho.strftime('%H:%M')} hrs a {hora_llegada.strftime('%H:%M')} hrs", style_tabla_celda)],
            [Paragraph("<b>TRANSPORTISTA / CONDUCTOR</b>", style_tabla_celda), Paragraph(str(conductor), style_tabla_celda)],
            [Paragraph("<b>ORDEN DE TRABAJO (OT)</b>", style_tabla_celda), Paragraph(str(orden_trabajo) if orden_trabajo else "N/A", style_tabla_celda)],
            [Paragraph("<b>CLIENTE (REMITENTE)</b>", style_tabla_celda), Paragraph(str(cliente), style_tabla_celda)],
            [Paragraph("<b>DOCUMENTOS (FACTURAS / GUÍAS)</b>", style_tabla_celda), Paragraph(", ".join(lista_documentos), style_tabla_celda)],
            [Paragraph("<b>CARGA (ETIQUETAS / CAJAS)</b>", style_tabla_celda), Paragraph(f"{int(cant_etiquetas)} Etiquetas / {int(cant_cajas)} Cajas", style_tabla_celda)],
            [Paragraph("<b>RECEPCIONADO POR</b>", style_tabla_celda), Paragraph(str(recepcionado_por), style_tabla_celda)],
            [Paragraph("<b>CONTROL KILOMETRAJE</b>", style_tabla_celda), Paragraph(f"Inicial: {km_inicial:.1f} km | Final: {km_final:.1f} km (Recorridos: {km_recorridos:.1f} km)" if es_interno else "No aplica (Transporte Externo)", style_tabla_celda)],
        ]
        
        # Diseño de la tabla PDF (Azul Marino Ejecutivo)
        tabla = Table(datos_tabla, colWidths=[200, 320])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A365D')),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
        ]))
        
        story.append(tabla)
        story.append(Spacer(1, 50))
        
        # Bloque formal de firmas de cierre
        style_firma = ParagraphStyle('FirmaStyle', parent=styles['Normal'], fontSize=10, alignment=1)
        story.append(Paragraph("________________________________________", style_firma))
        story.append(Paragraph("<b>Firma Responsable Logística</b>", style_firma))
        
        # Construir PDF en la memoria RAM
        doc.build(story)
        pdf_data = buffer_pdf.getvalue()
        
        # Mostrar el cuadro de descarga de forma inmediata y destacada
        st.success("🎉 ¡Los datos del informe han sido procesados con éxito!")
        st.markdown("### 📥 Descarga tu Documento")
        st.download_button(
            label="📄 Guardar Informe de Despacho (PDF)",
            data=pdf_data,
            file_name=f"despacho_OT_{orden_trabajo if orden_trabajo else fecha.strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
