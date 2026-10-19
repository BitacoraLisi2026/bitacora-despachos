import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

# Nombres de los archivos
ARCHIVO_ORIGEN = "bitacora_despachos.xlsx"
ARCHIVO_INFORME = "Informe_Despachos_Presentacion.xlsx"
LOGO_FILE = "logo.png"  # Asegúrate de tener tu logo en la misma carpeta

def dar_formato_profesional():
    if not os.path.exists(ARCHIVO_ORIGEN):
        print(f"❌ Error: No se encuentra el archivo '{ARCHIVO_ORIGEN}' en esta carpeta.")
        return

    # Cargar el libro de Excel
    wb = openpyxl.load_workbook(ARCHIVO_ORIGEN)
    ws = wb.active
    ws.title = "Control de Despachos"

    # 1. Insertar filas al principio para el Encabezado y el Logo
    ws.insert_rows(1, amount=5)

    # Insertar el logo si existe
    if os.path.exists(LOGO_FILE):
        try:
            img = Image(LOGO_FILE)
            img.width = 140
            img.height = 70
            ws.add_image(img, 'A1')
        except Exception as e:
            print(f"⚠️ Nota con el logo: {e}")

    # Configurar el Título del Informe
    ws['D2'] = "INFORME CONTROL DE DESPACHOS Y LOGÍSTICA"
    ws['D2'].font = Font(name="Segoe UI", size=18, bold=True, color="1A1A1A") # Negro elegante
    
    ws['D3'] = "Reporte consolidado - Generado de forma automatizada"
    ws['D3'].font = Font(name="Segoe UI", size=10, italic=True, color="555555")

    # Lista de títulos optimizada a 10 columnas esenciales
    NUEVOS_TITULOS = (
        "FECHA", "HORA SALIDA", "HORA LLEGADA", "CONDUCTOR", "OT", 
        "CLIENTE", "CANT. ETIQUETAS", "CANT. CAJAS", "FACT / GUIA", 
        "RECEPCIONADO POR"
    )

    # 2. Definición de Paleta de Colores (Códigos HEX)
    NEGRO_HEADER = "1A1A1A"      # Fondo de títulos principal
    GRIS_ALTERNO = "F8FAFC"      # Filas alternas limpias
    AZUL_DESTACADO = "1E3A8A"    # Para resaltar texto externo
    AMARILLO_AVISO = "FEF08A"    # Para fletes externos
    TEXTO_BLANCO = "FFFFFF"

    # Fuentes y Bordes
    fuente_header = Font(name="Segoe UI", size=10, bold=True, color=TEXTO_BLANCO)
    fuente_celda = Font(name="Segoe UI", size=10, color="000000")
    fuente_externo = Font(name="Segoe UI", size=10, color="1E3A8A", bold=True)
    
    borde_fino = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # 3. Dar formato y renombrar los títulos (Fila 6)
    fila_headers = 6
    for col in range(1, len(NUEVOS_TITULOS) + 1):
        celda = ws.cell(row=fila_headers, column=col)
        celda.value = NUEVOS_TITULOS[col - 1]
            
        celda.fill = PatternFill(start_color=NEGRO_HEADER, end_color=NEGRO_HEADER, fill_type="solid")
        celda.font = fuente_header
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = borde_fino
    
    ws.row_dimensions[fila_headers].height = 28

    # 4. Formatear las filas de datos de la bitácora (Fila 7 en adelante)
    total_cajas = 0

    for row in range(7, ws.max_row + 1):
        ws.row_dimensions[row].height = 20
        es_fila_par = (row % 2 == 0)
        
        # Identificar si es flete externo leyendo el Conductor (Columna 4)
        nom_conductor = str(ws.cell(row=row, column=4).value).strip()
        es_flete_externo = (nom_conductor != "Elizabeth - Colchagua" and nom_conductor != "None" and nom_conductor != "")

        # Sumatoria de cajas (Columna 8)
        try:
            total_cajas += int(ws.cell(row=row, column=8).value or 0)
        except Exception:
            pass

        # Procesamos solo las 10 columnas visibles
        for col in range(1, len(NUEVOS_TITULOS) + 1):
            celda = ws.cell(row=row, column=col)
            celda.font = fuente_celda
            celda.border = borde_fino
            
            # Ajuste de alineaciones para las 10 columnas
            columnas_centro = (1, 2, 3, 5)  # FECHA, HORA SALIDA, HORA LLEGADA, OT
            columnas_derecha = (7, 8)       # CANT. ETIQUETAS, CANT. CAJAS

            if col in columnas_centro:
                celda.alignment = Alignment(horizontal="center", vertical="center")
            elif col in columnas_derecha:
                celda.alignment = Alignment(horizontal="right", vertical="center")
            else:
                celda.alignment = Alignment(horizontal="left", vertical="center")

            # Aplicar colores corporativos
            if es_flete_externo:
                celda.fill = PatternFill(start_color=AMARILLO_AVISO, end_color=AMARILLO_AVISO, fill_type="solid")
                if col == 4:
                    celda.font = fuente_externo
            elif es_fila_par:
                celda.fill = PatternFill(start_color=GRIS_ALTERNO, end_color=GRIS_ALTERNO, fill_type="solid")

    # Eliminar las columnas sobrantes del archivo final (Kilómetros y Tipo de Transporte)
    ws.delete_cols(11, amount=4)

    # 5. Agregar bloque de Resumen y Totales al final
    fila_totales = ws.max_row + 2
    ws.cell(row=fila_totales, column=2, value="RESUMEN OPERATIVO").font = Font(name="Segoe UI", size=11, bold=True, color=NEGRO_HEADER)
    
    ws.cell(row=fila_totales+1, column=2, value="Total Cajas Despachadas:").font = Font(name="Segoe UI", size=10, bold=True)
    ws.cell(row=fila_totales+1, column=3, value=total_cajas).font = Font(name="Segoe UI", size=10)

    # 6. Auto-ajustar el ancho de las 10 columnas (CORREGIDO PARA ITER_COLS)
    for col in ws.iter_cols(max_col=len(NUEVOS_TITULOS)):
        max_len = 0
        primera_celda = col[0]  # <- Extrae la primera celda real de la tupla
        col_letter = get_column_letter(primera_celda.column)
        for cell in col:
            if cell.row < 6:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Guardar el informe definitivo
    wb.save(ARCHIVO_INFORME)
    print(f"🎯 ¡Informe generado con éxito (10 columnas) como '{ARCHIVO_INFORME}'!")

if __name__ == "__main__":
    dar_formato_profesional()
