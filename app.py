import os
from datetime import datetime, timedelta, timezone

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------
st.set_page_config(page_title="Registro de Entregas", page_icon="🚚", layout="centered")

ARCHIVO_EXCEL = "registro_entregas.xlsx"
NOMBRE_HOJA = "Entregas"

# El logo debe estar en la misma carpeta que app.py (en la raíz del repositorio)
RUTA_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

ENCABEZADOS = [
    "Fecha de salida",
    "Hora de entrega",
    "Cliente",
    "Quién recepciona",
    "Conductor",
    "Transporte",
    "Guía / Factura",
]

CLIENTES_BASE = [
    "Viñedos de Aguirre",
    "Viña Casadonoso",
    "Viña Aromo",
    "Viña Casa Solis",
    "Viña Concha y Toro",
    "Viña Luis Felipe Edwards",
    "Transportes Colchagua",
]

RECEPCIONISTAS_BASE = [
    "Camila Villasana",
    "Olga Gonzalez",
    "Jorge Diaz",
    "Javiera Ramirez",
    "Encargado Colchagua",
]

EMPRESAS_TRANSPORTE_BASE = [
    "Starken",
    "Chilexpress",
    "Correos de Chile",
    "Blue Express",
    "Turbus Cargo",
    "Pullman Cargo",
]

OPCION_NUEVO = "➕ Agregar nuevo"
PLACEHOLDER_CLIENTE = "Seleccione cliente..."
PLACEHOLDER_RECEPCION = "Seleccione quién recepciona..."
OPCION_OTRA_EMPRESA = "Otra Empresa"
VEHICULO_INTERNO = "Vehículo Interno"

# Colores del diseño corporativo
COLOR_AZUL_MARINO = "1F3864"
COLOR_GRIS_CLARO = "F2F2F2"
COLOR_BLANCO = "FFFFFF"
COLOR_BORDE = "BFBFBF"


# ----------------------------------------------------------------------
# UTILIDADES
# ----------------------------------------------------------------------
def ahora_chile():
    """Devuelve la fecha y hora actual de Chile (el servidor de Streamlit Cloud usa UTC)."""
    try:
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo("America/Santiago"))
    except Exception:
        return datetime.now(timezone(timedelta(hours=-3)))


def aplicar_diseno(ws):
    """Aplica el diseño corporativo: cabecera azul marino y filas zebra gris claro."""
    borde = Side(style="thin", color=COLOR_BORDE)
    borde_celda = Border(left=borde, right=borde, top=borde, bottom=borde)

    relleno_cabecera = PatternFill(
        start_color=COLOR_AZUL_MARINO, end_color=COLOR_AZUL_MARINO, fill_type="solid"
    )
    relleno_gris = PatternFill(
        start_color=COLOR_GRIS_CLARO, end_color=COLOR_GRIS_CLARO, fill_type="solid"
    )
    relleno_blanco = PatternFill(
        start_color=COLOR_BLANCO, end_color=COLOR_BLANCO, fill_type="solid"
    )

    fuente_cabecera = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    fuente_normal = Font(name="Calibri", size=11, color="000000")

    # Cabecera
    for celda in ws[1]:
        celda.fill = relleno_cabecera
        celda.font = fuente_cabecera
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = borde_celda
    ws.row_dimensions[1].height = 28

    # Filas de datos con efecto zebra
    for idx_fila in range(2, ws.max_row + 1):
        relleno = relleno_gris if idx_fila % 2 == 0 else relleno_blanco
        for celda in ws[idx_fila]:
            celda.fill = relleno
            celda.font = fuente_normal
            celda.border = borde_celda
            celda.alignment = Alignment(horizontal="center", vertical="center")
        # Formato de fecha en la columna A
        ws.cell(row=idx_fila, column=1).number_format = "DD/MM/YYYY"
        ws.row_dimensions[idx_fila].height = 20

    # Ancho de columnas según contenido
    for idx_col in range(1, ws.max_column + 1):
        letra = get_column_letter(idx_col)
        largo_max = 0
        for celda in ws[letra]:
            if celda.value is not None:
                valor = (
                    celda.value.strftime("%d/%m/%Y")
                    if hasattr(celda.value, "strftime") and idx_col == 1
                    else str(celda.value)
                )
                largo_max = max(largo_max, len(valor))
        ws.column_dimensions[letra].width = min(max(largo_max + 4, 16), 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def crear_libro_nuevo():
    wb = Workbook()
    ws = wb.active
    ws.title = NOMBRE_HOJA
    ws.append(ENCABEZADOS)
    return wb


def abrir_libro():
    """Abre el Excel existente o crea uno nuevo si aún no existe."""
    if os.path.exists(ARCHIVO_EXCEL):
        try:
            wb = load_workbook(ARCHIVO_EXCEL)
            if NOMBRE_HOJA in wb.sheetnames:
                return wb, wb[NOMBRE_HOJA]
            ws = wb.create_sheet(NOMBRE_HOJA)
            ws.append(ENCABEZADOS)
            return wb, ws
        except Exception:
            pass
    wb = crear_libro_nuevo()
    return wb, wb[NOMBRE_HOJA]


def guardar_filas(filas):
    """Agrega las filas al Excel (una por guía) y reaplica el diseño."""
    wb, ws = abrir_libro()
    for fila in filas:
        ws.append(fila)
    aplicar_diseno(ws)
    wb.save(ARCHIVO_EXCEL)


def valores_guardados(indice_columna):
    """Lee valores únicos ya guardados en el Excel (para recordar 'nuevos' agregados antes)."""
    valores = []
    if not os.path.exists(ARCHIVO_EXCEL):
        return valores
    try:
        wb = load_workbook(ARCHIVO_EXCEL, read_only=True)
        if NOMBRE_HOJA not in wb.sheetnames:
            return valores
        ws = wb[NOMBRE_HOJA]
        for fila in ws.iter_rows(min_row=2, values_only=True):
            if len(fila) > indice_columna and fila[indice_columna]:
                valor = str(fila[indice_columna]).strip()
                if valor and valor not in valores:
                    valores.append(valor)
        wb.close()
    except Exception:
        pass
    return valores


def construir_opciones(base, extras_sesion, indice_columna, excluir=()):
    """Une la lista base + lo guardado en el Excel + lo agregado en esta sesión."""
    opciones = list(base)
    for valor in valores_guardados(indice_columna) + list(extras_sesion):
        if valor not in opciones and valor not in excluir:
            opciones.append(valor)
    return opciones


def cargar_tabla():
    if not os.path.exists(ARCHIVO_EXCEL):
        return pd.DataFrame(columns=ENCABEZADOS)
    try:
        wb = load_workbook(ARCHIVO_EXCEL, read_only=True)
        ws = wb[NOMBRE_HOJA]
        filas = list(ws.iter_rows(min_row=2, values_only=True))
        wb.close()
    except Exception:
        return pd.DataFrame(columns=ENCABEZADOS)

    df = pd.DataFrame(filas, columns=ENCABEZADOS)
    if not df.empty:
        df["Fecha de salida"] = df["Fecha de salida"].apply(
            lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else x
        )
    return df


# ----------------------------------------------------------------------
# ESTADO DE SESIÓN
# ----------------------------------------------------------------------
if "clientes_extra" not in st.session_state:
    st.session_state["clientes_extra"] = []
if "recepcionistas_extra" not in st.session_state:
    st.session_state["recepcionistas_extra"] = []
if "empresas_extra" not in st.session_state:
    st.session_state["empresas_extra"] = []
if "mensaje" not in st.session_state:
    st.session_state["mensaje"] = None


# ----------------------------------------------------------------------
# LÓGICA DE GUARDADO (se ejecuta al presionar el botón)
# ----------------------------------------------------------------------
def guardar_registro():
    ss = st.session_state

    # Cliente
    cliente = ss.get("cliente_sel")
    if cliente == OPCION_NUEVO:
        cliente = (ss.get("cliente_nuevo") or "").strip()
    elif cliente == PLACEHOLDER_CLIENTE:
        cliente = ""

    # Quién recepciona
    recepciona = ss.get("recepciona_sel")
    if recepciona == OPCION_NUEVO:
        recepciona = (ss.get("recepciona_nuevo") or "").strip()
    elif recepciona == PLACEHOLDER_RECEPCION:
        recepciona = ""

    # Conductor y transporte
    conductor = ss.get("conductor")
    if conductor == "Elizabeth":
        transporte = VEHICULO_INTERNO
    else:
        transporte = ss.get("transporte_sel")
        if transporte == OPCION_OTRA_EMPRESA:
            transporte = (ss.get("transporte_nuevo") or "").strip()

    # Guías: una por línea
    texto_guias = ss.get("guias_texto") or ""
    guias = [linea.strip() for linea in texto_guias.splitlines() if linea.strip()]

    # Validaciones
    errores = []
    if not cliente:
        errores.append("Debes indicar el cliente.")
    if not recepciona:
        errores.append("Debes indicar quién recepciona.")
    if not transporte:
        errores.append("Debes indicar la empresa de transporte.")
    if not guias:
        errores.append("Debes ingresar al menos una guía o factura.")

    if errores:
        ss["mensaje"] = ("error", " ".join(errores))
        return

    fecha_salida = ss.get("fecha_salida")
    ahora = ahora_chile()

    # Una fila independiente por cada guía, con su hora de entrega automática
    filas = []
    for guia in guias:
        hora_entrega = ahora_chile().strftime("%H:%M:%S")
        filas.append(
            [fecha_salida, hora_entrega, cliente, recepciona, conductor, transporte, guia]
        )

    try:
        guardar_filas(filas)
    except Exception as e:
        ss["mensaje"] = ("error", f"No se pudo guardar el Excel: {e}")
        return

    # Recordar nuevos valores en la sesión
    if ss.get("cliente_sel") == OPCION_NUEVO and cliente not in ss["clientes_extra"]:
        ss["clientes_extra"].append(cliente)
    if ss.get("recepciona_sel") == OPCION_NUEVO and recepciona not in ss["recepcionistas_extra"]:
        ss["recepcionistas_extra"].append(recepciona)
    if (
        conductor == "Empresa Externa"
        and ss.get("transporte_sel") == OPCION_OTRA_EMPRESA
        and transporte not in ss["empresas_extra"]
    ):
        ss["empresas_extra"].append(transporte)

    # Limpiar campos para el siguiente registro
    ss["guias_texto"] = ""
    ss["cliente_nuevo"] = ""
    ss["recepciona_nuevo"] = ""
    ss["transporte_nuevo"] = ""
    ss["cliente_sel"] = PLACEHOLDER_CLIENTE
    ss["recepciona_sel"] = PLACEHOLDER_RECEPCION

    plural = "guía" if len(filas) == 1 else "guías"
    ss["mensaje"] = (
        "ok",
        f"Se guardaron {len(filas)} {plural} a las {ahora.strftime('%H:%M:%S')} hrs.",
    )


# ----------------------------------------------------------------------
# INTERFAZ
# ----------------------------------------------------------------------
if os.path.exists(RUTA_LOGO):
    col_izq, col_centro, col_der = st.columns([1, 2, 1])
    with col_centro:
        st.image(RUTA_LOGO, use_container_width=True)

st.title("🚚 Registro de Entregas")
st.caption("Completa los datos y guarda. Cada guía se registra en una fila independiente del Excel.")

# Mensaje del último guardado
if st.session_state["mensaje"]:
    tipo, texto = st.session_state["mensaje"]
    if tipo == "ok":
        st.success(texto)
    else:
        st.error(texto)
    st.session_state["mensaje"] = None

# 1. Fecha de salida
st.date_input("Fecha de salida", value=ahora_chile().date(), format="DD/MM/YYYY", key="fecha_salida")

# 2. Cliente
opciones_clientes = (
    [PLACEHOLDER_CLIENTE]
    + construir_opciones(CLIENTES_BASE, st.session_state["clientes_extra"], indice_columna=2)
    + [OPCION_NUEVO]
)
st.selectbox("Cliente", opciones_clientes, key="cliente_sel")
if st.session_state.get("cliente_sel") == OPCION_NUEVO:
    st.text_input("Nombre del nuevo cliente", key="cliente_nuevo")

# 3. Quién recepciona
opciones_recep = (
    [PLACEHOLDER_RECEPCION]
    + construir_opciones(
        RECEPCIONISTAS_BASE, st.session_state["recepcionistas_extra"], indice_columna=3
    )
    + [OPCION_NUEVO]
)
st.selectbox("Quién recepciona", opciones_recep, key="recepciona_sel")
if st.session_state.get("recepciona_sel") == OPCION_NUEVO:
    st.text_input("Nombre de quien recepciona", key="recepciona_nuevo")

# 4. Conductor
st.radio("Conductor", ["Elizabeth", "Empresa Externa"], horizontal=True, key="conductor")

# 5. Transporte
if st.session_state.get("conductor") == "Elizabeth":
    st.info(f"Transporte asignado automáticamente: **{VEHICULO_INTERNO}**")
else:
    opciones_transporte = construir_opciones(
        EMPRESAS_TRANSPORTE_BASE,
        st.session_state["empresas_extra"],
        indice_columna=5,
        excluir=(VEHICULO_INTERNO,),
    ) + [OPCION_OTRA_EMPRESA]
    st.selectbox("Empresa de transporte", opciones_transporte, key="transporte_sel")
    if st.session_state.get("transporte_sel") == OPCION_OTRA_EMPRESA:
        st.text_input("Nombre de la otra empresa", key="transporte_nuevo")

# 6. Guías / Facturas
st.text_area(
    "Guías / Facturas",
    height=200,
    placeholder="Escribe una guía por línea y presiona ENTER para la siguiente.\nEj:\n12345\n12346\n12347",
    key="guias_texto",
)

# Contador de guías detectadas
cantidad_guias = len([l for l in (st.session_state.get("guias_texto") or "").splitlines() if l.strip()])
if cantidad_guias:
    st.caption(f"Guías detectadas: {cantidad_guias}")

# 7. Guardar
st.button("💾 Guardar registro", type="primary", on_click=guardar_registro, use_container_width=True)

# ----------------------------------------------------------------------
# HISTORIAL Y DESCARGA
# ----------------------------------------------------------------------
st.divider()
st.subheader("📋 Registros guardados")

tabla = cargar_tabla()
if tabla.empty:
    st.write("Aún no hay registros.")
else:
    st.dataframe(tabla.iloc[::-1], use_container_width=True, hide_index=True)

if os.path.exists(ARCHIVO_EXCEL):
    with open(ARCHIVO_EXCEL, "rb") as f:
        st.download_button(
            "⬇️ Descargar Excel",
            data=f.read(),
            file_name="registro_entregas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.caption(
        "⚠️ En Streamlit Cloud el archivo en el servidor se reinicia cuando la app se reinicia "
        "o se vuelve a desplegar. Descarga el Excel con frecuencia."
    )
