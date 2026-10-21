import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

# Archivo de la base de datos fija en la nube
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

# 🔑 CAMBIA TU CONTRASEÑA AQUÍ PARA MODIFICAR O BORRAR DATOS
PASSWORD_ADMIN = "admin123"
# ==============================================================

st.set_page_config(page_title="Bitácora de Despachos", page_icon="🚚", layout="centered")

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

st.markdown("<h2 style='font-size: 30px; margin-bottom: 0px;'>🚚 Bitácora de Despachos (Nube)</h2>", unsafe_allow_html=True)
st.write("Ingresa los datos para actualizar la base de datos en tiempo real.")

# Formulario básico limpio de kilómetros
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
    
    if conductor == "Selecciona el conductor o transporte...":
        st.error("⚠️ Por favor, selecciona un Conductor o Empresa de transporte válida.")
    elif cliente == "Selecciona un cliente...":
        st.error("⚠️ Por favor, selecciona un Cliente válido.")
    elif not lista_documentos:
        st.error("⚠️ Por favor, ingresa al menos una Factura o Guía.")
    elif recepcionado_por == "Selecciona quién recibe...":
        st.error("⚠️ Por favor, selecciona el receptor.")
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
            st.success(f"✅ ¡Despacho registrado en la nube!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error al guardar: {e}")

# === ⚙️ PANEL DE EDICIÓN Y BORRADO (MODO ADMINISTRADOR) ===
try:
    if os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        df_completo = pd.read_sql_query("SELECT * FROM despachos ORDER BY id DESC;", conn)
        conn.close()
        
        if not df_completo.empty:
            st.markdown("---")
            st.subheader("📋 Historial de despachos")
            
            # Formulario de contraseña oculto
            with st.expander("🛠️ Hacer cambios (Modificar o borrar registros)"):
                clave = st.text_input("Ingresa la clave de administrador para desbloquear la edición:", type="password")
                
                if clave == PASSWORD_ADMIN:
                    st.info("💡 Haz doble clic en una celda para editarla. Para borrar filas completas, selecciónalas marcando el cuadro de la izquierda y presiona 'Supr' en tu teclado.")
                    
                    # El editor interactivo de Streamlit
                    df_editado = st.data_editor(
                        df_completo, 
                        num_rows="dynamic", 
                        use_container_width=True,
                        key="editor_bd"
                    )
                    
                    boton_guardar_cambios = st.button("🔒 Guardar Cambios en la Nube")
                    
                    if boton_guardar_cambios:
                        try:
                            # Reemplazamos la tabla de la base de datos con los datos modificados
                            conn = sqlite3.connect(DB_FILE)
                            df_editado.to_sql("despachos", conn, if_exists="replace", index=False)
                            conn.commit()
                            conn.close()
                            st.success("✅ ¡Base de datos actualizada con éxito en la nube!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error al actualizar los cambios: {e}")
                elif clave != "":
                    st.error("🔑 Clave incorrecta.")
            
            # Vista normal estática para choferes (No editable)
            st.dataframe(df_completo.drop(columns=["id"], errors="ignore").head(15), use_container_width=True)
            
            # Botón de descarga de siempre usando los últimos datos guardados
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_completo.drop(columns=["id"], errors="ignore").to_excel(writer, index=False, sheet_name='Despachos')
            
            st.download_button(
                label="📥 Descargar Base de Datos Completa (Excel)",
                data=buffer.getvalue(),
                file_name="bitacora_despachos_nube.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
except Exception as e:
    st.error(f"Error de visualización: {e}")
