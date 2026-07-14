import streamlit as st
import pdfplumber
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sqlite3
import hashlib
import os
import datetime
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Cartas SST - JER S.A.", page_icon="🩺", layout="wide")

# Estilos personalizados para una interfaz hermosa y moderna
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button {
        background-color: #1f4e79;
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: bold;
        border: none;
    }
    .stButton>button:hover { background-color: #153552; color: white; }
    .login-container {
        max-width: 450px;
        margin: 80px auto;
        padding: 30px;
        background-color: white;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

# --- SISTEMA DE SEGURIDAD (BASE DE DATOS SQLITE) ---
DB_NAME = "usuarios.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios 
                 (usuario TEXT PRIMARY KEY, contrasena TEXT, nombre TEXT)''')
    conn.commit()
    conn.close()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def registrar_usuario(user, pwd, name):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute("INSERT INTO usuarios VALUES (?, ?, ?)", (user, hash_password(pwd), name))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    return success

def verificar_usuario(user, pwd):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT nombre FROM usuarios WHERE usuario = ? AND contrasena = ?", (user, hash_password(pwd)))
    resultado = c.fetchone()
    conn.close()
    return resultado[0] if resultado else None

def tiene_usuarios():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM usuarios")
    count = c.fetchone()[0]
    conn.close()
    return count > 0

init_db()

# --- MANEJO DE SESIÓN DE LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

# --- PANTALLAS DE ACCESO ---
if not st.session_state.logged_in:
    if not tiene_usuarios():
        # Registro inicial (solo se muestra la primera vez)
        st.markdown("<div class='login-container'>", unsafe_allow_html=True)
        st.subheader("🔑 Crear Cuenta de Administrador")
        st.info("No hay usuarios registrados. Configura tu cuenta inicial para proteger el sistema.")
        reg_nombre = st.text_input("Nombre Completo")
        reg_user = st.text_input("Nombre de Usuario (para ingresar)")
        reg_pwd = st.text_input("Contraseña", type="password")
        if st.button("Crear Administrador"):
            if reg_nombre and reg_user and reg_pwd:
                if registrar_usuario(reg_user, reg_pwd, reg_nombre):
                    st.success("¡Usuario creado con éxito! Ya puedes iniciar sesión.")
                    st.rerun()
                else:
                    st.error("El usuario ya existe.")
            else:
                st.warning("Completa todos los campos.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    else:
        # Login normal
        st.markdown("<div class='login-container'>", unsafe_allow_html=True)
        st.subheader("🩺 Ingreso al Sistema - JER S.A.")
        log_user = st.text_input("Usuario")
        log_pwd = st.text_input("Contraseña", type="password")
        if st.button("Iniciar Sesión"):
            nombre_usuario = verificar_usuario(log_user, log_pwd)
            if nombre_usuario:
                st.session_state.logged_in = True
                st.session_state.username = nombre_usuario
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

# --- EXTRACCIÓN INTELIGENTE DE TEXTO DESDE EL PDF ---
def analizar_pdf_basico(texto):
    """Intenta extraer datos comunes usando búsquedas de patrones en el PDF"""
    datos = {
        "nombre": "", "cargo": "", "tipo_examen": "",
        "examenes": "", "recomendaciones": "", "vigilancia": "",
        "observaciones": "", "remisiones": ""
    }
    if not texto:
        return datos

    # Búsqueda de Nombre
    m_nom = re.search(r'(?:Nombre|Paciente|Colaborador|Trabajador):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_nom: datos["nombre"] = m_nom.group(1).strip()
    
    # Búsqueda de Cargo
    m_car = re.search(r'(?:Cargo|Ocupacion|Puesto):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_car: datos["cargo"] = m_car.group(1).strip()
    
    # Búsqueda de Tipo de Examen (Ingreso, Periódico, etc.)
    m_tipo = re.search(r'(?:Tipo de Examen|Concepto|Evaluacion):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_tipo: datos["tipo_examen"] = m_tipo.group(1).strip()

    # Búsqueda de Recomendaciones
    m_rec = re.search(r'(?:Recomendaciones|Indicaciones):\s*(.*?)(?:Observaciones|Remisiones|SST|$)', texto, re.IGNORECASE | re.DOTALL)
    if m_rec: datos["recomendaciones"] = m_rec.group(1).strip()

    # Búsqueda de Observaciones
    m_obs = re.search(r'(?:Observaciones|Nota):\s*(.*?)(?:Remisiones|SST|$)', texto, re.IGNORECASE | re.DOTALL)
    if m_obs: datos["observaciones"] = m_obs.group(1).strip()

    return datos

# --- INTERFAZ DE TRABAJO PRINCIPAL (LOGUEADO) ---
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.username}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("✍️ Cargar Firma Autorizada")
firma_file = st.sidebar.file_uploader("Sube la imagen de la firma (.png / .jpg)", type=["png", "jpg"])

# --- DIVISION DE PANTALLA EN 2 COLUMNAS ---
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.subheader("📁 1. Cargar PDF Médico")
    pdf_subido = st.file_uploader("Arrastra aquí el examen del proveedor", type="pdf")
    
    texto_raw = ""
    datos_sugeridos = {
        "nombre": "", "cargo": "", "tipo_examen": "",
        "examenes": "", "recomendaciones": "", "vigilancia": "",
        "observaciones": "", "remisiones": ""
    }
    
    if pdf_subido:
        with pdfplumber.open(pdf_subido) as pdf:
            for page in pdf.pages:
                texto_raw += page.extract_text() + "\n"
        
        st.info("✅ PDF leído con éxito.")
        datos_sugeridos = analizar_pdf_basico(texto_raw)
        
        with st.expander("🔍 Ver texto extraído del PDF"):
            st.text_area("Texto crudo:", value=texto_raw, height=300)

with col_der:
    st.subheader("📋 2. Formato de Envío (Datos Finales)")
    st.write("Verifica y edita la información que irá directamente al documento de Word:")
    
    # Campos del formulario con valores sugeridos del PDF
    consecutivo = st.text_input("Consecutivo de la Carta:", placeholder="Ej: JER-SST-001-2026")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        lugar = st.text_input("Lugar de Expedición:", value="Tunja")
    with col_f2:
        fecha = st.date_input("Fecha de la Carta:", value=datetime.date(2026, 7, 14)) # Cambiar al año actual 2026

    tipo_examen = st.text_input("Tipo de Examen (ASUNTO):", value=datos_sugeridos["tipo_examen"], placeholder="Ej: PERIÓDICO")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        nombre_persona = st.text_input("Nombre del Trabajador:", value=datos_sugeridos["nombre"])
    with col_p2:
        cargo_persona = st.text_input("Cargo del Trabajador:", value=datos_sugeridos["cargo"])
        
    examenes_realizados = st.text_area("Exámenes Realizados:", value=datos_sugeridos["examenes"], placeholder="Ej: Visiometría, Audiometría, Cuadro Hemático...")
    
    recom_medicas = st.text_area("Recomendaciones Médicas:", value=datos_sugeridos["recomendaciones"])
    vigilancia = st.text_area("Programa de Vigilancia Epidemiológica (PVE):", value=datos_sugeridos["vigilancia"], placeholder="Ej: Conservación Auditiva / Ninguno")
    
    observaciones = st.text_area("Observaciones:", value=datos_sugeridos["observaciones"])
    remisiones = st.text_area("Remisiones:", value=datos_sugeridos["remisiones"])

    # --- GENERAR DOCUMENTO WORD ---
    if st.button("💾 Generar y Descargar Word"):
        try:
            doc = Document()
            
            # Configurar fuente Arial 11 por defecto
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Arial'
            font.size = Pt(11)
            
            # Margenes estándar
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # 1. Consecutivo (Alineado a la Derecha)
            p_cons = doc.add_paragraph()
            p_cons.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run_cons = p_cons.add_run(f"Consecutivo: {consecutivo}")
            run_cons.bold = True
            
            doc.add_paragraph() # Espacio
            
            # 2. Asunto (Centrado y en Negrita)
            p_asunto = doc.add_paragraph()
            p_asunto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_asunto = p_asunto.add_run(f"ASUNTO: RECOMENDACIONES EXAMEN {tipo_examen.upper()}")
            run_asunto.bold = True
            
            doc.add_paragraph() # Espacio
            
            # 3. Lugar y Fecha
            p_fecha = doc.add_paragraph()
            fecha_formateada = fecha.strftime("%d de %B de %Y")
            # Traducir meses al español básico
            meses = {
                "January": "enero", "February": "febrero", "March": "marzo", "April": "abril",
                "May": "mayo", "June": "junio", "July": "julio", "August": "agosto",
                "September": "septiembre", "October": "octubre", "November": "noviembre", "December": "diciembre"
            }
            for eng, esp in meses.items():
                fecha_formateada = fecha_formateada.replace(eng, esp)
                
            p_fecha.add_run(f"{lugar}, {fecha_formateada}")
            
            doc.add_paragraph() # Espacio
            
            # 4. Destinatario
            doc.add_paragraph("Sr(a).")
            p_nom = doc.add_paragraph()
            p_nom.add_run(nombre_persona).bold = True
            doc.add_paragraph(cargo_persona)
            
            doc.add_paragraph() # Espacio
            
            # 5. Saludo y texto institucional
            p_saludo = doc.add_paragraph("Cordial saludo,")
            p_cuerpo = doc.add_paragraph(
                "Según los lineamientos del programa de medicina preventiva y del trabajo de JER S.A; "
                "se hace entrega de las recomendaciones establecidas por el Proveedor de servicios de "
                "Exámenes Médico Ocupacionales (Ingreso, Periódico, egreso, cambio de cargo y post incapacidad)"
            )
            p_cuerpo.paragraph_format.line_spacing = 1.15
            
            doc.add_paragraph() # Espacio
            
            # 6. Exámenes Realizados
            p_ex = doc.add_paragraph()
            p_ex.add_run("EXÁMENES REALIZADOS:").bold = True
            doc.add_paragraph(examenes_realizados if examenes_realizados else "Ninguno registrado.")
            
            doc.add_paragraph() # Espacio
            
            # 7. Recomendaciones y PVE
            p_recom = doc.add_paragraph()
            p_recom.add_run("Recomendaciones: ").bold = True
            p_recom.add_run(recom_medicas if recom_medicas else "No registra.")
            
            if vigilancia:
                p_recom.add_run("\nPrograma de vigilancia epidemiológica: ").bold = True
                p_recom.add_run(vigilancia)
                
            # 8. Observaciones
            p_obs = doc.add_paragraph()
            p_obs.add_run("observaciones: ").bold = True
            p_obs.add_run(observaciones if observaciones else "Ninguna.")
            
            # 9. Remisiones
            p_rem = doc.add_paragraph()
            p_rem.add_run("remisiones: ").bold = True
            p_rem.add_run(remisiones if remisiones else "Ninguna.")
            
            doc.add_paragraph() # Espacio
            
            # 10. Firma Institucional
            doc.add_paragraph("Atentamente,")
            
            if firma_file:
                # Estampar firma si fue subida
                doc.add_picture(firma_file, width=Inches(2.2))
            else:
                doc.add_paragraph() # Espacio vacío si no hay firma cargada
                
            p_firma_nombre = doc.add_paragraph()
            p_firma_nombre.add_run("VÍCTOR ALONSO MORENO CASAS").bold = True
            doc.add_paragraph("Coordinador SST")
            
            # Guardar en memoria
            b_io = io.BytesIO()
            doc.save(b_io)
            b_io.seek(0)
            
            st.success("🎉 ¡Documento de Word generado correctamente!")
            st.download_button(
                label="📥 Descargar Documento Word (.docx)",
                data=b_io,
                file_name=f"Recomendaciones_{nombre_persona.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Error al estructurar el Word: {e}")
