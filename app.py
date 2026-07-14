import streamlit as st
import pdfplumber
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import sqlite3
import hashlib
import os
import datetime
import re
import requests
import io
import zipfile
import tempfile

# Intentar importar FPDF para generación de PDF nativo
try:
    from fpdf import FPDF
    fpdf_disponible = True
except ImportError:
    fpdf_disponible = False

# --- CONFIGURACIÓN DE PÁGINA AVANZADA ---
st.set_page_config(
    page_title="Portal SST - JER S.A.", 
    page_icon="🩺", 
    layout="wide"
)

# --- INYECCIÓN DE CSS: MODO OSCURO PREMIUM ---
st.markdown("""
    <style>
    html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], [data-testid="stHeader"] {
        background-color: #0b0f19 !important;
        color: #f8fafc !important;
    }
    
    .login-box {
        max-width: 450px;
        margin: 80px auto;
        padding: 40px;
        background-color: #111827 !important;
        border-radius: 16px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
        border: 1px solid #1f2937;
    }
    
    .login-box h2 {
        color: #3b82f6 !important;
        text-align: center;
        margin-bottom: 5px;
        font-weight: 700;
    }
    .login-box p {
        color: #9ca3af !important;
        text-align: center;
        font-size: 14px;
    }
    
    div[data-testid="stRadio"] label p {
        color: #f3f4f6 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }
    
    div[data-testid="stWidgetLabel"] p {
        color: #60a5fa !important;
        font-weight: 600 !important;
        font-size: 15px !important;
    }
    
    div[data-baseweb="input"] {
        background-color: #1f2937 !important;
        border: 1px solid #374151 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="input"] input {
        color: #ffffff !important;
        background-color: #1f2937 !important;
    }
    div[data-testid="stTextArea"] textarea {
        color: #ffffff !important;
        background-color: #1f2937 !important;
        border: 1px solid #374151 !important;
        border-radius: 8px !important;
    }
    
    button[data-baseweb="tab"] p {
        color: #9ca3af !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p {
        color: #3b82f6 !important;
        font-weight: 700 !important;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        font-weight: 700 !important;
        border: none !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.3) !important;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.5) !important;
        filter: brightness(115%);
    }
    
    .header-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #0f172a 100%);
        padding: 22px;
        border-radius: 12px;
        color: #ffffff !important;
        margin-bottom: 25px;
        border: 1px solid #1e2937;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    
    .metric-card {
        background-color: #111827 !important;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        border-left: 5px solid #2563eb;
        margin-bottom: 12px;
        color: #e5e7eb !important;
    }
    
    div[data-testid="stExpander"] {
        background-color: #111827 !important;
        border: 1px solid #1f2937 !important;
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- SISTEMA DE SEGURIDAD (BASE DE DATOS) ---
DB_NAME = "usuarios.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios 
                 (usuario TEXT PRIMARY KEY, contrasena TEXT, nombre TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS configuracion 
                 (clave TEXT PRIMARY KEY, valor TEXT)''')
    conn.commit()
    conn.close()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def registrar_usuario(user, pwd, name):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute("INSERT INTO usuarios VALUES (?, ?, ?)", (user.lower().strip(), hash_password(pwd), name.strip()))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    return success

def verificar_usuario(user, pwd):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT nombre FROM usuarios WHERE usuario = ? AND contrasena = ?", (user.lower().strip(), hash_password(pwd)))
    resultado = c.fetchone()
    conn.close()
    return resultado[0] if resultado else None

def actualizar_contrasena(user, old_pwd, new_pwd):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT contrasena FROM usuarios WHERE usuario = ?", (user.lower().strip(),))
    resultado = c.fetchone()
    if resultado and resultado[0] == hash_password(old_pwd):
        c.execute("UPDATE usuarios SET contrasena = ? WHERE usuario = ?", (hash_password(new_pwd), user.lower().strip()))
        conn.commit()
        conn.close()
        return True
    conn.close()
    return False

def tiene_usuarios():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM usuarios")
    count = c.fetchone()[0]
    conn.close()
    return count > 0

def guardar_config(clave, valor):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO configuracion VALUES (?, ?)", (clave, valor))
    conn.commit()
    conn.close()

def obtener_config(clave):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT valor FROM configuracion WHERE clave = ?", (clave,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

init_db()

# --- MANEJO DE SESIÓN DE LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "documentos" not in st.session_state:
    st.session_state.documentos = {}
if "textos_raw" not in st.session_state:
    st.session_state.textos_raw = {}
if "export_bytes" not in st.session_state:
    st.session_state.export_bytes = None
if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None

# --- PANTALLAS DE ACCESO ---
if not st.session_state.logged_in:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.markdown("<h2>🔑 Acceso Seguro</h2>", unsafe_allow_html=True)
    st.markdown("<p>Portal Interno de Medicina Preventiva - JER S.A.</p>", unsafe_allow_html=True)
    
    if not tiene_usuarios():
        st.warning("🆕 Bienvenido. Configura tu cuenta inicial de Administrador.")
        reg_nombre = st.text_input("Nombre Completo")
        reg_user = st.text_input("Nombre de Usuario (Login)")
        reg_pwd = st.text_input("Contraseña", type="password")
        if st.button("Crear Administrador"):
            if reg_nombre and reg_user and reg_pwd:
                if registrar_usuario(reg_user, reg_pwd, reg_nombre):
                    st.success("¡Administrador creado con éxito!")
                    st.rerun()
            else:
                st.warning("Completa todos los campos.")
    else:
        opcion_acceso = st.radio("Elige una acción:", ["Iniciar Sesión", "Crear Nueva Cuenta", "Actualizar Contraseña"], horizontal=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if opcion_acceso == "Iniciar Sesión":
            log_user = st.text_input("Usuario")
            log_pwd = st.text_input("Contraseña", type="password")
            if st.button("Ingresar al Sistema"):
                nombre_usuario = verificar_usuario(log_user, log_pwd)
                if nombre_usuario:
                    st.session_state.logged_in = True
                    st.session_state.username = nombre_usuario
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas.")
                    
        elif opcion_acceso == "Crear Nueva Cuenta":
            reg_nombre = st.text_input("Nombre Completo")
            reg_user = st.text_input("Nombre de Usuario")
            reg_pwd = st.text_input("Contraseña", type="password")
            if st.button("Registrar Cuenta"):
                if reg_nombre and reg_user and reg_pwd:
                    if registrar_usuario(reg_user, reg_pwd, reg_nombre):
                        st.success("🎉 Cuenta creada. Cambia a 'Iniciar Sesión'.")
                    else:
                        st.error("❌ El usuario ya existe.")
                else:
                    st.warning("Completa todos los campos.")
                    
        elif opcion_acceso == "Actualizar Contraseña":
            upd_user = st.text_input("Usuario")
            upd_old_pwd = st.text_input("Contraseña Actual", type="password")
            upd_new_pwd = st.text_input("Nueva Contraseña", type="password")
            if st.button("Cambiar Contraseña"):
                if upd_user and upd_old_pwd and upd_new_pwd:
                    if actualizar_contrasena(upd_user, upd_old_pwd, upd_new_pwd):
                        st.success("✅ Contraseña actualizada con éxito.")
                    else:
                        st.error("❌ Error en los datos proporcionados.")
                        
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# --- REVISOR Y CORRECTOR DE ORTOGRAFÍA SST ---
def corregir_ortografia_sst(texto):
    if not texto: return ""
    diccionario_SST = {
        r'\brealziado\b': 'realizado', r'\brealziados\b': 'realizados',
        r'\baudiometria\b': 'audiometría', r'\bvisiometria\b': 'visiometría',
        r'\bespirometria\b': 'espirometría', r'\boptometria\b': 'optometría',
        r'\bfisica\b': 'física', r'\bmedico\b': 'médico', r'\bperiodico\b': 'periódico',
        r'\bproteccion\b': 'protección', r'\balimentacion\b': 'alimentación',
        r'\brecomendacion\b': 'recomendación', r'\bperfil\s+lipidico\b': 'perfil lipídico',
        r'\benfasis\b': 'énfasis', r'\bosteomuscular\b': 'osteomuscular',
        r'\bregion\b': 'región', r'\bhabitos\b': 'hábitos'
    }
    for patron, reemplazo in diccionario_SST.items():
        texto = re.sub(patron, reemplazo, texto, flags=re.IGNORECASE)
    return texto

def a_caso_oracion(texto):
    if not texto: return ""
    texto_min = corregir_ortografia_sst(texto).lower().strip()
    return re.sub(r'(^|[.!?]\s+|\n+)([a-zñáéíóúü])', lambda m: m.group(1) + m.group(2).upper(), texto_min)

def es_vacio_o_negativo(texto):
    if not texto: return True
    return texto.strip().lower().strip(" .-_/ '\"") in ["no", "ninguna", "ninguno", "no registra", "sin remisiones", "normal", "n/a", "sin remisión"]

def limpiar_campo(texto):
    if not texto: return ""
    partes = re.split(r'\b(Teléfono|Telefono|Tel|C\.C|CC|Documento|Cedula|Cargo|Fecha)\b', texto, flags=re.IGNORECASE)
    return re.sub(r'[:\-,_]+', '', partes[0]).strip().title()

# --- EXTRACTOR AVANZADO CON ACUMULADOR DE MEMORIA (MULTILÍNEA) ---
def analizar_pdf_inteligente(texto):
    datos = {
        "nombre": "", "cargo": "", "tipo_examen": "PERIODICO",
        "examenes_lista": [], "recomendaciones_lista": [], "vigilancia_lista": [],
        "observaciones": "", "remisiones": "No", "consecutivo": ""
    }
    if not texto: return datos

    m_nom = re.search(r'(?:Nombre|Paciente|Colaborador|Trabajador):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_nom: datos["nombre"] = limpiar_campo(m_nom.group(1))

    m_car = re.search(r'(?:Cargo|Ocupación|Ocupacion|Puesto):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_car: datos["cargo"] = limpiar_campo(m_car.group(1))

    for palabra in ["INGRESO", "PERIÓDICO", "PERIODICO", "EGRESO", "RETIRO", "CAMBIO DE CARGO", "POST-INCAPACIDAD", "POST INCAPACIDAD"]:
        if palabra in texto.upper():
            datos["tipo_examen"] = "PERIODICO" if "PERIOD" in palabra else palabra
            break

    EXAMS_MAP = {
        "AUDIOMETRIA DE TONOS": "Audiometría", "AUDIOMETRIA": "Audiometría",
        "ESPIROMETRIA": "Espirometría", "OPTOMETRIA": "Optometría",
        "EXAMEN MEDICO OCUPACIONAL": "Examen Clínico Ocupacional",
        "PERFIL LIPIDICO": "Perfil Lipídico", "GLICEMIA": "Glicemia",
        "ENFASIS OSTEOMUSCULAR": "Énfasis Osteomuscular", "ELECTROCARDIOGRAMA DE RITMO": "Electrocardiograma", 
        "ELECTROCARDIOGRAMA": "Electrocardiograma", "FROTIS": "Frotis",
        "CUADRO HEMATICO": "Cuadro Hemático", "COLESTEROL": "Colesterol",
        "TRIGLICERIDOS": "Triglicéridos", "PARCIAL DE ORINA": "Parcial de Orina",
        "VSH": "VSH", "PCR": "PCR"
    }

    examenes_detectados = []
    recoms_raw_dict = {}
    current_exam = None
    exam_text_accumulator = ""
    in_exams_section = True

    # Lectura Multilínea Inteligente
    lineas = texto.split('\n')
    for linea in lineas:
        linea_upper = linea.upper().strip()
        
        if any(stop in linea_upper for stop in ["OBSERVACIONES:", "OBSERVACION:", "REMISIONES:", "SISTEMA DE VIGILANCIA"]):
            in_exams_section = False
            if current_exam:
                recoms_raw_dict[current_exam] = recoms_raw_dict.get(current_exam, "") + " " + exam_text_accumulator
                current_exam = None
                
        if in_exams_section:
            matched_key = None
            for key in sorted(EXAMS_MAP.keys(), key=len, reverse=True):
                if key in linea_upper and linea_upper.find(key) < 15:
                    matched_key = key
                    break
            
            if matched_key:
                if current_exam:
                    recoms_raw_dict[current_exam] = recoms_raw_dict.get(current_exam, "") + " " + exam_text_accumulator
                
                current_exam = EXAMS_MAP[matched_key]
                if current_exam not in examenes_detectados:
                    examenes_detectados.append(current_exam)
                
                idx = linea_upper.find(matched_key) + len(matched_key)
                exam_text_accumulator = linea[idx:].strip(" :-,_/")
            else:
                if current_exam and linea.strip():
                    exam_text_accumulator += " " + linea.strip()

    if current_exam:
        recoms_raw_dict[current_exam] = recoms_raw_dict.get(current_exam, "") + " " + exam_text_accumulator

    recoms_por_examen = []
    pve_detectados = set()

    for exam in examenes_detectados:
        rec_part = recoms_raw_dict.get(exam, "").strip()
        status_exclusions = ["REALIZADO", "REALZIADO", "SIN ALTERACIONES", "NORMAL", "SANO", "NEGATIVO", "NO REGISTRA", "N/A", ""]
        
        if rec_part.upper().strip(" .") not in status_exclusions and len(rec_part) > 3:
            parts = re.split(r'//|;|\b\d+\.|\b\d+\-', rec_part)
            
            valid_parts = []
            for p in parts:
                p_clean = p.strip(" .-_/()[]")
                if p_clean and len(p_clean) > 3 and p_clean.upper() not in status_exclusions:
                    valid_parts.append(a_caso_oracion(p_clean))
                    
                    p_upper = p_clean.upper()
                    
                    # --- CORRECCIÓN CLAVE: BÚSQUEDA MEDIANTE LÍMITES DE PALABRA (\b) PARA EVITAR FALSOS POSITIVOS ---
                    if any(re.search(patron, p_upper) for patron in [r'\bAUDITIV', r'\bRUIDO', r'\bOIDO', r'\bOÍDO', r'\bAUDIO']):
                        pve_detectados.add("Conservación Auditiva")
                    elif any(re.search(patron, p_upper) for patron in [r'\bPOSTURAL', r'\bLUMBAR', r'\bOSTEOMUSCULAR', r'\bERGONOMIC', r'\bESPALDA', r'\bCARGA']):
                        pve_detectados.add("Prevención Osteomuscular (DME)")
                    # \bVISIÓN / \bVISION evita que "revisión" o "división" activen erróneamente la conservación visual
                    elif any(re.search(patron, p_upper) for patron in [r'\bVISUAL', r'\bGAFAS', r'\bVISION', r'\bVISIÓN', r'\bLENTE', r'\bOPTOMETR', r'\bRX\b']):
                        pve_detectados.add("Conservación Visual")
                    elif any(re.search(patron, p_upper) for patron in [r'\bRESPIRATORI', r'\bESPIROMETR', r'\bPOLVO', r'\bHUMO']):
                        pve_detectados.add("Conservación Respiratoria")
            
            if valid_parts:
                recoms_por_examen.append(f"{exam}: {' - '.join(valid_parts)}")

    datos["examenes_lista"] = examenes_detectados
    datos["recomendaciones_lista"] = recoms_por_examen
    datos["vigilancia_lista"] = list(pve_detectados)

    def extraer_seccion(texto_completo, palabras_inicio, palabras_fin):
        seccion = []
        dentro = False
        for l in texto_completo.split('\n'):
            l_upper = l.upper().strip()
            if not dentro:
                if any(h in l_upper for h in palabras_inicio):
                    dentro = True
                    for h in palabras_inicio:
                        if h in l_upper:
                            resto = l[l_upper.find(h) + len(h):].strip(" :-,_")
                            if resto: seccion.append(resto)
                            break
            else:
                if any(h in l_upper for h in palabras_fin): break
                seccion.append(l.strip())
        return "\n".join(seccion).strip()

    datos["observaciones"] = a_caso_oracion(extraer_seccion(texto, ["OBSERVACIONES:"], ["RECOMENDACIONES", "REMISIONES"]))
    
    rem_raw = extraer_seccion(texto, ["INFORMACION DE REMISIONES", "INFORMACIÓN DE REMISIONES"], ["CONSENTIMIENTO", "AUTORIZO"])
    datos["remisiones"] = "No" if es_vacio_o_negativo(rem_raw) else a_caso_oracion(rem_raw)

    return datos

# --- FORMATEADOR DINÁMICO DE NEGRITAS EN EL CUERPO ---
def aplicar_negrita_dinamica_cuerpo(paragraph, tipo_examen):
    texto_parrafo = paragraph.text
    if "Según los lineamientos del programa de medicina preventiva" not in texto_parrafo:
        return
        
    paragraph.text = "" 
    p1 = "Según los lineamientos del programa de medicina preventiva y del trabajo de JER S.A; se hace entrega de las recomendaciones establecidas por el Proveedor de servicios de Exámenes Médico Ocupacionales ("
    paragraph.add_run(p1)
    
    opciones = [
        ("Ingreso", "INGRESO" in tipo_examen.upper()),
        ("Periódico", "PERIODIC" in tipo_examen.upper() or "PERIÓDIC" in tipo_examen.upper()),
        ("egreso", "EGRESO" in tipo_examen.upper() or "RETIRO" in tipo_examen.upper()),
        ("cambio de cargo", "CAMBIO" in tipo_examen.upper()),
        ("post incapacidad", "POST" in tipo_examen.upper() or "INCAPACIDAD" in tipo_examen.upper())
    ]
    
    for i, (texto_opcion, condicion) in enumerate(opciones):
        run = paragraph.add_run(texto_opcion)
        run.bold = condicion 
        if i < len(opciones) - 1:
            if i == len(opciones) - 2: paragraph.add_run(" y ")
            else: paragraph.add_run(", ")
                
    paragraph.add_run(")")

# --- MANEJO DE CONSECUTIVO Y VIÑETAS CLONADAS ---
def obtener_siguiente_consecutivo_local():
    val = obtener_config("ultimo_consecutivo_local")
    return int(val) + 1 if val else 1

def incrementar_consecutivo_local():
    next_num = obtener_siguiente_consecutivo_local()
    guardar_config("ultimo_consecutivo_local", str(next_num))
    return f"SST-2026-{next_num}"

def replace_in_paragraph(paragraph, key, value):
    if key not in paragraph.text: return
    replaced_in_runs = False
    for run in paragraph.runs:
        if key in run.text:
            font_name, font_size, bold, italic, color = run.font.name, run.font.size, run.bold, run.italic, (run.font.color.rgb if run.font.color else None)
            run.text = run.text.replace(key, value)
            if font_name: run.font.name = font_name
            if font_size: run.font.size = font_size
            run.bold, run.italic = bold, italic
            if color: run.font.color.rgb = color
            replaced_in_runs = True
    if not replaced_in_runs: paragraph.text = paragraph.text.replace(key, value)

def replace_placeholder_with_bullets(cell, placeholder, items_list):
    for p in cell.paragraphs:
        if placeholder in p.text:
            font_name, font_size = "Arial", Pt(11)
            if p.runs:
                font_name = p.runs[0].font.name or "Arial"
                font_size = p.runs[0].font.size or Pt(11)
            p.text = ""
            if not items_list:
                p.add_run("Ninguno.").font.name = font_name
                return
            
            run = p.add_run("• " + items_list[0])
            run.font.name, run.font.size = font_name, font_size
            current_p = p
            for item in items_list[1:]:
                new_p = OxmlElement('w:p')
                current_p._p.addnext(new_p)
                new_para = Paragraph(new_p, cell)
                new_para.paragraph_format.line_spacing = p.paragraph_format.line_spacing
                run_new = new_para.add_run("• " + item)
                run_new.font.name, run_new.font.size = font_name, font_size
                current_p = new_para
            return

def procesar_remisiones_en_celda(cell, remisiones_text):
    for p in cell.paragraphs:
        if "{{Remisiones}}" in p.text:
            p.text = ""
            run_label = p.add_run("remisiones: ")
            run_label.bold = True
            if es_vacio_o_negativo(remisiones_text):
                p.add_run("no")
            else:
                p.add_run(remisiones_text)

# --- LIMPIEZA DE CARACTERES ESPECIALES PARA PDF (LATIN-1) ---
def clean_pdf_str(text):
    if not text: return ""
    replacements = {
        "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
        "\u2013": "-", "\u2014": "-", "\u2022": "*", "\xfa": "ú"
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text.encode('latin-1', 'replace').decode('latin-1')

# --- GENERADOR DE PDF NATIVO PROFESIONAL (SST) ---
def generar_pdf_nativo(datos, consecutivo_num, lugar, fecha, firma_file):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Paleta de colores Corporativa
    primary_color = (31, 78, 121)  # Azul #1f4e79
    dark_neutral = (40, 40, 40)
    
    def s(txt):
        return clean_pdf_str(txt)
    
    # Encabezado (Banda de fondo clara)
    pdf.set_fill_color(240, 244, 248)
    pdf.rect(10, 10, 190, 20, "F")
    
    # Texto de Encabezado
    pdf.set_text_color(*primary_color)
    pdf.set_font("Arial", "B", 11)
    pdf.set_xy(15, 12)
    pdf.cell(110, 8, s("JER S.A. - PORTAL DE MEDICINA PREVENTIVA Y DEL TRABAJO"), 0, 0, "L")
    
    # Caja de Consecutivo
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(*primary_color)
    pdf.set_xy(140, 13)
    pdf.cell(55, 14, s(consecutivo_num), 0, 0, "C", True)
    
    # Reset del cursor de texto
    pdf.set_xy(10, 36)
    pdf.set_text_color(*dark_neutral)
    
    # Fecha y lugar
    pdf.set_font("Arial", "", 10)
    fecha_texto = f"{lugar}, {fecha.strftime('%d de %B de %Y')}"
    pdf.cell(0, 10, s(fecha_texto), 0, 1, "L")
    pdf.ln(1)
    
    # Datos del Colaborador
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 5, s("Sr(a)."), 0, 1, "L")
    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(*primary_color)
    pdf.cell(0, 6, s(datos['nombre'].upper()), 0, 1, "L")
    pdf.set_font("Arial", "I", 10)
    pdf.set_text_color(*dark_neutral)
    pdf.cell(0, 5, s(f"Cargo: {datos['cargo']}"), 0, 1, "L")
    pdf.ln(4)
    
    # Línea Divisoria
    pdf.set_draw_color(*primary_color)
    pdf.set_line_width(0.5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)
    
    # Asunto
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(20, 5, s("ASUNTO:"), 0, 0, "L")
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(*dark_neutral)
    pdf.cell(0, 5, s(f"RECOMENDACIONES EXAMEN MÉDICO DE {datos['tipo_examen'].upper()}"), 0, 1, "L")
    pdf.ln(4)
    
    # Cuerpo de la carta
    pdf.set_font("Arial", "", 10)
    body_text = (
        "Cordial saludo:\n\n"
        "Según los lineamientos del programa de medicina preventiva y del trabajo de JER S.A; "
        "se hace entrega de las recomendaciones establecidas por el Proveedor de servicios de Exámenes "
        f"Médico Ocupacionales de tipo {datos['tipo_examen'].title()}:"
    )
    pdf.multi_cell(0, 5, s(body_text))
    pdf.ln(4)
    
    # Exámenes Realizados
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(0, 6, s("EXÁMENES REALIZADOS:"), 0, 1, "L")
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(*dark_neutral)
    for ex in datos['examenes_lista']:
        pdf.cell(5)
        pdf.cell(0, 5, s(f"- {ex}"), 0, 1, "L")
    pdf.ln(4)
    
    # Recomendaciones
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(0, 6, s("RECOMENDACIONES MÉDICAS:"), 0, 1, "L")
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(*dark_neutral)
    if datos['recomendaciones_lista']:
        for rec in datos['recomendaciones_lista']:
            pdf.set_x(15)
            pdf.multi_cell(0, 5, s(f"• {rec}"))
    else:
        pdf.cell(5)
        pdf.cell(0, 5, s("Ninguna."), 0, 1, "L")
    pdf.ln(4)
    
    # Observaciones
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(32, 5, s("OBSERVACIONES:"), 0, 0, "L")
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(*dark_neutral)
    obs_text = datos['observaciones'] if datos['observaciones'] else "Ninguna."
    pdf.multi_cell(0, 5, s(obs_text))
    pdf.ln(2)
    
    # Remisiones
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(26, 5, s("REMISIONES:"), 0, 0, "L")
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(*dark_neutral)
    rem_text = datos['remisiones'] if datos['remisiones'] else "No presenta remisiones."
    pdf.multi_cell(0, 5, s(rem_text))
    pdf.ln(8)
    
    # Prevenir que la firma quede huérfana en el final de página
    if pdf.get_y() > 220:
        pdf.add_page()
        
    # Estampa de Firma Autorizada
    if firma_file:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_img:
                temp_img.write(firma_file.getvalue())
                temp_path = temp_img.name
            pdf.image(temp_path, x=15, y=pdf.get_y(), w=42)
            pdf.ln(18)
            os.unlink(temp_path)
        except Exception:
            pdf.ln(10)
    else:
        pdf.ln(10)
        
    # Firmante
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(*primary_color)
    pdf.cell(0, 5, s("VÍCTOR ALONSO MORENO CASAS"), 0, 1, "L")
    pdf.set_font("Arial", "", 9)
    pdf.set_text_color(*dark_neutral)
    pdf.cell(0, 4, s("Coordinador SST"), 0, 1, "L")
    pdf.cell(0, 4, s("JER S.A."), 0, 1, "L")
    
    # Compatibilidad de bytes entre fpdf y fpdf2
    pdf_out = pdf.output()
    if isinstance(pdf_out, str):
        return pdf_out.encode('latin1')
    return bytes(pdf_out)

# --- GENERADOR DE HTML COMPATIBLE PARA VISTA DE IMPRESIÓN ---
def generar_html_vista(datos, consecutivo_num, lugar, fecha):
    return f"""
    <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; line-height: 1.5; background: white; border: 1px solid #ccc; max-width: 800px; margin: auto;">
        <style>
            @media print {{
                body {{ background: white; color: black; }}
                @page {{ size: letter; margin: 20mm; }}
            }}
        </style>
        <div style="text-align: right; font-weight: bold; color: #1f4e79;">Consecutivo: {consecutivo_num}</div>
        <div style="text-align: center; font-weight: bold; font-size: 16px; margin: 20px 0; color: #1f4e79; background: #f0f4f8; padding: 8px;">
            ASUNTO: RECOMENDACIONES EXAMEN {datos['tipo_examen'].upper()}
        </div>
        <div>{lugar}, {fecha.strftime('%d de %B de %Y')}</div><br>
        <div>Sr(a).<br><strong>{datos['nombre']}</strong><br>{datos['cargo']}</div><br>
        <p>Cordial saludo,</p>
        <p>Según los lineamientos del programa de medicina preventiva y del trabajo de JER S.A; se hace entrega de las recomendaciones establecidas por el Proveedor de servicios de Exámenes Médico Ocupacionales (Ingreso, Periódico, egreso, cambio de cargo y post incapacidad)</p>
        <p><strong>EXÁMENES REALIZADOS:</strong></p>
        <ul>{"".join([f"<li>{ex}</li>" for ex in datos['examenes_lista']])}</ul>
        <p><strong>Recomendaciones:</strong></p>
        <ul>{"".join([f"<li>{rec}</li>" for rec in datos['recomendaciones_lista']])}</ul>
        <p><strong>observaciones:</strong> {datos['observaciones']}</p>
        <p><strong>remisiones:</strong> {datos['remisiones']}</p><br>
        <p>Atentamente,</p><br>
        <p><strong>VÍCTOR ALONSO MORENO CASAS</strong><br>Coordinador SST</p>
    </div>
    """

# --- CONSTRUCTOR DE DOCUMENTO ÚNICO INTELIGENTE ---
def generar_word_unico(datos_trabajador, lugar, fecha, template_uploaded, firma_file):
    if template_uploaded:
        doc_word = Document(template_uploaded)
    elif os.path.exists("FORMATO RECOMENDACIONES MEDICAS BOT.docx"):
        doc_word = Document("FORMATO RECOMENDACIONES MEDICAS BOT.docx")
    else:
        doc_word = Document()
    
    consecutivo_final = datos_trabajador.get("consecutivo", "")
    if not consecutivo_final:
        g_url = obtener_config("google_sheets_url")
        if g_url:
            try:
                r = requests.get(g_url, params={
                    "name": datos_trabajador["nombre"], 
                    "cargo": datos_trabajador["cargo"], 
                    "examen": datos_trabajador["tipo_examen"], 
                    "fecha": fecha.strftime("%Y-%m-%d")
                }, timeout=10)
                consecutivo_final = r.json().get("consecutive") if r.json().get("status") == "success" else incrementar_consecutivo_local()
            except: consecutivo_final = incrementar_consecutivo_local()
        else: consecutivo_final = incrementar_consecutivo_local()
        datos_trabajador["consecutivo"] = consecutivo_final

    replacements = {
        "{{NUMERO DE CONSECUTIVO}}": consecutivo_final, 
        "{{TIPO DE EXAMEN}}": datos_trabajador["tipo_examen"].upper(),
        "{{LUGAR}}": lugar, 
        "{{FECHA HOY}}": fecha.strftime("%d de %B de %Y"),
        "{{NOMBRE DE LA PERSONA}}": datos_trabajador["nombre"], 
        "{{CARGO DE LA PERSONA}}": datos_trabajador["cargo"],
        "{{Observaciones}}": datos_trabajador["observaciones"]
    }

    for table in doc_word.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in replacements.items():
                        if k in p.text: p.text = p.text.replace(k, v)
                    aplicar_negrita_dinamica_cuerpo(p, datos_trabajador["tipo_examen"])
                replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS}}", datos_trabajador["examenes_lista"])
                replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS", datos_trabajador["examenes_lista"])
                replace_placeholder_with_bullets(cell, "{{Recomendaciones médicas}}", datos_trabajador["recomendaciones_lista"])
                procesar_remisiones_en_celda(cell, datos_trabajador["remisiones"])
                
                idx_victor = -1
                for idx, p in enumerate(cell.paragraphs):
                    if "VÍCTOR ALONSO MORENO CASAS" in p.text:
                        idx_victor = idx
                        break
                if idx_victor != -1 and firma_file:
                    p_firma = cell.paragraphs[idx_victor - 2]
                    p_firma.text = ""
                    p_firma.add_run().add_picture(firma_file, width=Inches(2.2))

    b_io = io.BytesIO()
    doc_word.save(b_io)
    return b_io.getvalue(), consecutivo_final

# --- VISTA STREAMLIT ---
st.markdown("<div class='header-banner'><h1>🩺 Portal de Control SST - JER S.A.</h1><p>Generación de Comunicaciones con Negrita Dinámica, Google Sheets y Firma Digital</p></div>", unsafe_allow_html=True)

# BARRA LATERAL (CONFIGURACIÓN)
st.sidebar.markdown(f"<h3 style='color:#60a5fa;'>👤 Perfil Activo</h3>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='metric-card'><strong>Usuario:</strong> {st.session_state.username}</div>", unsafe_allow_html=True)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.documentos = {}
    st.session_state.textos_raw = {}
    st.session_state.export_bytes = None
    st.session_state.zip_bytes = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔗 Configuración de Nube Sheets")
g_url_guardada = obtener_config("google_sheets_url")
g_url_input = st.sidebar.text_input("URL de Google Apps Script:", value=g_url_guardada, type="password")
if st.sidebar.button("Guardar Conexión"):
    guardar_config("google_sheets_url", g_url_input)
    st.sidebar.success("¡Enlace guardado!")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📋 Documentación Base")
template_uploaded = st.sidebar.file_uploader("Formato de Word Institucional (.docx)", type=["docx"])
firma_file = st.sidebar.file_uploader("Estampa de Firma Autorizada (.png / .jpg)", type=["png", "jpg"])

# DIVISION EN COLUMNAS
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.markdown("<h3 style='color:#60a5fa;'>📂 1. Carga de Documentos PDF</h3>", unsafe_allow_html=True)
    pdfs_subidos = st.file_uploader("Carga los archivos PDF:", type="pdf", accept_multiple_files=True)
    
    if pdfs_subidos:
        for pdf in pdfs_subidos:
            if pdf.name not in st.session_state.documentos:
                with pdfplumber.open(pdf) as p_file:
                    texto_raw = "".join([page.extract_text() + "\n" for page in p_file.pages])
                st.session_state.documentos[pdf.name] = analizar_pdf_inteligente(texto_raw)
        
        st.markdown(f"""
            <div style='display:flex; gap:10px; margin-top:15px;'>
                <div class='metric-card' style='flex:1;'><strong>PDFs Leídos</strong><br><span style='font-size:20px; font-weight:700; color:#60a5fa;'>{len(st.session_state.documentos)}</span></div>
                <div class='metric-card' style='flex:1;'><strong>Nube Sheets</strong><br><span style='font-size:14px; font-weight:600; color:#4ade80;'>{'Conectado' if g_url_guardada else 'Modo Local'}</span></div>
            </div>
        """, unsafe_allow_html=True)
        
        archivo_seleccionado = st.selectbox("🎯 Selecciona Colaborador:", list(st.session_state.documentos.keys()))
    else:
        archivo_seleccionado = None

with col_der:
    st.markdown("<h3 style='color:#60a5fa;'>📋 2. Editor del Trabajador Seleccionado</h3>", unsafe_allow_html=True)
    if archivo_seleccionado:
        doc_actual = st.session_state.documentos[archivo_seleccionado]
        
        col_f1, col_f2 = st.columns(2)
        with col_f1: lugar = st.text_input("Lugar:", value="Tunja")
        with col_f2: fecha = st.date_input("Fecha:", value=datetime.date.today())
        
        tipo_examen = st.text_input("Tipo de Examen:", value=doc_actual["tipo_examen"].upper())
        
        col_p1, col_p2 = st.columns(2)
        with col_p1: nombre_persona = st.text_input("Trabajador:", value=doc_actual["nombre"])
        with col_p2: cargo_persona = st.text_input("Cargo:", value=doc_actual["cargo"])
        
        examenes_realizados = st.text_area("Exámenes Realizados:", value="\n".join(doc_actual["examenes_lista"]))
        recom_medicas = st.text_area("Recomendaciones por Examen:", value="\n".join(doc_actual["recomendaciones_lista"]), height=130)
        observaciones = st.text_area("Observaciones:", value=doc_actual["observaciones"])
        remisiones = st.text_input("Remisiones (Escribe 'No' para marcarlo negativo):", value=doc_actual["remisiones"])

        doc_actual.update({
            "nombre": nombre_persona, "cargo": cargo_persona, "tipo_examen": tipo_examen,
            "examenes_lista": [l.strip() for l in examenes_realizados.split('\n') if l.strip()],
            "recomendaciones_lista": [l.strip() for l in recom_medicas.split('\n') if l.strip()],
            "observaciones": observaciones, "remisiones": remisiones
        })

        st.markdown("---")
        formato_salida = st.radio("⚡ Elige formato de generación:", ["Documento PDF Oficial (.pdf)", "Microsoft Word (.docx)", "Impresión de Respaldo Web (HTML)"], horizontal=True)
        
        col_act1, col_gen2 = st.columns(2)
        
        with col_act1:
            if st.button("✨ Procesar y Descargar este Colaborador"):
                with st.spinner("Procesando documento..."):
                    if "Word" in formato_salida:
                        bytes_word, consec_num = generar_word_unico(doc_actual, lugar, fecha, template_uploaded, firma_file)
                        if bytes_word:
                            st.success(f"🟢 Guardado en Sheets (Consecutivo: {consec_num})")
                            st.download_button("📥 Descargar Word (.docx)", data=bytes_word, file_name=f"Informe_{nombre_persona.replace(' ','_')}.docx")
                    elif "PDF" in formato_salida:
                        if fpdf_disponible:
                            _, consec_num = generar_word_unico(doc_actual, lugar, fecha, template_uploaded, None)
                            bytes_pdf = generar_pdf_nativo(doc_actual, consec_num, lugar, fecha, firma_file)
                            st.success(f"🟢 Guardado en Sheets (Consecutivo: {consec_num})")
                            st.download_button("📥 Descargar PDF Oficial (.pdf)", data=bytes_pdf, file_name=f"Informe_{nombre_persona.replace(' ','_')}.pdf", mime="application/pdf")
                        else:
                            st.error("Librería fpdf2 no instalada en el entorno. Por favor ejecute: pip install fpdf2")
                    else:
                        _, consec_num = generar_word_unico(doc_actual, lugar, fecha, template_uploaded, None)
                        html_out = generar_html_vista(doc_actual, consec_num, lugar, fecha)
                        st.success(f"🟢 Guardado en Sheets (Consecutivo: {consec_num})")
                        st.download_button("📥 Descargar Documento Imprimible (.html)", data=html_out.encode('utf-8'), file_name=f"Informe_{nombre_persona.replace(' ','_')}.html")
                        
        with col_gen2:
            if len(st.session_state.documentos) > 1:
                if st.button("📦 Generar TODOS los Colaboradores (ZIP)"):
                    with st.spinner("Conectando en lote con la nube..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for filename, datos_trab in st.session_state.documentos.items():
                                if "Word" in formato_salida:
                                    bytes_word, consec_num = generar_word_unico(datos_trab, lugar, fecha, template_uploaded, firma_file)
                                    if bytes_word:
                                        zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.docx", bytes_word)
                                elif "PDF" in formato_salida:
                                    _, consec_num = generar_word_unico(datos_trab, lugar, fecha, template_uploaded, None)
                                    bytes_pdf = generar_pdf_nativo(datos_trab, consec_num, lugar, fecha, firma_file)
                                    zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.pdf", bytes_pdf)
                                else:
                                    _, consec_num = generar_word_unico(datos_trab, lugar, fecha, template_uploaded, None)
                                    html_out = generar_html_vista(datos_trab, consec_num, lugar, fecha)
                                    zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.html", html_out.encode('utf-8'))
                                    
                        zip_buffer.seek(0)
                        st.session_state.zip_bytes = zip_buffer.getvalue()
                        st.success("🎉 ZIP de lote masivo compilado con éxito.")
                        
                if st.session_state.zip_bytes:
                    st.download_button("📥 Descargar ZIP Masivo", data=st.session_state.zip_bytes, file_name=f"Lote_SST_JER_SA_{fecha.strftime('%Y%m%d')}.zip", mime="application/zip")
    else:
        st.markdown("<div style='text-align:center; padding: 40px; color:#64748b;'><h3>👋 Tablero Listo</h3><p>Por favor, arrastra tus archivos PDF en la sección izquierda para activar el procesamiento automático.</p></div>", unsafe_allow_html=True)
