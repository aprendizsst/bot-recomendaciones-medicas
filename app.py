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

# --- CONFIGURACIÓN DE PÁGINA AVANZADA ---
st.set_page_config(
    page_title="Portal SST - JER S.A.", 
    page_icon="🩺", 
    layout="wide"
)

# --- INYECCIÓN DE CSS ESTRICTO PARA CORREGIR CONTRASTE Y COLORES ---
st.markdown("""
    <style>
    /* Fondo global de la aplicación */
    [data-testid="stAppViewContainer"] {
        background-color: #f8fafc !important;
    }
    
    /* CONTROL DE ALTO CONTRASTE PARA LOGIN */
    .login-box {
        max-width: 450px;
        margin: 60px auto;
        padding: 35px;
        background-color: #ffffff !important;
        border-radius: 14px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }
    
    /* Forzar visibilidad de títulos y subtítulos */
    .login-box h2 {
        color: #1f4e79 !important;
        text-align: center;
        margin-bottom: 5px;
        font-weight: 700;
    }
    .login-box p {
        color: #475569 !important;
        text-align: center;
        font-size: 14px;
    }
    
    /* FORZAR VISIBILIDAD EN TEXTOS DE RADIOS (Botones de selección) */
    div[data-testid="stRadio"] label p {
        color: #1e293b !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }
    
    /* FORZAR VISIBILIDAD EN LAS ETIQUETAS DE LOS INPUTS (Usuario, Contraseña, etc.) */
    div[data-testid="stWidgetLabel"] p {
        color: #1f4e79 !important;
        font-weight: 600 !important;
        font-size: 15px !important;
    }
    
    /* CORRECCIÓN ABSOLUTA DE LAS CAJAS DE TEXTO (INPUTS BLANCOS) */
    div[data-baseweb="input"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="input"] input {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }
    div[data-testid="stTextArea"] textarea {
        color: #0f172a !important;
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }
    
    /* DISEÑO DE BOTONES PREMIUM */
    .stButton>button {
        background: linear-gradient(135deg, #1f4e79 0%, #2a6f97 100%) !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        border: none !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 10px rgba(31, 78, 121, 0.15) !important;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 15px rgba(31, 78, 121, 0.25) !important;
        filter: brightness(105%);
    }
    
    /* Banner institucional interno */
    .header-banner {
        background: linear-gradient(135deg, #1f4e79 0%, #2a6f97 100%);
        padding: 22px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(31, 78, 121, 0.15);
    }
    
    /* Tarjetas informativas internas */
    .metric-card {
        background: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
        border-left: 5px solid #1f4e79;
        margin-bottom: 12px;
        color: #334155;
    }
    </style>
""", unsafe_allow_html=True)

# --- SISTEMA DE SEGURIDAD Y CONFIGURACIÓN (BASE DE DATOS) ---
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

# --- REVISOR Y CORRECTOR DE ORTOGRAFÍA SST AUTÓNOMO ---
def corregir_ortografia_sst(texto):
    if not texto: 
        return ""
    diccionario_SST = {
        r'\brealziado\b': 'realizado',
        r'\brealziados\b': 'realizados',
        r'\baudiometria\b': 'audiometría',
        r'\bvisiometria\b': 'visiometría',
        r'\bespirometria\b': 'espiometría',
        r'\boptometria\b': 'optometría',
        r'\bfisica\b': 'física',
        r'\bmedico\b': 'médico',
        r'\bperiodico\b': 'periódico',
        r'\bproteccion\b': 'protección',
        r'\balimentacion\b': 'alimentación',
        r'\brecomendacion\b': 'recomendación',
        r'\brecomendaciones\s+medicas\b': 'recomendaciones médicas',
        r'\boptometra\b': 'optómetra',
        r'\boftalmologia\b': 'oftalmología',
        r'\baudiologo\b': 'audiólogo',
        r'\bposicion\b': 'posición',
        r'\bactiba\b': 'activa',
        r'\bactibas\b': 'activas',
        r'\bevaluacion\b': 'evaluación',
        r'\bcoordinacion\b': 'coordinación',
        r'\butilizacion\b': 'utilización',
        r'\bexposicion\b': 'exposición',
        r'\bhidratacion\b': 'hidratación',
        r'\bsegun\b': 'según',
        r'\bperfil\s+lipidico\b': 'perfil lipídico',
        r'\benfasis\b': 'énfasis',
        r'\bosteomuscular\b': 'osteomuscular'
    }
    texto_corregido = texto
    for patron, reemplazo in diccionario_SST.items():
        texto_corregido = re.sub(patron, reemplazo, texto_corregido, flags=re.IGNORECASE)
    return texto_corregido

# --- FUNCIÓN DE FORMATO: CASO ORACIÓN (SENTENCE CASE) ---
def a_caso_oracion(texto):
    if not texto: 
        return ""
    texto_sano = corregir_ortografia_sst(texto)
    texto_min = texto_sano.lower().strip()
    
    def capitalizar_match(match):
        return match.group(1) + match.group(2).upper()
    
    texto_formateado = re.sub(r'(^|[.!?]\s+|\n+)([a-zñáéíóúü])', capitalizar_match, texto_min)
    return texto_formateado

# --- VALIDACIÓN DE VALOR VACÍO O NEGATIVO ---
def es_vacio_o_negativo(texto):
    if not texto:
        return True
    t_clean = texto.strip().lower().strip(" .-_/ '\"")
    exclusiones = ["no", "ninguna", "ninguno", "no registra", "sin remisiones", "normal", "n/a", "sin remisión", "sin remision", "no se registran", "no aplica", "ninguno."]
    return t_clean in exclusiones

# --- FUNCIONES DE EXTRACCIÓN ---
def limpiar_campo(texto):
    if not texto: return ""
    exclusiones = r'\b(Teléfono|Telefono|Tel|C\.C|CC|Documento|Identificac|Cedula|Cédula|Edad|Sexo|Cargo|Fecha|Estado|Empresa|Ciudad)\b'
    partes = re.split(exclusiones, texto, flags=re.IGNORECASE)
    texto_limpio = partes[0]
    texto_limpio = re.sub(r'[:\-,_]+', '', texto_limpio)
    return texto_limpio.strip().title()

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

    m_tipo = re.search(r'(?:Tipo de Examen|Concepto|Evaluación|Evaluacion|Motivo|Clase de Examen):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_tipo: 
        datos["tipo_examen"] = limpiar_campo(m_tipo.group(1)).upper()
    else:
        for palabra in ["INGRESO", "PERIÓDICO", "PERIODICO", "EGRESO", "RETIRO", "CAMBIO DE CARGO", "POST-INCAPACIDAD", "POST INCAPACIDAD"]:
            if palabra in texto.upper():
                datos["tipo_examen"] = palabra
                break

    EXAMS_MAP = {
        "AUDIOMETRIA DE TONOS": "Audiometría",
        "AUDIOMETRIA": "Audiometría",
        "ESPIROMETRIA": "Espiometría",
        "OPTOMETRIA": "Optometría",
        "EXAMEN MEDICO OCUPACIONAL": "Examen Clínico Ocupacional",
        "PERFIL LIPIDICO": "Perfil Lipídico",
        "GLICEMIA": "Glicemia",
        "ENFASIS OSTEOMUSCULAR": "Énfasis Osteomuscular",
        "ELECTROCARDIOGRAMA DE RITMO O DE SUPERFICIE SOD": "Electrocardiograma",
        "ELECTROCARDIOGRAMA": "Electrocardiograma",
        "FROTIS": "Frotis",
        "CUADRO HEMATICO": "Cuadro Hemático",
        "COLESTEROL": "Colesterol",
        "TRIGLICERIDOS": "Triglicéridos",
        "PARCIAL DE ORINA": "Parcial de Orina",
        "VSH": "VSH",
        "PCR": "PCR"
    }

    examenes_detectados = []
    recomendaciones_detectadas = []
    pve_detectados = set()

    lineas = texto.split('\n')
    for linea in lineas:
        linea_upper = linea.upper().strip()
        matched_key = None
        for key in sorted(EXAMS_MAP.keys(), key=len, reverse=True):
            if key in linea_upper:
                matched_key = key
                break
        
        if matched_key:
            nombre_examen = EXAMS_MAP[matched_key]
            if nombre_examen not in examenes_detectados:
                examenes_detectados.append(nombre_examen)
            
            idx = linea_upper.find(matched_key) + len(matched_key)
            rec_part = linea[idx:].strip(" :-,_/")
            
            status_exclusions = ["REALIZADO", "REALZIADO", "SIN ALTERACIONES", "NORMAL", "SANO", "NEGATIVO", "NO REGISTRA", "N/A", ""]
            if rec_part.upper().strip(" .") not in status_exclusions:
                parts = re.split(r'//|,|\b\d+\.|\b\d+\-', rec_part)
                for p in parts:
                    p_clean = p.strip(" .-_/()[]")
                    if p_clean and len(p_clean) > 3:
                        if p_clean.upper() not in status_exclusions and p_clean.upper() != "SST":
                            recomendaciones_detectadas.append(a_caso_oracion(p_clean))
                            
                            p_upper = p_clean.upper()
                            if any(w in p_upper for w in ["AUDITIV", "RUIDO", "OIDO", "AUDIO"]):
                                pve_detectados.add("Vigilancia Epidemiológica de Conservación Auditiva")
                            elif any(w in p_upper for w in ["OSTEOMUSCULAR", "POSTURAL", "ERGONOMIC", "LUMBAR", "ESPALDA", "DESORDEN", "ESQUEL", "COLUMNA"]):
                                pve_detectados.add("Vigilancia Epidemiológica de Prevención Osteomuscular (DME)")
                            elif any(w in p_upper for w in ["VISUAL", "OPTIC", "GAFAS", "OJOS", "RX", "VISION", "LENTES"]):
                                pve_detectados.add("Vigilancia Epidemiológica de Conservación Visual")
                            elif any(w in p_upper for w in ["RESPIRATORI", "ESPIROMETR", "POLVO", "HUMO", "RESPIRACION"]):
                                pve_detectados.add("Vigilancia Epidemiológica de Conservación Respiratoria")

    datos["examenes_lista"] = examenes_detectados
    datos["recomendaciones_lista"] = recomendaciones_detectadas
    datos["vigilancia_lista"] = list(pve_detectados)

    def extraer_seccion(texto_completo, palabras_inicio, palabras_fin):
        lineas_bloque = texto_completo.split('\n')
        seccion = []
        dentro = False
        for l in lineas_bloque:
            l_upper = l.upper().strip()
            if not dentro:
                if any(h in l_upper for h in palabras_inicio):
                    dentro = True
                    for h in palabras_inicio:
                        if h in l_upper:
                            idx = l_upper.find(h) + len(h)
                            resto = l[idx:].strip(" :-,_")
                            if resto: seccion.append(resto)
                            break
            else:
                if any(h in l_upper for h in palabras_fin):
                    break
                seccion.append(l.strip())
        return "\n".join(seccion).strip()

    datos["observaciones"] = a_caso_oracion(extraer_seccion(texto,
        ["OBSERVACIONES:", "OBSERVACION:", "OBSERVACIONES"],
        ["RECOMENDACIONES", "REMISIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE"]
    ))

    # Extracción estricta de remisiones
    rem_raw = extraer_seccion(texto,
        ["INFORMACION DE REMISIONES", "INFORMACIÓN DE REMISIONES"],
        ["RECOMENDACIONES", "OBSERVACIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE", "DIAGNOSTICOS", "DIAGNÓSTICOS", "CONSENTIMIENTO", "AUTORIZO"]
    )
    for stop_w in ["CONSENTIMIENTO", "AUTORIZO", "DOCTOR", "PARACLINICO", "PARACLÍNICO", "ABAJO MENCIONADO"]:
        if stop_w in rem_raw.upper():
            idx = rem_raw.upper().find(stop_w)
            rem_raw = rem_raw[:idx].strip(" :-,_/'\"")

    if es_vacio_o_negativo(rem_raw):
        datos["remisiones"] = "No"
    else:
        datos["remisiones"] = a_caso_oracion(rem_raw)

    return datos

# --- MANEJO DE CONSECUTIVO LOCAL ---
def obtener_siguiente_consecutivo_local():
    val = obtener_config("ultimo_consecutivo_local")
    if not val: return 1
    try: return int(val) + 1
    except: return 1

def incrementar_consecutivo_local():
    next_num = obtener_siguiente_consecutivo_local()
    guardar_config("ultimo_consecutivo_local", str(next_num))
    return f"SST-2026-{next_num}"

# --- REEMPLAZO DE MARCADORES ---
def replace_in_paragraph(paragraph, key, value):
    if key not in paragraph.text:
        return
    replaced_in_runs = False
    for run in paragraph.runs:
        if key in run.text:
            font_name = run.font.name
            font_size = run.font.size
            bold = run.bold
            italic = run.italic
            color = run.font.color.rgb if run.font.color else None
            
            run.text = run.text.replace(key, value)
            
            if font_name: run.font.name = font_name
            if font_size: run.font.size = font_size
            run.bold = bold
            run.italic = italic
            if color: run.font.color.rgb = color
            replaced_in_runs = True
            
    if not replaced_in_runs:
        paragraph.text = paragraph.text.replace(key, value)

# --- REEMPLAZO DINÁMICO CLONANDO LA CALIGRAFÍA ---
def replace_placeholder_with_bullets(cell, placeholder, items_list):
    for p in cell.paragraphs:
        if placeholder in p.text:
            font_name = "Arial"
            font_size = Pt(11)
            if p.runs:
                font_name = p.runs[0].font.name or "Arial"
                font_size = p.runs[0].font.size or Pt(11)
                
            p.text = ""
            if not items_list:
                run = p.add_run("Ninguno.")
                run.font.name = font_name
                run.font.size = font_size
                return
            
            run = p.add_run("• " + items_list[0])
            run.font.name = font_name
            run.font.size = font_size
            
            current_p = p
            for item in items_list[1:]:
                new_p = OxmlElement('w:p')
                current_p._p.addnext(new_p)
                new_para = Paragraph(new_p, cell)
                
                new_para.paragraph_format.line_spacing = p.paragraph_format.line_spacing
                new_para.paragraph_format.space_after = p.paragraph_format.space_after
                new_para.paragraph_format.left_indent = p.paragraph_format.left_indent
                
                run_new = new_para.add_run("• " + item)
                run_new.font.name = font_name
                run_new.font.size = font_size
                current_p = new_para
            return

# --- PROCESAMIENTO INTELIGENTE DE REMISIONES ---
def procesar_remisiones_en_celda(cell, remisiones_text):
    for p in cell.paragraphs:
        if "{{Remisiones}}" in p.text:
            font_name = "Arial"
            font_size = Pt(11)
            if p.runs:
                font_name = p.runs[0].font.name or "Arial"
                font_size = p.runs[0].font.size or Pt(11)
                
            p.text = ""
            
            if es_vacio_o_negativo(remisiones_text):
                run_label = p.add_run("remisiones: ")
                run_label.bold = True
                run_label.font.name = font_name
                run_label.font.size = font_size
                
                run_val = p.add_run("no")
                run_val.font.name = font_name
                run_val.font.size = font_size
                return
            else:
                run_label = p.add_run("remisiones:")
                run_label.bold = True
                run_label.font.name = font_name
                run_label.font.size = font_size
                
                remisiones_lista = []
                for r in re.split(r'\n|,|;', remisiones_text):
                    r_clean = r.strip(" .-_/*")
                    if r_clean and not es_vacio_o_negativo(r_clean):
                        remisiones_lista.append(a_caso_oracion(r_clean))
                
                if not remisiones_lista:
                    run_val = p.add_run(" no")
                    run_val.font.name = font_name
                    run_val.font.size = font_size
                    return
                
                current_p = p
                for item in remisiones_lista:
                    new_p = OxmlElement('w:p')
                    current_p._p.addnext(new_p)
                    new_para = Paragraph(new_p, cell)
                    
                    new_para.paragraph_format.line_spacing = p.paragraph_format.line_spacing
                    new_para.paragraph_format.space_after = p.paragraph_format.space_after
                    new_para.paragraph_format.left_indent = p.paragraph_format.left_indent
                    
                    run_new = new_para.add_run("• " + item)
                    run_new.font.name = font_name
                    run_new.font.size = font_size
                    current_p = new_para
                return

def cargar_plantilla_base(archivo_cargado):
    if archivo_cargado:
        return Document(archivo_cargado)
    elif os.path.exists("FORMATO RECOMENDACIONES MEDICAS BOT.docx"):
        return Document("FORMATO RECOMENDACIONES MEDICAS BOT.docx")
    return None

# --- BANNER DE BIENVENIDA (AQUÍ ADENTRO SÓLO LOGUEADO) ---
st.markdown(f"""
    <div class='header-banner'>
        <h1 style='margin:0; font-size:26px; color:#ffffff;'>🩺 Automatización de Medicina Preventiva</h1>
        <p style='margin:5px 0 0 0; opacity:0.9; font-size:14px; color:#ffffff;'>JER S.A. · Generador Inteligente de Comunicaciones de Recomendaciones Ocupacionales</p>
    </div>
""", unsafe_allow_html=True)

# --- PANEL LATERAL CON ESTILO ---
st.sidebar.markdown(f"<h3 style='color:#1f4e79;'>👤 Perfil Activo</h3>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='metric-card'><strong>Usuario:</strong> {st.session_state.username}</div>", unsafe_allow_html=True)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.documentos = {}
    st.session_state.textos_raw = {}
    st.session_state.zip_bytes = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔗 Configuración de Nube")
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

# --- TABLERO PRINCIPAL EN 2 COLUMNAS ---
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.markdown("<h3 style='color:#1f4e79;'>📂 1. Carga de Documentos</h3>", unsafe_allow_html=True)
    pdfs_subidos = st.file_uploader("Arrastra aquí los archivos PDF médicos:", type="pdf", accept_multiple_files=True)
    
    if pdfs_subidos:
        for pdf in pdfs_subidos:
            if pdf.name not in st.session_state.documentos:
                with pdfplumber.open(pdf) as p_file:
                    texto_raw = ""
                    for page in p_file.pages:
                        texto_raw += page.extract_text() + "\n"
                datos = analizar_pdf_inteligente(texto_raw)
                st.session_state.documentos[pdf.name] = datos
                st.session_state.textos_raw[pdf.name] = texto_raw
        
        st.markdown(f"""
            <div style='display:flex; gap:10px; margin-top:15px;'>
                <div class='metric-card' style='flex:1;'><strong>PDFs Leídos</strong><br><span style='font-size:20px; font-weight:700; color:#1f4e79;'>{len(st.session_state.documentos)}</span></div>
                <div class='metric-card' style='flex:1;'><strong>Nube Sheets</strong><br><span style='font-size:14px; font-weight:600; color:green;'>{'Conectado' if g_url_guardada else 'Modo Respaldo'}</span></div>
            </div>
        """, unsafe_allow_html=True)
        
        lista_trabajadores = list(st.session_state.documentos.keys())
        archivo_seleccionado = st.selectbox("🎯 Elige un colaborador para auditar:", lista_trabajadores)
    else:
        st.session_state.documentos = {}
        st.session_state.textos_raw = {}
        archivo_seleccionado = None

    if archivo_seleccionado:
        with st.expander("📄 Auditoría de texto extraído del PDF"):
            st.text_area("Texto de origen:", value=st.session_state.textos_raw[archivo_seleccionado], height=250)

with col_der:
    st.markdown("<h3 style='color:#1f4e79;'>📋 2. Editor del Trabajador Seleccionado</h3>", unsafe_allow_html=True)
    
    if archivo_seleccionado:
        doc_actual = st.session_state.documentos[archivo_seleccionado]
        
        with st.container():
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                lugar = st.text_input("Lugar de Expedición:", value="Tunja", key=f"lugar_{archivo_seleccionado}")
            with col_f2:
                fecha = st.date_input("Fecha de la Carta:", value=datetime.date(2026, 7, 14), key=f"fecha_{archivo_seleccionado}")

            tipo_examen = st.text_input("Tipo de Examen (ASUNTO):", value=doc_actual["tipo_examen"].upper(), key=f"tipo_{archivo_seleccionado}")
            
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                nombre_persona = st.text_input("Nombre del Trabajador:", value=doc_actual["nombre"], key=f"nombre_{archivo_seleccionado}")
            with col_p2:
                cargo_persona = st.text_input("Cargo del Trabajador:", value=doc_actual["cargo"], key=f"cargo_{archivo_seleccionado}")
        
        tab1, tab2, tab3 = st.tabs(["🔬 Historial Clínico", "📋 Plan de Cuidado", "⚠️ Observación y Alertas"])
        
        with tab1:
            examenes_unificados = "\n".join(doc_actual["examenes_lista"])
            examenes_realizados = st.text_area("Exámenes Realizados (Uno por línea):", value=examenes_unificados, key=f"ex_{archivo_seleccionado}", height=130)
            
        with tab2:
            recom_unificadas = "; ".join(doc_actual["recomendaciones_lista"])
            recom_medicas = st.text_area("Recomendaciones Médicas:", value=recom_unificadas, key=f"recom_{archivo_seleccionado}", height=100)
            
            vigilancia_unificada = "; ".join(doc_actual["vigilancia_lista"])
            vigilancia = st.text_area("Programa de Vigilancia (PVE):", value=vigilancia_unificada, key=f"vig_{archivo_seleccionado}", height=80)
            
        with tab3:
            observaciones = st.text_area("Observaciones Generales:", value=doc_actual["observaciones"], key=f"obs_{archivo_seleccionado}", height=80)
            remisiones = st.text_input("Remisiones Especializadas:", value=doc_actual["remisiones"], key=f"rem_{archivo_seleccionado}")

        # Guardar cambios en caliente
        doc_actual["nombre"] = nombre_persona
        doc_actual["cargo"] = cargo_persona
        doc_actual["tipo_examen"] = tipo_examen
        doc_actual["examenes_lista"] = [linea.strip() for linea in examenes_realizados.split('\n') if linea.strip()]
        doc_actual["recomendaciones"] = a_caso_oracion(recom_medicas)
        doc_actual["vigilancia"] = a_caso_oracion(vigilancia)
        doc_actual["observaciones"] = a_caso_oracion(observaciones)
        doc_actual["remisiones"] = remisiones

        # --- CONSTRUCTOR DE WORD ---
        def generar_word_unico(datos_trabajador):
            doc_word = cargar_plantilla_base(template_uploaded)
            if not doc_word:
                st.error("Plantilla no encontrada.")
                return None
            
            consecutivo_final = datos_trabajador.get("consecutivo", "")
            if not consecutivo_final:
                g_url = obtener_config("google_sheets_url")
                if g_url:
                    try:
                        params = {
                            "name": datos_trabajador["nombre"],
                            "cargo": datos_trabajador["cargo"],
                            "examen": datos_trabajador["tipo_examen"],
                            "fecha": fecha.strftime("%Y-%m-%d")
                        }
                        r = requests.get(g_url, params=params, timeout=12)
                        data = r.json()
                        if data.get("status") == "success":
                            consecutivo_final = data.get("consecutive")
                        else:
                            consecutivo_final = incrementar_consecutivo_local()
                    except:
                        consecutivo_final = incrementar_consecutivo_local()
                else:
                    consecutivo_final = incrementar_consecutivo_local()
                datos_trabajador["consecutivo"] = consecutivo_final

            fecha_formateada = fecha.strftime("%d de %B de %Y")
            meses = {"January": "enero", "February": "febrero", "March": "marzo", "April": "abril", "May": "mayo", "June": "junio", "July": "julio", "August": "agosto", "September": "septiembre", "October": "octubre", "November": "noviembre", "December": "diciembre"}
            for eng, esp in meses.items():
                fecha_formateada = fecha_formateada.replace(eng, esp)

            replacements = {
                "{{NUMERO DE CONSECUTIVO}}": consecutivo_final,
                "{{TIPO DE EXAMEN}}": datos_trabajador["tipo_examen"].upper(),
                "{{LUGAR}}": lugar,
                "{{FECHA HOY}}": fecha_formateada,
                "{{NOMBRE DE LA PERSONA}}": datos_trabajador["nombre"],
                "{{CARGO DE LA PERSONA}}": datos_trabajador["cargo"],
                "{{Recomendaciones médicas}}": a_caso_oracion(datos_trabajador["recomendaciones"]) if datos_trabajador["recomendaciones"] else "No registra.",
                "{{Programa de vigilancia epidemiológica}}": a_caso_oracion(datos_trabajador["vigilancia"]) if datos_trabajador["vigilancia"] else "Ninguno.",
                "{{Observaciones}}": a_caso_oracion(datos_trabajador["observaciones"]) if datos_trabajador["observaciones"] else "Ninguna."
            }

            for table in doc_word.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, val in replacements.items():
                                replace_in_paragraph(p, key, val)
                        replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS}}", datos_trabajador["examenes_lista"])
                        replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS", datos_trabajador["examenes_lista"])
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
            b_io.seek(0)
            return b_io.getvalue()

        # --- SECCIÓN DE ACCIONES DE DESCARGA INTERACTIVAS ---
        st.markdown("<br>", unsafe_allow_html=True)
        col_gen1, col_gen2 = st.columns(2)
        
        with col_gen1:
            if st.button("✨ Generar Word"):
                with st.spinner("Compilando plantilla corporativa..."):
                    bytes_word = generar_word_unico(doc_actual)
                    if bytes_word:
                        st.download_button(
                            label=f"📥 Bajar Word de {doc_actual['nombre']}",
                            data=bytes_word,
                            file_name=f"Recomendaciones_{doc_actual['nombre'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
        with col_gen2:
            if len(st.session_state.documentos) > 1:
                if st.button("📦 Empaquetar todo en ZIP"):
                    with st.spinner("Sincronizando masivamente con la nube..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for filename, datos_trab in st.session_state.documentos.items():
                                bytes_word = generar_word_unico(datos_trab)
                                if bytes_word:
                                    archivo_word_nombre = f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.docx"
                                    zf.writestr(archivo_word_nombre, bytes_word)
                                    
                        zip_buffer.seek(0)
                        st.session_state.zip_bytes = zip_buffer.getvalue()
                        st.success(f"🎉 ZIP listo para descargar.")
                        
                if st.session_state.zip_bytes:
                    st.download_button(
                        label="📥 Descargar ZIP Masivo",
                        data=st.session_state.zip_bytes,
                        file_name=f"Recomendaciones_SST_JER_SA_{fecha.strftime('%Y%m%d')}.zip",
                        mime="application/zip"
                    )
    else:
        st.markdown("""
            <div style='text-align:center; padding: 40px; color:#64748b;'>
                <h3>👋 Sistema Autónomo Listo</h3>
                <p>Por favor, carga uno o varios exámenes médicos en PDF a la izquierda para desplegar el editor dinámico.</p>
            </div>
        """, unsafe_allow_html=True)
