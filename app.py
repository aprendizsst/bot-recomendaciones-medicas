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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Cartas SST - JER S.A.", page_icon="🩺", layout="wide")

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
        max-width: 480px;
        margin: 60px auto;
        padding: 30px;
        background-color: white;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
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
    st.markdown("<div class='login-container'>", unsafe_allow_html=True)
    st.subheader("🩺 Gestión de Acceso - JER S.A.")
    
    if not tiene_usuarios():
        st.info("🆕 Bienvenido. Configura tu cuenta inicial de Administrador.")
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
        opcion_acceso = st.radio("Acción:", ["Iniciar Sesión", "Crear Nueva Cuenta", "Actualizar Contraseña"], horizontal=True)
        st.markdown("---")
        
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
                        st.success("🎉 Cuenta creada. Ve a 'Iniciar Sesión'.")
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

# --- FUNCIÓN DE FORMATO: CASO ORACIÓN (SENTENCE CASE) ---
def a_caso_oracion(texto):
    if not texto: 
        return ""
    texto_min = texto.lower().strip()
    
    def capitalizar_match(match):
        return match.group(1) + match.group(2).upper()
    
    texto_formateado = re.sub(r'(^|[.!?]\s+|\n+)([a-zñáéíóúü])', capitalizar_match, texto_min)
    return texto_formateado

# --- VALIDACIÓN DE VALOR VACÍO O NEGATIVO ---
def es_vacio_o_negativo(texto):
    if not texto:
        return True
    t_clean = texto.strip().lower().strip(" .-_/")
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
        "observaciones": "", "remisiones": "", "consecutivo": ""
    }
    if not texto: return datos

    # 1. Nombre
    m_nom = re.search(r'(?:Nombre|Paciente|Colaborador|Trabajador):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_nom: datos["nombre"] = limpiar_campo(m_nom.group(1))

    # 2. Cargo
    m_car = re.search(r'(?:Cargo|Ocupación|Ocupacion|Puesto):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_car: datos["cargo"] = limpiar_campo(m_car.group(1))

    # 3. Tipo de Examen
    m_tipo = re.search(r'(?:Tipo de Examen|Concepto|Evaluación|Evaluacion|Motivo|Clase de Examen):\s*([^\n]+)', texto, re.IGNORECASE)
    if m_tipo: 
        datos["tipo_examen"] = limpiar_campo(m_tipo.group(1)).upper()
    else:
        for palabra in ["INGRESO", "PERIÓDICO", "PERIODICO", "EGRESO", "RETIRO", "CAMBIO DE CARGO", "POST-INCAPACIDAD", "POST INCAPACIDAD"]:
            if palabra in texto.upper():
                datos["tipo_examen"] = palabra
                break

    # --- MAPA DE EXÁMENES CLAVE ---
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

    # Extraer observaciones normales
    datos["observaciones"] = a_caso_oracion(extraer_seccion(texto,
        ["OBSERVACIONES:", "OBSERVACION:", "OBSERVACIONES"],
        ["RECOMENDACIONES", "REMISIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE"]
    ))

    # --- EXTRACCIÓN DE REMISIONES ESTRICTA DESDE LA TABLA DEL PROVEEDOR ---
    datos["remisiones"] = a_caso_oracion(extraer_seccion(texto,
        ["INFORMACION DE REMISIONES", "INFORMACIÓN DE REMISIONES"],
        ["RECOMENDACIONES", "OBSERVACIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE", "DIAGNOSTICOS", "DIAGNÓSTICOS"]
    ))

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

# --- REMPLAZO EN PÁRRAFOS ---
def replace_in_paragraph(paragraph, key, value):
    if key not in paragraph.text:
        return
    replaced_in_runs = False
    for run in paragraph.runs:
        if key in run.text:
            run.text = run.text.replace(key, value)
            replaced_in_runs = True
    if not replaced_in_runs:
        paragraph.text = paragraph.text.replace(key, value)

# --- REEMPLAZO DINÁMICO POR VIÑETAS NATIVAS EN WORD (A PRUEBA DE CRASHES) ---
def replace_placeholder_with_bullets(cell, placeholder, items_list):
    for p in cell.paragraphs:
        if placeholder in p.text:
            p.text = ""
            if not items_list:
                p.text = "Ninguno."
                return
            
            try:
                p.style = 'List Bullet'
                p.add_run(items_list[0])
            except KeyError:
                p.add_run("• " + items_list[0])
            
            current_p = p
            for item in items_list[1:]:
                new_p = OxmlElement('w:p')
                current_p._p.addnext(new_p)
                new_para = Paragraph(new_p, cell)
                try:
                    new_para.style = 'List Bullet'
                    new_para.add_run(item)
                except KeyError:
                    new_para.add_run("• " + item)
                current_p = new_para
            return

# --- PROCESAMIENTO INTELIGENTE DE REMISIONES ---
def procesar_remisiones_en_celda(cell, remisiones_text):
    for p in cell.paragraphs:
        if "{{Remisiones}}" in p.text:
            if es_vacio_o_negativo(remisiones_text):
                p.text = ""
                return
            else:
                p.text = ""
                run = p.add_run("remisiones:")
                run.bold = True
                
                remisiones_lista = []
                for r in re.split(r'\n|,|;', remisiones_text):
                    r_clean = r.strip(" .-_/*")
                    if r_clean and not es_vacio_o_negativo(r_clean):
                        remisiones_lista.append(a_caso_oracion(r_clean))
                
                if not remisiones_lista:
                    p.text = ""
                    return
                
                current_p = p
                for item in remisiones_lista:
                    new_p = OxmlElement('w:p')
                    current_p._p.addnext(new_p)
                    new_para = Paragraph(new_p, cell)
                    try:
                        new_para.style = 'List Bullet'
                        new_para.add_run(item)
                    except KeyError:
                        new_para.add_run("• " + item)
                    current_p = new_para
                return

# --- CONFIGURACIÓN DE PLANTILLA WORD ---
def cargar_plantilla_base(archivo_cargado):
    if archivo_cargado:
        return Document(archivo_cargado)
    elif os.path.exists("FORMATO RECOMENDACIONES MEDICAS BOT.docx"):
        return Document("FORMATO RECOMENDACIONES MEDICAS BOT.docx")
    return None

# --- BARRA LATERAL ---
st.sidebar.markdown(f"👤 **Usuario Activo:** {st.session_state.username}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.documentos = {}
    st.session_state.textos_raw = {}
    st.session_state.zip_bytes = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔗 Conexión Google Sheets")
g_url_guardada = obtener_config("google_sheets_url")
g_url_input = st.sidebar.text_input("URL de Google Apps Script:", value=g_url_guardada, type="password")
if st.sidebar.button("Guardar Conexión"):
    guardar_config("google_sheets_url", g_url_input)
    st.sidebar.success("¡URL de Google Sheets guardada!")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📄 Subir Plantilla Corporativa")
template_uploaded = st.sidebar.file_uploader("Sube el formato de Word original (.docx)", type=["docx"])

st.sidebar.markdown("---")
st.sidebar.subheader("✍️ Cargar Firma Autorizada")
firma_file = st.sidebar.file_uploader("Sube la firma de Víctor (.png / .jpg)", type=["png", "jpg"])

# --- COLUMNAS DE TRABAJO ---
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.subheader("📁 1. Carga Masiva de PDFs")
    pdfs_subidos = st.file_uploader("Arrastra aquí uno o varios PDFs", type="pdf", accept_multiple_files=True)
    
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
        
        st.success(f"🟢 ¡Se cargaron {len(st.session_state.documentos)} trabajadores con éxito!")
        lista_trabajadores = list(st.session_state.documentos.keys())
        archivo_seleccionado = st.selectbox("📝 Selecciona un trabajador para editar:", lista_trabajadores)
    else:
        st.session_state.documentos = {}
        st.session_state.textos_raw = {}
        archivo_seleccionado = None

    if archivo_seleccionado:
        with st.expander("🔍 Ver texto crudo de este PDF"):
            st.text_area("Texto:", value=st.session_state.textos_raw[archivo_seleccionado], height=250)

with col_der:
    st.subheader("📋 2. Editor del Trabajador Seleccionado")
    
    if archivo_seleccionado:
        doc_actual = st.session_state.documentos[archivo_seleccionado]
        
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
            
        examenes_unificados = "\n".join(doc_actual["examenes_lista"])
        examenes_realizados = st.text_area("Exámenes Realizados (Uno por línea para viñetas en Word):", value=examenes_unificados, key=f"ex_{archivo_seleccionado}", height=120)
        
        recom_unificadas = "; ".join(doc_actual["recomendaciones_lista"])
        recom_medicas = st.text_area("Recomendaciones Médicas:", value=recom_unificadas, key=f"recom_{archivo_seleccionado}")
        
        vigilancia_unificada = "; ".join(doc_actual["vigilancia_lista"])
        vigilancia = st.text_area("Programa de Vigilancia Epidemiológica (PVE):", value=vigilancia_unificada, key=f"vig_{archivo_seleccionado}")
        
        observaciones = st.text_area("Observaciones:", value=doc_actual["observaciones"], key=f"obs_{archivo_seleccionado}")
        
        remisiones = st.text_area("Remisiones (Escribe 'Ninguna' o déjalo vacío para ocultarlo de la carta):", value=doc_actual["remisiones"], key=f"rem_{archivo_seleccionado}")

        # Guardar cambios en memoria en tiempo real
        doc_actual["nombre"] = nombre_persona
        doc_actual["cargo"] = cargo_persona
        doc_actual["tipo_examen"] = tipo_examen
        doc_actual["examenes_lista"] = [linea.strip() for linea in examenes_realizados.split('\n') if linea.strip()]
        doc_actual["recomendaciones"] = a_caso_oracion(recom_medicas)
        doc_actual["vigilancia"] = a_caso_oracion(vigilancia)
        doc_actual["observaciones"] = a_caso_oracion(observaciones)
        doc_actual["remisiones"] = a_caso_oracion(remisiones)

        # --- CONSTRUCTOR DE WORD ---
        def generar_word_unico(datos_trabajador):
            doc_word = cargar_plantilla_base(template_uploaded)
            if not doc_word:
                st.error("No se encontró la plantilla Word en el sistema.")
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
            meses = {
                "January": "enero", "February": "febrero", "March": "marzo", "April": "abril",
                "May": "mayo", "June": "junio", "July": "julio", "August": "agosto",
                "September": "septiembre", "October": "octubre", "November": "noviembre", "December": "diciembre"
            }
            for eng, esp in meses.items():
                fecha_formateada = fecha_formateada.replace(eng, esp)

            # Diccionario de Reemplazos Simples de Texto
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
                        # 1. Reemplazo de marcadores normales
                        for p in cell.paragraphs:
                            for key, val in replacements.items():
                                replace_in_paragraph(p, key, val)
                        
                        # 2. Reemplazo por Viñetas de Exámenes Realizados
                        replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS}}", datos_trabajador["examenes_lista"])
                        replace_placeholder_with_bullets(cell, "{{LISTA DE EXAMENES REALIZADOS", datos_trabajador["examenes_lista"])
                        
                        # 3. Procesamiento Inteligente de Remisiones (Habilitar Listas o Borrar la Sección)
                        procesar_remisiones_en_celda(cell, datos_trabajador["remisiones"])
                        
                        # 4. Estampado de Firma Autorizada
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

        # --- BOTONES DE ACCIÓN ---
        st.markdown("---")
        col_gen1, col_gen2 = st.columns(2)
        
        with col_gen1:
            if st.button("✨ Generar Word de este Trabajador"):
                with st.spinner("Procesando plantilla..."):
                    bytes_word = generar_word_unico(doc_actual)
                    if bytes_word:
                        st.success(f"🟢 Consecutivo '{doc_actual['consecutivo']}' reservado.")
                        st.download_button(
                            label=f"📥 Descargar Word: {doc_actual['nombre']}",
                            data=bytes_word,
                            file_name=f"Recomendaciones_{doc_actual['nombre'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
        with col_gen2:
            if len(st.session_state.documentos) > 1:
                if st.button("📦 Generar TODOS en un ZIP"):
                    with st.spinner("Conectando con Sheets y empaquetando..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for filename, datos_trab in st.session_state.documentos.items():
                                bytes_word = generar_word_unico(datos_trab)
                                if bytes_word:
                                    archivo_word_nombre = f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.docx"
                                    zf.writestr(archivo_word_nombre, bytes_word)
                                    
                        zip_buffer.seek(0)
                        st.session_state.zip_bytes = zip_buffer.getvalue()
                        st.success(f"🎉 ZIP creado con {len(st.session_state.documentos)} informes.")
                        
                if st.session_state.zip_bytes:
                    st.download_button(
                        label="📥 Descargar ZIP Masivo",
                        data=st.session_state.zip_bytes,
                        file_name=f"Recomendaciones_SST_JER_SA_{fecha.strftime('%Y%m%d')}.zip",
                        mime="application/zip"
                    )
    else:
        st.info("👋 Sube archivos PDF en el panel de la izquierda para comenzar.")
