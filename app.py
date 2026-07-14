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
import requests

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
if "word_bytes" not in st.session_state:
    st.session_state.word_bytes = None
if "doc_listo" not in st.session_state:
    st.session_state.doc_listo = False

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

# --- FUNCIONES DE EXTRACCIÓN MEJORADAS (MÉTODO HEURÍSTICO) ---
def limpiar_campo(texto):
    if not texto: return ""
    # Detener la captura si se encuentran palabras claves de formularios en la misma línea
    exclusiones = r'\b(Teléfono|Telefono|Tel|C\.C|CC|Documento|Identificac|Cedula|Cédula|Edad|Sexo|Cargo|Fecha|Estado|Empresa|Ciudad)\b'
    partes = re.split(exclusiones, texto, flags=re.IGNORECASE)
    texto_limpio = partes[0]
    texto_limpio = re.sub(r'[:\-,_]+', '', texto_limpio)
    return texto_limpio.strip()

def extraer_seccion(texto, palabras_inicio, palabras_fin):
    lineas = texto.split('\n')
    seccion = []
    dentro = False
    for linea in lineas:
        linea_upper = linea.upper().strip()
        if not dentro:
            if any(h in linea_upper for h in palabras_inicio):
                dentro = True
                for h in palabras_inicio:
                    if h in linea_upper:
                        idx = linea_upper.find(h) + len(h)
                        resto = linea[idx:].strip(" :-,_")
                        if resto: seccion.append(resto)
                        break
        else:
            if any(h in linea_upper for h in palabras_fin):
                break
            seccion.append(linea.strip())
    return "\n".join(seccion).strip()

def analizar_pdf_inteligente(texto):
    datos = {
        "nombre": "", "cargo": "", "tipo_examen": "",
        "examenes": "", "recomendaciones": "", "vigilancia": "",
        "observaciones": "", "remisiones": ""
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
        datos["tipo_examen"] = limpiar_campo(m_tipo.group(1))
    else:
        for palabra in ["INGRESO", "PERIÓDICO", "PERIODICO", "EGRESO", "RETIRO", "CAMBIO DE CARGO", "POST-INCAPACIDAD", "POST INCAPACIDAD"]:
            if palabra in texto.upper():
                datos["tipo_examen"] = palabra
                break

    # 4. Exámenes Realizados (Heurística de términos comunes)
    examenes_comunes = ["Visiometría", "Visiometria", "Audiometría", "Audiometria", "Espiometría", "Espirometria", "Frotis", "Cuadro Hemático", "Optometría", "Laboratorio Clínico", "Glicemia", "Colesterol"]
    detectados = [ex for ex in examenes_comunes if re.search(r'\b' + re.escape(ex) + r'\b', texto, re.IGNORECASE)]
    datos["examenes"] = ", ".join(detectados) if detectados else ""

    # 5. Recomendaciones Médicas (Bloque multilínea)
    datos["recomendaciones"] = extraer_seccion(texto, 
        ["RECOMENDACIONES MEDICAS", "RECOMENDACIONES:", "INDICACIONES MEDICAS"],
        ["OBSERVACIONES", "REMISIONES", "SISTEMA DE VIGILANCIA", "PVE", "VIGILANCIA", "FIRMA", "ATENTAMENTE"]
    )

    # 6. Vigilancia Epidemiológica (PVE)
    datos["vigilancia"] = extraer_seccion(texto,
        ["SISTEMA DE VIGILANCIA", "VIGILANCIA EPIDEMIOLOGICA", "PVE", "SVE", "VIGILANCIA:"],
        ["RECOMENDACIONES", "OBSERVACIONES", "REMISIONES", "FIRMA", "ATENTAMENTE"]
    )

    # 7. Observaciones
    datos["observaciones"] = extraer_seccion(texto,
        ["OBSERVACIONES:", "OBSERVACION:", "OBSERVACIONES"],
        ["RECOMENDACIONES", "REMISIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE"]
    )

    # 8. Remisiones
    datos["remisiones"] = extraer_seccion(texto,
        ["REMISIONES:", "REMISION:", "REMISIONES", "REMITIDO A:"],
        ["RECOMENDACIONES", "OBSERVACIONES", "VIGILANCIA", "FIRMA", "ATENTAMENTE"]
    )

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

# --- BARRA LATERAL (CONFIGURACIÓN) ---
st.sidebar.markdown(f"👤 **Usuario Activo:** {st.session_state.username}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.doc_listo = False
    st.rerun()

st.sidebar.markdown("---")

# Guardar URL de Google Sheets
st.sidebar.subheader("🔗 Conexión Google Sheets")
g_url_guardada = obtener_config("google_sheets_url")
g_url_input = st.sidebar.text_input("URL de Google Apps Script:", value=g_url_guardada, type="password")

if st.sidebar.button("Guardar Conexión"):
    guardar_config("google_sheets_url", g_url_input)
    st.sidebar.success("¡URL de Google Sheets guardada!")
    st.rerun()

if g_url_guardada:
    st.sidebar.markdown("🟢 **Estado:** Enlazado con Google Sheets")
else:
    st.sidebar.markdown("🟡 **Estado:** Modo Local (Sin Google Sheets)")

st.sidebar.markdown("---")
st.sidebar.subheader("✍️ Cargar Firma Autorizada")
firma_file = st.sidebar.file_uploader("Sube la firma de Víctor (.png / .jpg)", type=["png", "jpg"])

# --- COLUMNAS DE TRABAJO ---
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.subheader("📁 1. Cargar PDF Médico")
    pdf_subido = st.file_uploader("Arrastra el examen en PDF", type="pdf")
    
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
        datos_sugeridos = analizar_pdf_inteligente(texto_raw)
        
        with st.expander("🔍 Ver texto extraído del PDF"):
            st.text_area("Texto crudo:", value=texto_raw, height=300)

with col_der:
    st.subheader("📋 2. Formato de Envío[cite: 2]")
    st.write("Edita la información antes de generar tu documento final:")
    
    consecutivo_actual = obtener_config("google_sheets_url")
    if consecutivo_actual:
        st.caption("ℹ️ El consecutivo se reservará automáticamente en tu Google Sheet al hacer clic en 'Generar Word'.")
    else:
        num_local = obtener_siguiente_consecutivo_local()
        st.info(f"ℹ️ Consecutivo local de respaldo estimado: `SST-2026-{num_local}`")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        lugar = st.text_input("Lugar de Expedición[cite: 2]:", value="Tunja")
    with col_f2:
        fecha = st.date_input("Fecha de la Carta[cite: 2]:", value=datetime.date(2026, 7, 14))

    tipo_examen = st.text_input("Tipo de Examen (ASUNTO)[cite: 2]:", value=datos_sugeridos["tipo_examen"].upper(), placeholder="Ej: PERIÓDICO")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        nombre_persona = st.text_input("Nombre del Trabajador[cite: 2]:", value=datos_sugeridos["nombre"])
    with col_p2:
        cargo_persona = st.text_input("Cargo del Trabajador[cite: 2]:", value=datos_sugeridos["cargo"])
        
    examenes_realizados = st.text_area("Exámenes Realizados[cite: 2]:", value=datos_sugeridos["examenes"], placeholder="Ej: Audiometría, Visiometría, Laboratorios...")
    
    recom_medicas = st.text_area("Recomendaciones Médicas[cite: 2]:", value=datos_sugeridos["recomendaciones"])
    vigilancia = st.text_area("Programa de Vigilancia Epidemiológica (PVE)[cite: 2]:", value=datos_sugeridos["vigilancia"])
    
    observaciones = st.text_area("Observaciones[cite: 2]:", value=datos_sugeridos["observaciones"])
    remisiones = st.text_area("Remisiones[cite: 2]:", value=datos_sugeridos["remisiones"])

    # Botón para Procesar y Estructurar
    if st.button("✨ Generar Word"):
        with st.spinner("Reservando consecutivo y maquetando..."):
            consecutivo_final = ""
            g_url = obtener_config("google_sheets_url")
            
            # Intentar conexión con Google Sheets
            if g_url:
                try:
                    params = {
                        "name": nombre_persona,
                        "cargo": cargo_persona,
                        "examen": tipo_examen,
                        "fecha": fecha.strftime("%Y-%m-%d")
                    }
                    r = requests.get(g_url, params=params, timeout=12)
                    data = r.json()
                    if data.get("status") == "success":
                        consecutivo_final = data.get("consecutive")
                        st.success(f"🟢 Consecutivo '{consecutivo_final}' registrado con éxito en Google Sheets.")
                    else:
                        st.warning("⚠️ Error en Google Sheets. Se usará el consecutivo local.")
                        consecutivo_final = incrementar_consecutivo_local()
                except Exception as e:
                    st.warning(f"⚠️ Sin conexión a Google Sheets ({e}). Consecutivo local asignado.")
                    consecutivo_final = incrementar_consecutivo_local()
            else:
                consecutivo_final = incrementar_consecutivo_local()

            # --- CONSTRUCCIÓN DEL WORD ---
            try:
                doc = Document()
                style = doc.styles['Normal']
                style.font.name = 'Arial'
                style.font.size = Pt(11)
                
                for section in doc.sections:
                    section.top_margin = Inches(1)
                    section.bottom_margin = Inches(1)
                    section.left_margin = Inches(1)
                    section.right_margin = Inches(1)

                # 1. Consecutivo[cite: 2]
                p_cons = doc.add_paragraph()
                p_cons.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                run_cons = p_cons.add_run(f"Consecutivo: {consecutivo_final}")[cite: 2]
                run_cons.bold = True
                
                doc.add_paragraph()
                
                # 2. Asunto[cite: 2]
                p_asunto = doc.add_paragraph()
                p_asunto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_asunto = p_asunto.add_run(f"ASUNTO: RECOMENDACIONES EXAMEN {tipo_examen.upper()}")[cite: 2]
                run_asunto.bold = True
                
                doc.add_paragraph()
                
                # 3. Lugar y Fecha[cite: 2]
                p_fecha = doc.add_paragraph()
                fecha_formateada = fecha.strftime("%d de %B de %Y")
                meses = {
                    "January": "enero", "February": "febrero", "March": "marzo", "April": "abril",
                    "May": "mayo", "June": "junio", "July": "julio", "August": "agosto",
                    "September": "septiembre", "October": "octubre", "November": "noviembre", "December": "diciembre"
                }
                for eng, esp in meses.items():
                    fecha_formateada = fecha_formateada.replace(eng, esp)
                p_fecha.add_run(f"{lugar}, {fecha_formateada}")[cite: 2]
                
                doc.add_paragraph()
                
                # 4. Destinatario[cite: 2]
                doc.add_paragraph("Sr(a).")[cite: 2]
                p_nom = doc.add_paragraph()
                p_nom.add_run(nombre_persona).bold = True[cite: 2]
                doc.add_paragraph(cargo_persona)[cite: 2]
                
                doc.add_paragraph()
                
                # 5. Saludo y Cuerpo[cite: 2]
                doc.add_paragraph("Cordial saludo,")[cite: 2]
                p_cuerpo = doc.add_paragraph(
                    "Según los lineamientos del programa de medicina preventiva y del trabajo de JER S.A; "
                    "se hace entrega de las recomendaciones establecidas por el Proveedor de servicios de "
                    "Exámenes Médico Ocupacionales (Ingreso, Periódico, egreso, cambio de cargo y post incapacidad)"[cite: 2]
                )
                p_cuerpo.paragraph_format.line_spacing = 1.15
                
                doc.add_paragraph()
                
                # 6. Exámenes[cite: 2]
                p_ex = doc.add_paragraph()
                p_ex.add_run("EXÁMENES REALIZADOS:").bold = True[cite: 2]
                doc.add_paragraph(examenes_realizados if examenes_realizados else "Ninguno registrado.")
                
                doc.add_paragraph()
                
                # 7. Recomendaciones[cite: 2]
                p_recom = doc.add_paragraph()
                p_recom.add_run("Recomendaciones: ").bold = True[cite: 2]
                p_recom.add_run(recom_medicas if recom_medicas else "No registra.")[cite: 2]
                
                if vigilancia:
                    p_recom.add_run("\nPrograma de vigilancia epidemiológica: ").bold = True[cite: 2]
                    p_recom.add_run(vigilancia)[cite: 2]
                    
                # 8. Observaciones[cite: 2]
                p_obs = doc.add_paragraph()
                p_obs.add_run("observaciones: ").bold = True[cite: 2]
                p_obs.add_run(observaciones if observaciones else "Ninguna.")[cite: 2]
                
                # 9. Remisiones[cite: 2]
                p_rem = doc.add_paragraph()
                p_rem.add_run("remisiones: ").bold = True[cite: 2]
                p_rem.add_run(remisiones if remisiones else "Ninguna.")[cite: 2]
                
                doc.add_paragraph()
                
                # 10. Firma[cite: 2]
                doc.add_paragraph("Atentamente,")[cite: 2]
                if firma_file:
                    doc.add_picture(firma_file, width=Inches(2.2))
                else:
                    doc.add_paragraph()
                    
                p_firma_nombre = doc.add_paragraph()
                p_firma_nombre.add_run("VÍCTOR ALONSO MORENO CASAS").bold = True[cite: 2]
                doc.add_paragraph("Coordinador SST")[cite: 2]
                
                # Guardar en memoria
                b_io = io.BytesIO()
                doc.save(b_io)
                b_io.seek(0)
                
                # Guardamos las variables de estado
                st.session_state.word_bytes = b_io.getvalue()
                st.session_state.doc_listo = True
                st.session_state.nombre_archivo = f"Recomendaciones_{nombre_persona.replace(' ', '_')}.docx"
                
            except Exception as e:
                st.error(f"Error al estructurar el Word: {e}")

    # BOTÓN DE DESCARGA: Se muestra de manera fija al finalizar la maquetación
    if st.session_state.doc_listo and st.session_state.word_bytes:
        st.markdown("---")
        st.success("🎉 ¡Maquetación finalizada con éxito!")
        st.download_button(
            label="📥 Descargar Documento Word (.docx)",
            data=st.session_state.word_bytes,
            file_name=st.session_state.nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
