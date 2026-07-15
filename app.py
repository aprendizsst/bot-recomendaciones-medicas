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
import unicodedata
import requests
import io
import zipfile
import tempfile
import subprocess
import base64

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

# --- MANEJO DE SESIÓN DE LOGIN Y ESTADOS DE ARCHIVOS ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "documentos" not in st.session_state:
    st.session_state.documentos = {}
if "pdfs_raw_bytes" not in st.session_state:
    st.session_state.pdfs_raw_bytes = {}
if "textos_raw" not in st.session_state:
    st.session_state.textos_raw = {}
if "export_bytes" not in st.session_state:
    st.session_state.export_bytes = None
if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None
if "processed_doc" not in st.session_state:
    st.session_state.processed_doc = None
if "prev_colaborador" not in st.session_state:
    st.session_state.prev_colaborador = None
if "document_count" not in st.session_state:
    st.session_state.document_count = 0

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
        r'\bregion\b': 'región', r'\bhabitos\b': 'hábitos', r'\badiministrativo\b': 'administrativo'
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

def es_vacio_o_estado(texto):
    if not texto: return True
    t_clean = texto.strip().upper()
    t_clean_norm = re.sub(r'[^A-ZÁÉÍÓÚÑ\s]', '', t_clean).strip()
    t_clean_norm = re.sub(r'\s+', ' ', t_clean_norm)
    
    frases_estado = {
        "REALIZADO", "REALZIADO", "SIN ALTERACIONES", "NORMAL", "SANO", "NEGATIVO", 
        "NO REGISTRA", "NA", "SIN REMISIONES", "SIN REMISIÓN", "VISUAL", "CARDIOVASCULAR", 
        "DME", "OSTEOMUSCULAR", "AUDITIVO", "RESPIRATORIO", "SVE", "SISTEMA", "VIGILANCIA",
        "SANO Y SIN ALTERACIONES", "NINGUNO", "NINGUNA", "NO PRESENTAS", "NO PRESENTA", 
        "NO REGISTRA RECOMENDACIONES", "NORMALES", "NORMAL", "SIN ALTERACION", "NO APLICA",
        "RECOMENDACIONES MÉDICAS", "RECOMENDACIONES OCUPACIONALES", "HABITOS Y ESTILO DE VIDA SALUDABLES",
        "HABITOS SALUDABLES", "OTRAS OBSERVACIONES Y RECOMENDACIONES", "RECOMENDACIONES MEDICAS"
    }
    
    if t_clean_norm in frases_estado:
        return True
    if len(texto.strip()) <= 3:
        return True
    return False

def limpiar_campo(texto):
    if not texto: return ""
    partes = re.split(r'\b(Teléfono|Telefono|Tel|C\.C|CC|Documento|Cedula|Cargo|Fecha)\b', texto, flags=re.IGNORECASE)
    return re.sub(r'[:\-,_]+', '', partes[0]).strip().title()

def limpiar_linea_ruido_lateral(linea):
    patron_ruido = r'\s{2,}(VISUAL|DME|CARDIOVASCULAR|SVE|AUDITIVO|RESPIRATORIO|SISTEMA|VIGILANCIA)\s*$'
    linea_limpia = re.sub(patron_ruido, '', linea, flags=re.IGNORECASE)
    return linea_limpia.strip()

def limpiar_ruido_columnas_final(texto):
    if not texto: return ""
    patrones_ruido = [
        r'\bvisual\b', r'\bdme\b', r'\bcardiovascular\b', r'\bsve\b', 
        r'\bauditivo\b', r'\brespiratorio\b', r'\bsistema\b', r'\bvigilancia\b'
    ]
    for patron in patrones_ruido:
        texto = re.sub(patron + r'\s*$', '', texto, flags=re.IGNORECASE)
    return texto.strip(" :-,_/")

def intentar_parsear_fecha(fecha_str):
    fecha_str = fecha_str.lower().strip(" :-,_/.()[]|")
    m_letras = re.search(r'(\d{1,2})\s+de\s+([a-zñáéíóúü]+)\s+de\s+(20\d{2})', fecha_str)
    if m_letras:
        dia = int(m_letras.group(1))
        mes_str = m_letras.group(2)
        anio = int(m_letras.group(3))
        meses = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }
        return datetime.date(anio, meses.get(mes_str, 1), dia)
        
    m_ymd = re.search(r'(20\d{2})[-/](\d{1,2})[-/](\d{1,2})', fecha_str)
    if m_ymd: return datetime.date(int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3)))
        
    m_dmy = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](20\d{2})', fecha_str)
    if m_dmy: return datetime.date(int(m_dmy.group(3)), int(m_dmy.group(2)), int(m_dmy.group(1)))
        
    return datetime.date.today()


# --- EXTRACCIÓN ROBUSTA DE IDENTIDAD, CARGO Y LUGAR ---
# Estas funciones complementan el analizador existente sin cambiar su flujo,
# sus campos de salida ni la generación de documentos.

_ETIQUETAS_CORTE = [
    "APELLIDOS Y NOMBRES", "NOMBRES Y APELLIDOS", "NOMBRE DEL TRABAJADOR",
    "NOMBRE TRABAJADOR", "NOMBRE COMPLETO", "TRABAJADOR", "PACIENTE",
    "CARGO ACTUAL", "CARGO", "OCUPACIÓN", "OCUPACION", "OFICIO", "PUESTO",
    "DOCUMENTO", "IDENTIFICACIÓN", "IDENTIFICACION", "CÉDULA", "CEDULA", "C.C.",
    "CC", "GÉNERO", "GENERO", "EDAD", "TELÉFONO", "TELEFONO", "CELULAR",
    "EPS", "AFP", "ARL", "EMPRESA", "NIT", "FECHA", "CIUDAD", "MUNICIPIO",
    "LUGAR", "SEDE", "DIRECCIÓN", "DIRECCION"
]

_RUIDO_IDENTIDAD = {
    "DATOS", "DATOS DEL TRABAJADOR", "INFORMACION DEL TRABAJADOR",
    "INFORMACIÓN DEL TRABAJADOR", "APELLIDOS Y NOMBRES", "NOMBRES Y APELLIDOS",
    "NOMBRE", "TRABAJADOR", "PACIENTE", "GENERO", "GÉNERO", "EDAD",
    "DOCUMENTO", "IDENTIFICACION", "IDENTIFICACIÓN", "CEDULA", "CÉDULA",
    "EMPRESA", "IPS", "EPS", "AFP", "ARL", "FIRMA", "CERTIFICADO"
}

_RUIDO_CARGO = {
    "CARGO", "CARGO ACTUAL", "OCUPACION", "OCUPACIÓN", "OFICIO", "PUESTO",
    "TRABAJADOR", "DATOS", "EMPRESA", "EPS", "AFP", "ARL", "GENERO", "GÉNERO",
    "DOCUMENTO", "IDENTIFICACION", "IDENTIFICACIÓN", "CERTIFICADO"
}

_RUIDO_LUGAR = {
    "LUGAR", "CIUDAD", "MUNICIPIO", "SEDE", "FECHA", "DIA", "DÍA", "MES",
    "AÑO", "ANO", "REALIZACION", "REALIZACIÓN", "EXAMEN", "EXÁMEN",
    "CERTIFICADO", "PAGINA", "PÁGINA", "LOGOTIPO", "AM", "PM"
}


def normalizar_etiqueta(texto):
    """Normaliza únicamente para comparar etiquetas; no modifica el valor final."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip(" :|_-./")


def dividir_columnas_estructuradas(linea):
    """Divide filas de tablas extraídas por pdfplumber sin romper nombres normales."""
    if not linea:
        return []
    columnas = [
        re.sub(r"\s+", " ", c).strip(" |/-,_.:")
        for c in re.split(r"\s{2,}|\t+|\|", linea)
    ]
    return [c for c in columnas if c]


def recortar_en_siguiente_etiqueta(valor):
    if not valor:
        return ""
    patron = r"\b(?:" + "|".join(
        sorted((re.escape(e) for e in _ETIQUETAS_CORTE), key=len, reverse=True)
    ) + r")\b"
    coincidencias = list(re.finditer(patron, valor, flags=re.IGNORECASE))
    if coincidencias:
        # Se conserva la primera parte solo cuando la siguiente etiqueta aparece
        # después de haber extraído algún valor.
        primera = coincidencias[0]
        if primera.start() > 0:
            valor = valor[:primera.start()]
    return valor.strip(" |/-,_.:")


def limpiar_candidato_campo(valor, tipo):
    if not valor:
        return ""
    valor = str(valor).replace("\x00", " ")
    valor = re.sub(r"\s+", " ", valor).strip(" |/-,_.:")
    valor = recortar_en_siguiente_etiqueta(valor)

    # Elimina identificaciones y datos laterales que suelen quedar pegados.
    valor = re.split(
        r"\b(?:C\.?\s*C\.?|CÉDULA|CEDULA|DOCUMENTO|IDENTIFICACIÓN|IDENTIFICACION|"
        r"TELÉFONO|TELEFONO|CELULAR|EDAD|GÉNERO|GENERO|EPS|AFP|ARL)\b",
        valor,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip(" |/-,_.:")

    if tipo in {"nombre", "cargo"}:
        valor = re.sub(r"\([^)]*\)", "", valor).strip()
    elif tipo == "lugar":
        # Conserva nombres de IPS entre paréntesis solo si son el único dato.
        valor = re.sub(r"\s+\([^)]*(?:PÁGINA|PAGINA|HORA|AM|PM)[^)]*\)", "", valor, flags=re.IGNORECASE)
        # En algunas tablas la línea siguiente contiene DÍA | MES | AÑO | CIUDAD.
        # Se eliminan los componentes de fecha antes de validar el municipio.
        valor = re.sub(
            r"^\s*(?:(?:20\d{2})\s*[|/\-.]\s*\d{1,2}\s*[|/\-.]\s*\d{1,2}|"
            r"\d{1,2}\s*[|/\-.]\s*\d{1,2}\s*[|/\-.]\s*20\d{2})\s*[|/\-,_.:]*\s*",
            "",
            valor,
        )
        if "|" in valor:
            partes_lugar = [
                p.strip(" |/-,_.:")
                for p in valor.split("|")
                if p.strip(" |/-,_.:")
            ]
            partes_con_letras = [
                p for p in partes_lugar
                if sum(c.isalpha() for c in p) >= 3
            ]
            if partes_con_letras:
                valor = partes_con_letras[-1]

    return re.sub(r"\s+", " ", valor).strip(" |/-,_.:")


def candidato_nombre_valido(valor):
    valor = limpiar_candidato_campo(valor, "nombre")
    if not valor or len(valor) < 5 or len(valor) > 100:
        return False
    if re.search(r"\d|@|https?://", valor):
        return False
    tokens = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'-]+", valor)
    if len(tokens) < 2 or len(tokens) > 8:
        return False
    norm = normalizar_etiqueta(valor)
    if norm in {normalizar_etiqueta(x) for x in _RUIDO_IDENTIDAD}:
        return False
    if any(
        ruido in norm
        for ruido in [
            "CERTIFICADO", "MEDICINA", "OCUPACIONAL", "EMPRESA", "INSTITUCION",
            "INSTITUCIÓN", "LABORATORIO", "CENTRO MEDICO", "CENTRO MÉDICO",
            "SALUD OCUPACIONAL", "FIRMA DEL", "NOMBRE DEL MEDICO", "NOMBRE DEL MÉDICO"
        ]
    ):
        return False
    letras = sum(c.isalpha() for c in valor)
    return letras / max(len(valor), 1) >= 0.65


def candidato_cargo_valido(valor):
    valor = limpiar_candidato_campo(valor, "cargo")
    if not valor or len(valor) < 3 or len(valor) > 120:
        return False
    norm = normalizar_etiqueta(valor)
    if norm in {normalizar_etiqueta(x) for x in _RUIDO_CARGO}:
        return False
    if any(
        ruido in norm
        for ruido in [
            "CERTIFICADO", "MEDICO OCUPACIONAL", "MÉDICO OCUPACIONAL",
            "FIRMA", "DOCUMENTO", "IDENTIFICACION", "IDENTIFICACIÓN",
            "FECHA DE NACIMIENTO", "GENERO", "GÉNERO"
        ]
    ):
        return False
    if re.fullmatch(r"[\d\s./-]+", valor):
        return False
    return sum(c.isalpha() for c in valor) >= 3


def candidato_lugar_valido(valor):
    valor = limpiar_candidato_campo(valor, "lugar")
    if not valor or len(valor) < 3 or len(valor) > 100:
        return False
    if re.search(r"https?://|www\.|@", valor, flags=re.IGNORECASE):
        return False
    if re.fullmatch(r"[\d\s:./-]+", valor):
        return False
    norm = normalizar_etiqueta(valor)
    if norm in {normalizar_etiqueta(x) for x in _RUIDO_LUGAR}:
        return False
    if any(
        ruido in norm
        for ruido in [
            "PAGINA", "CERTIFICADO", "LOGOTIPO", "FIRMA", "CONSENTIMIENTO",
            "HORA DE IMPRESION", "HORA DE IMPRESIÓN", "A.M", "P.M"
        ]
    ):
        return False
    return sum(c.isalpha() for c in valor) >= 3


def elegir_mejor_candidato(candidatos, tipo):
    validadores = {
        "nombre": candidato_nombre_valido,
        "cargo": candidato_cargo_valido,
        "lugar": candidato_lugar_valido,
    }
    validador = validadores[tipo]
    mejores = []

    for puntaje, valor, origen in candidatos:
        limpio = limpiar_candidato_campo(valor, tipo)
        if not validador(limpio):
            continue

        tokens = limpio.split()
        if tipo == "nombre":
            if 2 <= len(tokens) <= 5:
                puntaje += 8
            if limpio.upper() == limpio:
                puntaje += 3
        elif tipo == "cargo":
            if 1 <= len(tokens) <= 7:
                puntaje += 4
        elif tipo == "lugar":
            if 1 <= len(tokens) <= 5:
                puntaje += 4

        mejores.append((puntaje, -len(limpio), limpio, origen))

    if not mejores:
        return ""

    mejores.sort(reverse=True)
    valor = mejores[0][2]
    if tipo == "nombre":
        return valor.title()
    if tipo == "cargo":
        return corregir_ortografia_sst(valor).title()
    return valor.title()


def _coincide_etiqueta(columna, etiquetas):
    norm = normalizar_etiqueta(columna)
    return any(
        norm == normalizar_etiqueta(etiqueta)
        or normalizar_etiqueta(etiqueta) in norm
        for etiqueta in etiquetas
    )


def extraer_campo_por_etiquetas(lineas, etiquetas, tipo):
    """
    Obtiene candidatos en tres estructuras frecuentes:
    1. ETIQUETA: valor
    2. ETIQUETA    valor
    3. fila de encabezados + fila de valores.
    """
    candidatos = []
    etiquetas_ordenadas = sorted(etiquetas, key=len, reverse=True)
    patron_etiquetas = "|".join(re.escape(e) for e in etiquetas_ordenadas)

    for idx, linea in enumerate(lineas):
        if not linea or not linea.strip():
            continue

        # Valor en la misma línea después de dos puntos, guion o separación de tabla.
        m_inline = re.search(
            rf"(?:{patron_etiquetas})(?:\s*[:=]\s*|\s+-\s+|\s{{2,}}|\|)\s*(.+)$",
            linea,
            flags=re.IGNORECASE,
        )
        if m_inline:
            candidatos.append((115, m_inline.group(1), "etiqueta en línea"))

        columnas_header = dividir_columnas_estructuradas(linea)
        indices = [
            pos
            for pos, col in enumerate(columnas_header)
            if _coincide_etiqueta(col, etiquetas)
        ]

        if indices:
            for offset in range(1, 5):
                if idx + offset >= len(lineas):
                    break
                siguiente = lineas[idx + offset].strip()
                if not siguiente:
                    continue

                columnas_valor = dividir_columnas_estructuradas(siguiente)
                for pos in indices:
                    if pos < len(columnas_valor):
                        candidatos.append(
                            (105 - offset, columnas_valor[pos], "tabla encabezado/valor")
                        )

                # Si la etiqueta ocupa toda la línea, la línea siguiente completa
                # suele ser el dato, aunque el PDF haya perdido columnas.
                norm_linea = normalizar_etiqueta(linea)
                if any(
                    norm_linea == normalizar_etiqueta(etiqueta)
                    for etiqueta in etiquetas
                ):
                    candidatos.append((100 - offset, siguiente, "línea siguiente"))
                break

        # Etiqueta ubicada al inicio, pero sin separador reconocible.
        norm_linea = normalizar_etiqueta(linea)
        for etiqueta in etiquetas_ordenadas:
            norm_etiqueta = normalizar_etiqueta(etiqueta)
            if norm_linea.startswith(norm_etiqueta) and len(norm_linea) > len(norm_etiqueta):
                resto = linea[len(etiqueta):].strip(" |/-,_.:")
                if resto:
                    candidatos.append((92, resto, "etiqueta inicial"))
                break

    return elegir_mejor_candidato(candidatos, tipo)


def extraer_fecha_y_lugar_robusto(lineas, texto_completo):
    candidatos_lugar = []
    fechas = []

    patron_fecha_dmy = re.compile(
        r"\b(\d{1,2})\s*(?:[|/\-.]|\s)\s*(\d{1,2})\s*(?:[|/\-.]|\s)\s*(20\d{2})\b"
    )
    patron_fecha_ymd = re.compile(
        r"\b(20\d{2})\s*(?:[|/\-.]|\s)\s*(\d{1,2})\s*(?:[|/\-.]|\s)\s*(\d{1,2})\b"
    )

    # Tablas con encabezados DÍA | MES | AÑO | CIUDAD/MUNICIPIO.
    for idx, linea in enumerate(lineas):
        columnas = dividir_columnas_estructuradas(linea)
        columnas_norm = [normalizar_etiqueta(c) for c in columnas]

        ciudad_indices = [
            i for i, c in enumerate(columnas_norm)
            if any(k in c for k in ["CIUDAD", "MUNICIPIO", "LUGAR", "SEDE"])
        ]

        if ciudad_indices and any(
            k in " ".join(columnas_norm)
            for k in ["DIA", "MES", "ANO", "FECHA", "REALIZACION"]
        ):
            for offset in range(1, 4):
                if idx + offset >= len(lineas):
                    break
                valores = dividir_columnas_estructuradas(lineas[idx + offset])
                for pos in ciudad_indices:
                    if pos < len(valores):
                        candidatos_lugar.append(
                            (125 - offset, valores[pos], "tabla fecha/lugar")
                        )
                linea_valores = lineas[idx + offset]
                m_dmy = patron_fecha_dmy.search(linea_valores)
                m_ymd = patron_fecha_ymd.search(linea_valores)
                try:
                    if m_ymd:
                        fechas.append(
                            (
                                120 - offset,
                                datetime.date(
                                    int(m_ymd.group(1)),
                                    int(m_ymd.group(2)),
                                    int(m_ymd.group(3)),
                                ),
                            )
                        )
                    elif m_dmy:
                        fechas.append(
                            (
                                120 - offset,
                                datetime.date(
                                    int(m_dmy.group(3)),
                                    int(m_dmy.group(2)),
                                    int(m_dmy.group(1)),
                                ),
                            )
                        )
                except ValueError:
                    pass
                if valores:
                    break

        m_dmy = patron_fecha_dmy.search(linea)
        m_ymd = patron_fecha_ymd.search(linea)
        m_fecha = m_ymd or m_dmy
        if m_fecha:
            try:
                if m_ymd:
                    fecha_detectada = datetime.date(
                        int(m_ymd.group(1)),
                        int(m_ymd.group(2)),
                        int(m_ymd.group(3)),
                    )
                else:
                    fecha_detectada = datetime.date(
                        int(m_dmy.group(3)),
                        int(m_dmy.group(2)),
                        int(m_dmy.group(1)),
                    )
                fechas.append((90, fecha_detectada))
            except ValueError:
                pass

            antes = linea[:m_fecha.start()].strip(" |/-,_.:")
            despues = linea[m_fecha.end():].strip(" |/-,_.:")

            contexto = normalizar_etiqueta(linea)
            puntaje = 118 if any(
                k in contexto
                for k in ["REALIZACION", "CIUDAD", "MUNICIPIO", "LUGAR", "SEDE"]
            ) else 78

            if despues:
                candidatos_lugar.append((puntaje, despues, "después de fecha"))
            if antes and len(antes.split()) <= 7:
                # Evita usar el encabezado como ciudad.
                antes = re.sub(
                    r"(?i)\b(?:FECHA|CIUDAD|MUNICIPIO|LUGAR|REALIZACI[ÓO]N|DEL EXAMEN|DEL EXÁMEN)\b",
                    "",
                    antes,
                ).strip(" |/-,_.:")
                if antes:
                    candidatos_lugar.append((puntaje - 5, antes, "antes de fecha"))

    # Formatos como "Tunja, 15 de junio de 2026".
    meses = (
        "enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
        "septiembre|octubre|noviembre|diciembre"
    )
    patron_ciudad_fecha = re.compile(
        rf"\b([A-Za-zÁÉÍÓÚÜÑáéíóúüñ][A-Za-zÁÉÍÓÚÜÑáéíóúüñ .'-]{{2,45}}),?\s+"
        rf"(\d{{1,2}}\s+de\s+(?:{meses})\s+de\s+20\d{{2}})\b",
        flags=re.IGNORECASE,
    )
    for m in patron_ciudad_fecha.finditer(texto_completo):
        lugar = m.group(1).strip()
        # Toma solo el segmento final si quedó unido a texto anterior.
        lugar = re.split(r"\n|[:;]", lugar)[-1].strip()
        candidatos_lugar.append((112, lugar, "ciudad y fecha en letras"))
        try:
            fechas.append((112, intentar_parsear_fecha(m.group(2))))
        except Exception:
            pass

    etiquetas_lugar = [
        "FECHA Y CIUDAD DE REALIZACIÓN",
        "FECHA Y CIUDAD DE REALIZACION",
        "CIUDAD DE REALIZACIÓN",
        "CIUDAD DE REALIZACION",
        "LUGAR DE REALIZACIÓN",
        "LUGAR DE REALIZACION",
        "MUNICIPIO DE REALIZACIÓN",
        "MUNICIPIO DE REALIZACION",
        "LUGAR DEL EXAMEN",
        "LUGAR DONDE SE REALIZÓ EL EXAMEN",
        "LUGAR DONDE SE REALIZO EL EXAMEN",
        "LUGAR DONDE SE REALIZARON LOS EXÁMENES",
        "LUGAR DONDE SE REALIZARON LOS EXAMENES",
        "CIUDAD DEL EXAMEN",
        "MUNICIPIO DEL EXAMEN",
        "SEDE DE ATENCIÓN",
        "SEDE DE ATENCION",
        "CENTRO MÉDICO",
        "CENTRO MEDICO",
        "IPS PRESTADORA",
        "IPS QUE REALIZA EL EXAMEN",
        "CIUDAD",
        "MUNICIPIO",
        "SEDE",
    ]
    lugar_etiquetado = extraer_campo_por_etiquetas(
        lineas, etiquetas_lugar, "lugar"
    )
    if lugar_etiquetado:
        candidatos_lugar.append((130, lugar_etiquetado, "etiqueta explícita"))

    lugar = elegir_mejor_candidato(candidatos_lugar, "lugar")
    fecha = max(fechas, key=lambda item: item[0])[1] if fechas else datetime.date.today()
    return fecha, lugar


def extraer_identidad_cargo_lugar(texto):
    lineas = [line.rstrip() for line in texto.splitlines()]
    etiquetas_nombre = [
        "APELLIDOS Y NOMBRES DEL TRABAJADOR",
        "NOMBRES Y APELLIDOS DEL TRABAJADOR",
        "APELLIDOS Y NOMBRES",
        "NOMBRES Y APELLIDOS",
        "NOMBRE DEL TRABAJADOR",
        "NOMBRE TRABAJADOR",
        "NOMBRE COMPLETO",
        "NOMBRES COMPLETOS",
        "NOMBRE(S) Y APELLIDO(S)",
        "APELLIDO(S) Y NOMBRE(S)",
        "NOMBRE Y APELLIDOS",
        "NOMBRES DEL TRABAJADOR",
        "TRABAJADOR",
        "PACIENTE",
    ]
    etiquetas_cargo = [
        "CARGO ACTUAL DEL TRABAJADOR",
        "CARGO DEL TRABAJADOR",
        "CARGO ACTUAL",
        "OCUPACIÓN ACTUAL",
        "OCUPACION ACTUAL",
        "OCUPACIÓN",
        "OCUPACION",
        "PUESTO DE TRABAJO",
        "OCUPACIÓN DEL TRABAJADOR",
        "OCUPACION DEL TRABAJADOR",
        "CARGO U OCUPACIÓN",
        "CARGO U OCUPACION",
        "CARGO / OCUPACIÓN",
        "CARGO / OCUPACION",
        "PUESTO",
        "OFICIO",
        "LABOR",
        "CARGO",
    ]

    nombre = extraer_campo_por_etiquetas(
        lineas, etiquetas_nombre, "nombre"
    )
    cargo = extraer_campo_por_etiquetas(
        lineas, etiquetas_cargo, "cargo"
    )
    fecha, lugar = extraer_fecha_y_lugar_robusto(
        lineas, texto
    )

    return {
        "nombre": nombre,
        "cargo": cargo,
        "fecha": fecha,
        "lugar": lugar,
    }


def extraer_texto_pdf_robusto(pdf_raw_data):
    """
    Conserva el texto normal y agrega una representación de las tablas.
    Esto mejora especialmente nombre, cargo y ciudad sin alterar el resto
    del análisis existente.
    """
    lineas_salida = []
    vistos = set()

    def agregar(fragmento):
        if not fragmento:
            return
        for linea in str(fragmento).splitlines():
            # Se mantienen las separaciones amplias porque representan columnas
            # y permiten asociar correctamente encabezados con sus valores.
            linea = linea.replace("\t", "    ").strip()
            if not linea:
                continue
            clave = normalizar_etiqueta(linea)
            if clave and clave not in vistos:
                vistos.add(clave)
                lineas_salida.append(linea)

    with pdfplumber.open(io.BytesIO(pdf_raw_data)) as p_file:
        for numero_pagina, page in enumerate(p_file.pages, start=1):
            texto_layout = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
                layout=True,
            ) or ""
            texto_normal = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
                layout=False,
            ) or ""
            agregar(texto_layout)
            agregar(texto_normal)

            tablas = page.extract_tables() or []
            for tabla in tablas:
                for fila in tabla or []:
                    celdas = []
                    for celda in fila or []:
                        celda_limpia = re.sub(
                            r"\s+",
                            " ",
                            (celda or "").replace("\n", " "),
                        ).strip()
                        celdas.append(celda_limpia)
                    if any(celdas):
                        agregar(" | ".join(celdas))

    return "\n".join(lineas_salida)


# --- ANALIZADOR INTELIGENTE MULTILÍNEA GENERALIZADO ---
def analizar_pdf_inteligente(texto):
    datos = {
        "nombre": "", "cargo": "", "tipo_examen": "PERIODICO",
        "examenes_lista": [], "recomendaciones_lista": [], "vigilancia_lista": [],
        "observaciones": "", "remisiones": "No", "consecutivo": "",
        "vigilancia_programa": "NINGUNO",
        "lugar": "Tunja",
        "fecha": datetime.date.today()
    }
    if not texto:
        return datos

    lineas_raw = texto.split("\n")

    # Extracción prioritaria y robusta de los tres campos que suelen venir
    # dentro de tablas: trabajador, cargo y lugar de realización.
    identificacion = extraer_identidad_cargo_lugar(texto)

    if identificacion.get("nombre"):
        datos["nombre"] = identificacion["nombre"]

    if identificacion.get("cargo"):
        datos["cargo"] = identificacion["cargo"]

    if identificacion.get("lugar"):
        datos["lugar"] = identificacion["lugar"]

    if identificacion.get("fecha"):
        datos["fecha"] = identificacion["fecha"]

    EXAMS_MAP = {
        "AUDIOMETRIA DE TONOS": "Audiometría", "AUDIOMETRIA": "Audiometría",
        "ESPIROMETRIA": "Espirometría", "ESPIROMETRÍA": "Espirometría",
        "OPTOMETRIA": "Optometría", "OPTOMETRÍA": "Optometría",
        "VISIOMETRIA": "Visiometría", "VISIOMETRÍA": "Visiometría",
        "EXAMEN MEDICO OCUPACIONAL": "Examen Clínico Ocupacional",
        "EXAMEN MEDICO": "Examen Clínico Ocupacional",
        "EXAMEN OCUPACIONAL ENFASIS OSTEOMUSCULAR": "Énfasis Osteomuscular",
        "ENFASIS OSTEOMUSCULAR": "Énfasis Osteomuscular", "ÉNFASIS OSTEOMUSCULAR": "Énfasis Osteomuscular",
        "ELECTROCARDIOGRAMA DE RITMO": "Electrocardiograma", "ELECTROCARDIOGRAMA": "Electrocardiograma", 
        "FROTIS": "Frotis",
        "CUADRO HEMATICO": "Cuadro Hemático", "CUADRO HEMÁTICO": "Cuadro Hemático",
        "COLESTEROL": "Colesterol",
        "TRIGLICERIDOS": "Triglicéridos", "TRIGLICÉRIDOS": "Triglicéridos",
        "PARCIAL DE ORINA": "Parcial de Orina",
        "VSH": "VSH", "PCR": "PCR"
    }

    examenes_detectados = []
    recoms_raw_dict = {}
    current_exam = None
    in_exams_section = True
    formato_grilla_detectado = False
    recoms_grilla_acumuladas = []

    for idx_l, linea in enumerate(lineas_raw):
        linea_limpia = limpiar_linea_ruido_lateral(linea)
        linea_upper = linea_limpia.upper().strip()
        
        if "EL CONCEPTO DE APTITUD SE DEFINIÓ A PARTIR DE LOS SIGUIENTES EXÁMENES PRACTICADOS" in linea_upper:
            for offset in range(1, 4):
                if idx_l + offset < len(lineas_raw):
                    l_sig = lineas_raw[idx_l + offset].upper()
                    for k_ex, v_ex in EXAMS_MAP.items():
                        if k_ex in l_sig and v_ex not in examenes_detectados:
                            examenes_detectados.append(v_ex)
            continue
            
        if any(h in linea_upper for h in ["RECOMENDACIONES MÉDICAS", "RECOMENDACIONES OCUPACIONALES", "HABITOS Y ESTILO DE VIDA SALUDABLES"]):
            formato_grilla_detectado = True
            continue
            
        if formato_grilla_detectado:
            if any(stop in linea_upper for stop in ["OTRAS OBSERVACIONES", "REMISIONES:", "ATENTAMENTE"]):
                formato_grilla_detectado = False
            else:
                columnas = [col.strip(" |/-,_.") for col in re.split(r'\s{2,}|\|', linea_limpia) if col.strip()]
                for col in columnas:
                    if not es_vacio_o_estado(col):
                        rec_fmt = a_caso_oracion(col)
                        if rec_fmt and rec_fmt not in recoms_grilla_acumuladas:
                            recoms_grilla_acumuladas.append(rec_fmt)
                continue

        if any(stop in linea_upper for stop in ["OBSERVACIONES:", "OBSERVACION:", "REMISIONES:", "SISTEMA DE VIGILANCIA"]):
            in_exams_section = False
            if current_exam:
                recoms_raw_dict[current_exam] = recoms_raw_dict.get(current_exam, "")
                current_exam = None

        matched_key = None
        for key in sorted(EXAMS_MAP.keys(), key=len, reverse=True):
            if key in linea_upper and linea_upper.find(key) < 15:
                matched_key = key
                break
        
        if matched_key:
            in_exams_section = True
            current_exam = EXAMS_MAP[matched_key]
            if current_exam not in examenes_detectados:
                examenes_detectados.append(current_exam)
            idx = linea_upper.find(matched_key) + len(matched_key)
            recoms_raw_dict[current_exam] = linea_limpia[idx:].strip(" :-,_/")
        else:
            if in_exams_section and current_exam and linea_limpia.strip():
                if not (linea_limpia.isupper() and len(linea_limpia) > 10):
                    recoms_raw_dict[current_exam] = recoms_raw_dict.get(current_exam, "") + " " + linea_limpia.strip()

    recoms_por_examen = []
    pve_detectados = set()

    if recoms_grilla_acumuladas:
        for rec in recoms_grilla_acumuladas:
            recoms_por_examen.append(rec)
            rec_up = rec.upper()
            if any(re.search(patron, rec_up) for patron in [r'\bAUDITIV', r'\bRUIDO', r'\bOIDO', r'\bOÍDO', r'\bAUDIO']): pve_detectados.add("Conservación Auditiva")
            if any(re.search(patron, rec_up) for patron in [r'\bPOSTURAL', r'\bLUMBAR', r'\bOSTEOMUSCULAR', r'\bERGONOMIC', r'\bESPALDA', r'\bCARGA']): pve_detectados.add("Prevención Osteomuscular (DME)")
            if any(re.search(patron, rec_up) for patron in [r'\bVISUAL', r'\bGAFAS', r'\bVISION', r'\bVISIÓN', r'\bLENTE', r'\bOPTOMETR', r'\bRX\b']): pve_detectados.add("Conservación Visual")
            if any(re.search(patron, rec_up) for patron in [r'\bRESPIRATORI', r'\bESPIROMETR', r'\bPOLVO', r'\bHUMO']): pve_detectados.add("Conservación Respiratoria")
    else:
        for exam in examenes_detectados:
            rec_part = recoms_raw_dict.get(exam, "").strip()
            rec_part = re.sub(r'\s+', ' ', rec_part)
            rec_part = limpiar_ruido_columnas_final(rec_part)
            
            if not es_vacio_o_estado(rec_part):
                parts = re.split(r'//|;|\b\d+\.|\b\d+\-', rec_part)
                valid_parts = []
                for p in parts:
                    p_clean = p.strip(" .-_/()[]")
                    if not es_vacio_o_estado(p_clean):
                        valid_parts.append(a_caso_oracion(p_clean))
                        p_upper = p_clean.upper()
                        if any(re.search(patron, p_upper) for patron in [r'\bAUDITIV', r'\bRUIDO', r'\bOIDO', r'\bOÍDO', r'\bAUDIO']): pve_detectados.add("Conservación Auditiva")
                        if any(re.search(patron, p_upper) for patron in [r'\bPOSTURAL', r'\bLUMBAR', r'\bOSTEOMUSCULAR', r'\bERGONOMIC', r'\bESPALDA', r'\bCARGA']): pve_detectados.add("Prevención Osteomuscular (DME)")
                        if any(re.search(patron, p_upper) for patron in [r'\bVISUAL', r'\bGAFAS', r'\bVISION', r'\bVISIÓN', r'\bLENTE', r'\bOPTOMETR', r'\bRX\b']): pve_detectados.add("Conservación Visual")
                        if any(re.search(patron, p_upper) for patron in [r'\bRESPIRATORI', r'\bESPIROMETR', r'\bPOLVO', r'\bHUMO']): pve_detectados.add("Conservación Respiratoria")
                
                if valid_parts:
                    recoms_por_examen.append(f"{exam}: {' - '.join(valid_parts)}")

    # CLAVE UNIFICADA EN ESPAÑOL DEFINITIVA
    datos["examenes_lista"] = examenes_detectados
    datos["recomendaciones_lista"] = recoms_por_examen
    datos["vigilancia_lista"] = list(pve_detectados)

    # --- CORRECCIÓN INTEGRAL: EXTRACCIÓN ACOTA DE PROGRAMAS PVE ---
    programas_encontrados = []
    sve_clinical_keywords = {
        "VISUAL": "Conservación Visual", "AUDITIV": "Conservación Auditiva", 
        "RUIDO": "Conservación Auditiva", "OIDO": "Conservación Auditiva", "OÍDO": "Conservación Auditiva",
        "AUDIO": "Conservación Auditiva", "OSTEOMUSCULAR": "Prevención Osteomuscular (DME)",
        "POSTURAL": "Prevención Osteomuscular (DME)", "LUMBAR": "Prevención Osteomuscular (DME)",
        "ERGONOMIC": "Prevención Osteomuscular (DME)", "ESPALDA": "Prevención Osteomuscular (DME)",
        "DME": "Prevención Osteomuscular (DME)", "RESPIRATORI": "Conservación Respiratoria",
        "ESPIROMETR": "Conservación Respiratoria", "POLVO": "Conservación Respiratoria",
        "HUMO": "Conservación Respiratoria", "CARDIOVASCULAR": "Riesgo Cardiovascular"
    }

    # Búsqueda vecinal: Solo lee las 2 líneas inmediatamente inferiores a la cabecera PVE
    for idx, line in enumerate(lineas_raw):
        l_up = line.upper()
        if "INGRESAR AL PROGRAMA DE VIGILANCIA" in l_up or "PROGRAMA DE VIGILANCIA EPIDEMIOL" in l_up:
            for offset in [0, 1, 2]:
                if idx + offset < len(lineas_raw):
                    text_target = lineas_raw[idx + offset].upper()
                    # Freno de emergencia si cruza a la declaración de firmas/consentimientos legales
                    if offset > 0 and any(stop in text_target for stop in ["REMISIONES:", "OBSERVACIONES:", "ATENTAMENTE", "CONSENTIMIENTO"]):
                        break
                    for kw, prog_name in sve_clinical_keywords.items():
                        if kw in text_target and prog_name not in programas_encontrados:
                            programas_encontrados.append(prog_name)
            break

    if not programas_encontrados and pve_detectados:
        programas_encontrados = [p.upper() for p in pve_detectados]
        
    datos["vigilancia_programa"] = ", ".join(programas_encontrados) if programas_encontrados else "NINGUNO"

    def extraer_seccion_limpia(texto_completo, palabras_inicio, palabras_fin):
        seccion = []
        dentro = False
        for l in texto_completo.split('\n'):
            l_limpia = limpiar_linea_ruido_lateral(l)
            l_upper = l_limpia.upper().strip()
            if not dentro:
                if any(h in l_upper for h in palabras_inicio):
                    dentro = True
                    for h in palabras_inicio:
                        if h in l_upper:
                            resto = l_limpia[l_upper.find(h) + len(h):].strip(" :-,_")
                            if resto: seccion.append(resto)
                            break
            else:
                if any(h in l_upper for h in palabras_fin): break
                seccion.append(l_limpia)
        return "\n".join([s for s in seccion if s]).strip()

    obs_fmt_nuevo = ""
    m_obs_nuevo = re.search(r'OTRAS OBSERVACIONES Y RECOMENDACIONES\s*\n\s*([^\n]+)', texto, re.IGNORECASE)
    if m_obs_nuevo: obs_fmt_nuevo = m_obs_nuevo.group(1).strip()
        
    if obs_fmt_nuevo and not es_vacio_o_estado(obs_fmt_nuevo):
        datos["observaciones"] = a_caso_oracion(obs_fmt_nuevo)
    else:
        datos["observaciones"] = a_caso_oracion(extraer_seccion_limpia(
            texto, ["OBSERVACIONES:"], ["RECOMENDACIONES", "REMISIONES", "INGRESAR AL PROGRAMA", "PROGRAMA DE VIGILANCIA"]
        ))
    
    rem_raw = extraer_seccion_limpia(texto, ["INFORMACION DE REMISIONES", "INFORMACIÓN DE REMISIONES"], ["CONSENTIMIENTO", "AUTORIZO"])
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

def replace_placeholder_in_paragraph_runs(paragraph, placeholder, value):
    if placeholder not in paragraph.text: return False
    replaced = False
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            replaced = True
            
    if not replaced:
        font_name = "Arial"; font_size = Pt(11); bold = False; italic = False; color = None
        if paragraph.runs:
            for r in paragraph.runs:
                if r.text.strip():
                    font_name = r.font.name or font_name
                    font_size = r.font.size or font_size
                    bold = r.bold if r.bold is not None else bold
                    italic = r.italic if r.italic is not None else italic
                    color = r.font.color.rgb if r.font.color else color
                    break
        full_text = paragraph.text.replace(placeholder, value)
        paragraph.text = ""
        new_run = paragraph.add_run(full_text)
        new_run.font.name = font_name; new_run.font.size = font_size; new_run.bold = bold; new_run.italic = italic
        if color: new_run.font.color.rgb = color
    return True

def insert_bullets_in_placeholder(parent_container, paragraph, items_list):
    if not paragraph.runs:
        font_name = "Arial"; font_size = Pt(11); bold = False; color = None
    else:
        first_run = paragraph.runs[0]
        font_name = first_run.font.name or "Arial"
        font_size = first_run.font.size or Pt(11)
        bold = first_run.bold
        color = first_run.font.color.rgb if first_run.font.color else None

    paragraph.text = ""
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.space_before = Pt(0)
    
    if not items_list:
        run = paragraph.add_run("Ninguno.")
        run.font.name = font_name; run.font.size = font_size; run.bold = bold
        if color: run.font.color.rgb = color
        return

    run = paragraph.add_run("• " + items_list[0])
    run.font.name = font_name; run.font.size = font_size; run.bold = bold
    if color: run.font.color.rgb = color

    current_p = paragraph
    for item in items_list[1:]:
        new_p_element = OxmlElement('w:p')
        current_p._p.addnext(new_p_element)
        new_para = Paragraph(new_p_element, parent_container)
        new_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
        new_para.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
        new_para.paragraph_format.space_after = Pt(2)
        new_para.paragraph_format.space_before = Pt(0)
        new_para.paragraph_format.left_indent = paragraph.paragraph_format.left_indent or Inches(0.25)
        run_new = new_para.add_run("• " + item)
        run_new.font.name = font_name; run_new.font.size = font_size; run_new.bold = bold
        if color: run_new.font.color.rgb = color
        current_p = new_para

def insert_recommendations_in_placeholder(parent_container, paragraph, recom_list):
    combined_items = list(recom_list)
    if paragraph.runs:
        font_name = paragraph.runs[0].font.name or "Arial"
        font_size = paragraph.runs[0].font.size or Pt(11)
        color = paragraph.runs[0].font.color.rgb if paragraph.runs[0].font.color else None
    else:
        font_name = "Arial"; font_size = Pt(11); color = None

    paragraph.text = ""
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.space_before = Pt(0)
    run_lbl = paragraph.add_run("Recomendaciones: ")
    run_lbl.bold = True; run_lbl.font.name = font_name; run_lbl.font.size = font_size
    if color: run_lbl.font.color.rgb = color

    if not combined_items:
        run_none = paragraph.add_run("Ninguna.")
        run_none.font.name = font_name; run_none.font.size = font_size
        if color: run_none.font.color.rgb = color
        return

    run_first = paragraph.add_run("• " + combined_items[0])
    run_first.font.name = font_name; run_first.font.size = font_size
    if color: run_first.font.color.rgb = color

    current_p = paragraph
    for item in combined_items[1:]:
        new_p_element = OxmlElement('w:p')
        current_p._p.addnext(new_p_element)
        new_para = Paragraph(new_p_element, parent_container)
        new_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
        new_para.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
        new_para.paragraph_format.space_after = Pt(2)
        new_para.paragraph_format.space_before = Pt(0)
        new_para.paragraph_format.left_indent = Inches(0.25)
        run_new = new_para.add_run("• " + item)
        run_new.font.name = font_name; run_new.font.size = font_size
        if color: run_new.font.color.rgb = color
        current_p = new_para

def replace_label_placeholder(paragraph, label_text, placeholder, value):
    if placeholder not in paragraph.text: return False
    if paragraph.runs:
        font_name = paragraph.runs[0].font.name or "Arial"
        font_size = paragraph.runs[0].font.size or Pt(11)
        color = paragraph.runs[0].font.color.rgb if paragraph.runs[0].font.color else None
    else:
        font_name = "Arial"; font_size = Pt(11); color = None
        
    paragraph.text = ""
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.space_before = Pt(0)
    run_lbl = paragraph.add_run(label_text)
    run_lbl.bold = True; run_lbl.font.name = font_name; run_lbl.font.size = font_size
    if color: run_lbl.font.color.rgb = color
    
    val_clean = value.strip() if value else "Ninguna."
    if es_vacio_o_negativo(val_clean): val_clean = "Ninguna."
        
    run_val = paragraph.add_run(val_clean)
    run_val.font.name = font_name; run_val.font.size = font_size
    if color: run_val.font.color.rgb = color
    return True

def obtener_siguiente_consecutivo_local():
    val = obtener_config("ultimo_consecutivo_local")
    return int(val) + 1 if val else 1

def incrementar_consecutivo_local():
    next_num = obtener_siguiente_consecutivo_local()
    guardar_config("ultimo_consecutivo_local", str(next_num))
    return f"SST-2026-{next_num}"

# --- CONSTRUCTOR DE DOCUMENTO ÚNICO INTELIGENTE ---
def generar_word_unico(datos_trabajador, lugar, fecha, template_uploaded, firma_file):
    if template_uploaded: doc_word = Document(template_uploaded)
    elif os.path.exists("FORMATO RECOMENDACIONES MEDICAS BOT.docx"): doc_word = Document("FORMATO RECOMENDACIONES MEDICAS BOT.docx")
    else: doc_word = Document()
    
    consecutivo_final = datos_trabajador.get("consecutivo", "")
    if not consecutivo_final:
        g_url = obtener_config("google_sheets_url")
        if g_url:
            try:
                r = requests.get(g_url, params={
                    "name": datos_trabajador["nombre"], "cargo": datos_trabajador["cargo"], 
                    "examen": datos_trabajador["tipo_examen"], "fecha": fecha.strftime("%Y-%m-%d")
                }, timeout=10)
                consecutivo_final = r.json().get("consecutive") if r.json().get("status") == "success" else incrementar_consecutivo_local()
            except: consecutivo_final = incrementar_consecutivo_local()
        else: consecutivo_final = incrementar_consecutivo_local()
        datos_trabajador["consecutivo"] = consecutivo_final

    simple_replacements = {
        "{{NUMERO DE CONSECUTIVO}}": consecutivo_final, 
        "{{TIPO DE EXAMEN}}": datos_trabajador["tipo_examen"].upper(),
        "{{LUGAR}}": lugar, "{{FECHA HOY}}": fecha.strftime("%d de %B de %Y"),
        "{{NOMBRE DE LA PERSONA}}": datos_trabajador["nombre"].upper(), 
        "{{CARGO DE LA PERSONA}}": datos_trabajador["cargo"].upper(),
        "{{Programa de vigilancia epidemiológica}}": datos_trabajador.get("vigilancia_programa", "NINGUNO").upper()
    }

    def procesar_parrafo(p, container):
        if "{{LISTA DE EXAMENES REALIZADOS}}" in p.text:
            insert_bullets_in_placeholder(container, p, datos_trabajador["examenes_lista"])
            return True
        if "{{Recomendaciones médicas}}" in p.text:
            insert_recommendations_in_placeholder(container, p, datos_trabajador["recomendaciones_lista"])
            return True
        if "{{Observaciones}}".lower() in p.text.lower():
            replace_label_placeholder(p, "Observaciones: ", p.text, datos_trabajador["observaciones"])
            return True
        if "{{Remisiones}}".lower() in p.text.lower():
            replace_label_placeholder(p, "Remisiones: ", p.text, datos_trabajador["remisiones"])
            return True
            
        for k, v in simple_replacements.items():
            if k in p.text: replace_placeholder_in_paragraph_runs(p, k, v)
        aplicar_negrita_dinamica_cuerpo(p, datos_trabajador["tipo_examen"])

    for p in list(doc_word.paragraphs): procesar_parrafo(p, doc_word)

    for table in doc_word.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in list(cell.paragraphs): procesar_parrafo(p, cell)
                
                idx_victor = -1
                for idx, p in enumerate(cell.paragraphs):
                    if "VÍCTOR ALONSO MORENO CASAS" in p.text:
                        idx_victor = idx
                        break
                if idx_victor != -1 and firma_file:
                    target_idx = max(0, idx_victor - 1)
                    p_firma = cell.paragraphs[target_idx]
                    p_firma.text = ""
                    p_firma.add_run().add_picture(firma_file, width=Inches(1.6))

    b_io = io.BytesIO()
    doc_word.save(b_io)
    return b_io.getvalue(), consecutivo_final

# --- COMPILADOR DE WORD A PDF ---
def convertir_docx_a_pdf(docx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
        temp_docx.write(docx_bytes)
        temp_docx_path = temp_docx.name
    pdf_path = temp_docx_path.replace(".docx", ".pdf")
    
    try:
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", 
            "--outdir", os.path.dirname(temp_docx_path), temp_docx_path
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f: pdf_bytes = f.read()
            os.unlink(temp_docx_path); os.unlink(pdf_path)
            return pdf_bytes, True
    except: pass

    try:
        from docx2pdf import convert
        convert(temp_docx_path, pdf_path)
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f: pdf_bytes = f.read()
            os.unlink(temp_docx_path); os.unlink(pdf_path)
            return pdf_bytes, True
    except: pass

    if os.path.exists(temp_docx_path): os.unlink(temp_docx_path)
    return None, False

def generar_html_vista(datos, consecutivo_num, lugar, fecha):
    return f"""
    <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; line-height: 1.5; background: white; border: 1px solid #ccc; max-width: 800px; margin: auto;">
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
        <p><strong>Programa de Vigilancia:</strong> {datos.get('vigilancia_programa', 'NINGUNO')}</p>
        <p><strong>observaciones:</strong> {datos['observaciones']}</p>
        <p><strong>remisiones:</strong> {datos['remisiones']}</p><br>
        <p>Atentamente,</p><br>
        <p><strong>VÍCTOR ALONSO MORENO CASAS</strong><br>Coordinador SST</p>
    </div>
    """

# --- VISTA PRINCIPAL STREAMLIT ---
st.markdown("<div class='header-banner'><h1>🩺 Portal de Control SST - JER S.A.</h1><p>Generación de Comunicaciones con Negrita Dinámica, Google Sheets y Firma Digital</p></div>", unsafe_allow_html=True)

st.sidebar.markdown(f"<h3 style='color:#60a5fa;'>👤 Perfil Activo</h3>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='metric-card'><strong>Usuario:</strong> {st.session_state.username}</div>", unsafe_allow_html=True)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.documentos = {}
    st.session_state.pdfs_raw_bytes = {}
    st.session_state.textos_raw = {}
    st.session_state.export_bytes = None
    st.session_state.zip_bytes = None
    st.session_state.processed_doc = None
    st.session_state.prev_colaborador = None
    st.session_state.document_count = 0
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

col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.markdown("<h3 style='color:#60a5fa;'>📂 1. Carga de Documentos PDF</h3>", unsafe_allow_html=True)
    pdfs_subidos = st.file_uploader("Carga los archivos PDF:", type="pdf", accept_multiple_files=True)
    
    if pdfs_subidos:
        if len(pdfs_subidos) != st.session_state.document_count:
            st.session_state.documentos = {}
            st.session_state.pdfs_raw_bytes = {}
            st.session_state.zip_bytes = None
            st.session_state.document_count = len(pdfs_subidos)
            
        for pdf in pdfs_subidos:
            if pdf.name not in st.session_state.documentos:
                pdf_raw_data = pdf.read()
                st.session_state.pdfs_raw_bytes[pdf.name] = pdf_raw_data
                texto_raw = extraer_texto_pdf_robusto(pdf_raw_data)
                st.session_state.textos_raw[pdf.name] = texto_raw
                st.session_state.documentos[pdf.name] = analizar_pdf_inteligente(texto_raw)
        
        st.markdown(f"""
            <div style='display:flex; gap:10px; margin-top:15px;'>
                <div class='metric-card' style='flex:1;'><strong>PDFs Leídos</strong><br><span style='font-size:20px; font-weight:700; color:#60a5fa;'>{len(st.session_state.documentos)}</span></div>
                <div class='metric-card' style='flex:1;'><strong>Nube Sheets</strong><br><span style='font-size:14px; font-weight:600; color:#4ade80;'>{'Conectado' if g_url_guardada else 'Modo Local'}</span></div>
            </div>
        """, unsafe_allow_html=True)
        
        archivo_seleccionado = st.selectbox("🎯 Selecciona Colaborador:", list(st.session_state.documentos.keys()))
        
        if archivo_seleccionado and archivo_seleccionado in st.session_state.pdfs_raw_bytes:
            st.markdown("---")
            st.markdown("<h4 style='color:#60a5fa;'>📄 Soporte Visual de Comparación</h4>", unsafe_allow_html=True)
            
            bytes_originales = st.session_state.pdfs_raw_bytes[archivo_seleccionado]
            st.download_button(
                label="📥 Descargar PDF de Origen (IPS)",
                data=bytes_originales,
                file_name=f"ORIGINAL_{archivo_seleccionado}",
                mime="application/pdf",
                key="btn_download_original_source"
            )
            
            with st.expander("👁️ Ver / Ocultar PDF de Origen Subido", expanded=True):
                with pdfplumber.open(io.BytesIO(bytes_originales)) as preview_pdf:
                    for i, page in enumerate(preview_pdf.pages):
                        img_pil = page.to_image(resolution=130).original
                        st.image(img_pil, caption=f"Página {i+1} - {archivo_seleccionado}", use_container_width=True)
    else:
        archivo_seleccionado = None
        st.session_state.documentos = {}
        st.session_state.pdfs_raw_bytes = {}
        st.session_state.zip_bytes = None
        st.session_state.document_count = 0

with col_der:
    st.markdown("<h3 style='color:#60a5fa;'>📋 2. Editor del Trabajador Seleccionado</h3>", unsafe_allow_html=True)
    if archivo_seleccionado:
        if archivo_seleccionado != st.session_state.prev_colaborador:
            st.session_state.processed_doc = None
            st.session_state.prev_colaborador = archivo_seleccionado
            
        doc_actual = st.session_state.documentos[archivo_seleccionado]
        
        col_f1, col_f2 = st.columns(2)
        with col_f1: lugar = st.text_input("Lugar:", value=doc_actual.get("lugar", "Tunja"))
        with col_f2: fecha = st.date_input("Fecha:", value=doc_actual.get("fecha", datetime.date.today()))
        
        tipo_examen = st.text_input("Tipo de Examen:", value=doc_actual["tipo_examen"].upper())
        
        col_p1, col_p2 = st.columns(2)
        with col_p1: nombre_persona = st.text_input("Trabajador:", value=doc_actual["nombre"])
        with col_p2: cargo_persona = st.text_input("Cargo:", value=doc_actual["cargo"])
        
        examenes_realizados = st.text_area("Exámenes Realizados:", value="\n".join(doc_actual["examenes_lista"]))
        recom_medicas = st.text_area("Recomendaciones por Examen:", value="\n".join(doc_actual["recomendaciones_lista"]), height=130)
        
        programa_vigilancia = st.text_input("Programa de Vigilancia Epidemiológica (PVE):", value=doc_actual.get("vigilancia_programa", "NINGUNO"))
        
        observaciones = st.text_area("Observaciones:", value=doc_actual["observaciones"])
        remisiones = st.text_input("Remisiones (Escribe 'No' para marcarlo negativo):", value=doc_actual["remisiones"])

        valores_actualizados = {
            "nombre": nombre_persona, "cargo": cargo_persona, "tipo_examen": tipo_examen,
            "examenes_lista": [l.strip() for l in examenes_realizados.split('\n') if l.strip()],
            "recomendaciones_lista": [l.strip() for l in recom_medicas.split('\n') if l.strip()],
            "observaciones": observaciones.strip(), "remisiones": remisiones.strip(),
            "vigilancia_programa": programa_vigilancia.strip(),
            "lugar": lugar, "fecha": fecha
        }
        
        for clave, valor in valores_actualizados.items():
            if doc_actual.get(clave) != valor:
                st.session_state.processed_doc = None
                doc_actual[clave] = valor

        st.markdown("---")
        formato_salida = st.radio("⚡ Elige formato de generación:", ["Microsoft Word (.docx)", "Documento PDF Oficial (.pdf)", "Impresión de Respaldo Web (HTML)"], horizontal=True)
        
        col_act1, col_gen2 = st.columns(2)
        
        with col_act1:
            if st.button("✨ Procesar este Colaborador"):
                with st.spinner("Procesando y registrando documento..."):
                    bytes_word, consec_num = generar_word_unico(doc_actual, lugar, fecha, template_uploaded, firma_file)
                    
                    if "Word" in formato_salida:
                        st.session_state.processed_doc = {
                            "bytes": bytes_word, "consec_num": consec_num,
                            "filename": f"Informe_{nombre_persona.replace(' ','_')}.docx",
                            "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        }
                    elif "PDF" in formato_salida:
                        bytes_pdf, exito = convertir_docx_a_pdf(bytes_word)
                        if exito:
                            st.session_state.processed_doc = {
                                "bytes": bytes_pdf, "consec_num": consec_num,
                                "filename": f"Informe_{nombre_persona.replace(' ','_')}.pdf", "mime": "application/pdf"
                            }
                        else: st.error("⚠️ No se pudo compilar el PDF de manera directa para coincidir con el Word.")
                    else:
                        html_out = generar_html_vista(doc_actual, consec_num, lugar, fecha)
                        st.session_state.processed_doc = {
                            "bytes": html_out.encode('utf-8'), "consec_num": consec_num,
                            "filename": f"Informe_{nombre_persona.replace(' ','_')}.html", "mime": "text/html"
                        }
                st.rerun()
            
            if st.session_state.processed_doc is not None:
                doc_info = st.session_state.processed_doc
                st.success(f"🟢 Guardado con éxito en base (Consecutivo: {doc_info['consec_num']})")
                st.download_button(
                    label=f"📥 Descargar archivo generado ({doc_info['filename'].split('.')[-1].upper()})",
                    data=doc_info["bytes"], file_name=doc_info["filename"], mime=doc_info["mime"]
                )
                        
        with col_gen2:
            if len(st.session_state.documentos) > 1:
                if st.button("📦 Procesar TODOS los Colaboradores (ZIP)"):
                    with st.spinner("Compilando lote masivo..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for filename, datos_trab in st.session_state.documentos.items():
                                bytes_word, consec_num = generar_word_unico(datos_trab, lugar, fecha, template_uploaded, firma_file)
                                if "Word" in formato_salida:
                                    zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.docx", bytes_word)
                                elif "PDF" in formato_salida:
                                    bytes_pdf, exito = convertir_docx_a_pdf(bytes_word)
                                    if exito: zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.pdf", bytes_pdf)
                                    else: zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.docx", bytes_word)
                                else:
                                    html_out = generar_html_vista(datos_trab, consec_num, lugar, fecha)
                                    zf.writestr(f"Recomendaciones_{datos_trab['nombre'].replace(' ', '_')}.html", html_out.encode('utf-8'))
                        zip_buffer.seek(0)
                        st.session_state.zip_bytes = zip_buffer.getvalue()
                        st.success("🎉 ZIP de lote masivo compilado con éxito.")
                        st.rerun()
                        
                if st.session_state.zip_bytes is not None:
                    st.download_button(
                        label="📥 Descargar ZIP Masivo", data=st.session_state.zip_bytes, 
                        file_name=f"Lote_SST_JER_SA_{fecha.strftime('%Y%m%d')}.zip", mime="application/zip"
                    )
    else:
        st.markdown("<div style='text-align:center; padding: 40px; color:#64748b;'><h3>👋 Tablero Listo</h3><p>Por favor, arrastra tus archivos PDF en la sección izquierda para activar el procesamiento automático.</p></div>", unsafe_allow_html=True)
