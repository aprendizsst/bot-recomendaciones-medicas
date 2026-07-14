import streamlit as st
import pdfplumber
from docx import Document
from docx.shared import Pt, Inches, RGBColor
import io

# Configuración de la página web
st.set_page_config(page_title="Creador de Informes Médicos", page_icon="📝", layout="wide")

st.title("📝 Generador Personalizado de Informes Médicos")
st.write("Sube tu PDF, edita el contenido extraído y personaliza el diseño de tu Word sin depender de APIs externas.")

# --- BARRA LATERAL DE DISEÑO Y FORMATO ---
st.sidebar.header("🎨 Configuración del Formato Word")

# 1. Selección de Fuente
tipo_fuente = st.sidebar.selectbox(
    "Fuente del documento:",
    ["Arial", "Calibri", "Times New Roman", "Georgia", "Verdana"]
)

# 2. Tamaños de Letra
col_t1, col_t2 = st.sidebar.columns(2)
with col_t1:
    tamano_titulo = st.sidebar.slider("Tamaño de Título:", 16, 28, 20)
with col_t2:
    tamano_cuerpo = st.sidebar.slider("Tamaño de Texto:", 10, 14, 11)

# 3. Colores Personalizados (Para títulos y subtítulos)
color_principal_hex = st.sidebar.color_picker("Color para Títulos:", "#1f4e79")

# 4. Configuración de Márgenes
margen_seleccionado = st.sidebar.radio(
    "Márgenes de página:",
    ["Normal (2.5 cm)", "Estrecho (1.27 cm)"]
)

# Función para convertir Color Hexadecimal a RGB para Word
def hex_a_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

# --- PROCESO DE DOCUMENTO ---
uploaded_file = st.file_uploader("Sube tu archivo PDF médico aquí", type="pdf")

if uploaded_file is not None:
    # Si el archivo cambia o se sube por primera vez, extraemos el texto
    if "texto_extraido" not in st.session_state or st.session_state.get("ultimo_archivo") != uploaded_file.name:
        with st.spinner("Leyendo el PDF..."):
            texto_pdf = ""
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    texto_pdf += page.extract_text() + "\n"
            st.session_state.texto_extraido = texto_pdf
            st.session_state.ultimo_archivo = uploaded_file.name

    # Área de edición interactiva
    st.subheader("✍️ Revisa, edita o limpia la información extraída:")
    st.write("Puedes escribir, borrar o formatear el texto aquí mismo. Usa `#` para títulos, `##` para subtítulos y `-` para viñetas.")
    
    texto_editable = st.text_area(
        "Contenido del documento:",
        value=st.session_state.texto_extraido,
        height=350
    )

    # Botón para generar Word
    if st.button("💾 Generar Word con este Diseño"):
        try:
            doc = Document()
            
            # --- APLICAR MÁRGENES ---
            valor_margen = 1.0 if "Normal" in margen_seleccionado else 0.5
            for section in doc.sections:
                section.top_margin = Inches(valor_margen)
                section.bottom_margin = Inches(valor_margen)
                section.left_margin = Inches(valor_margen)
                section.right_margin = Inches(valor_margen)

            # --- CONFIGURAR FUENTE BASE (Normal) ---
            estilo_normal = doc.styles['Normal']
            fuente_normal = estilo_normal.font
            fuente_normal.name = tipo_fuente
            fuente_normal.size = Pt(tamano_cuerpo)

            # Procesar el texto editado línea por línea
            color_rgb = hex_a_rgb(color_principal_hex)
            
            for linea in texto_editable.split('\n'):
                linea_limpia = linea.strip()
                if not linea_limpia:
                    continue
                
                # Detectar Título Principal (#)
                if linea_limpia.startswith('# '):
                    p = doc.add_heading(linea_limpia.replace('# ', '').strip(), level=1)
                    p.style.font.name = tipo_fuente
                    p.style.font.size = Pt(tamano_titulo)
                    p.style.font.color.rgb = color_rgb
                    
                # Detectar Subtítulo (##)
                elif linea_limpia.startswith('## '):
                    p = doc.add_heading(linea_limpia.replace('## ', '').strip(), level=2)
                    p.style.font.name = tipo_fuente
                    p.style.font.size = Pt(tamano_titulo - 4)
                    p.style.font.color.rgb = color_rgb
                    
                # Detectar Viñetas (- o *)
                elif linea_limpia.startswith('- ') or linea_limpia.startswith('* '):
                    texto_vineta = linea_limpia.replace('- ', '').replace('* ', '').strip()
                    p = doc.add_paragraph(texto_vineta, style='List Bullet')
                    p.style.font.name = tipo_fuente
                    p.style.font.size = Pt(tamano_cuerpo)
                    
                # Texto Normal
                else:
                    p = doc.add_paragraph(linea_limpia)
                    p.style.font.name = tipo_fuente
                    p.style.font.size = Pt(tamano_cuerpo)

            # Guardar en memoria para descarga
            b_io = io.BytesIO()
            doc.save(b_io)
            b_io.seek(0)

            st.success("🎉 ¡Documento de Word maquetado con éxito!")
            
            st.download_button(
                label="📥 Descargar archivo Word Personalizado",
                data=b_io,
                file_name="Informe_Medico_Formateado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Error al generar el Word: {e}")
