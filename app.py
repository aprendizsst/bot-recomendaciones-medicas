import streamlit as st
import pdfplumber
from docx import Document
import google.generativeai as genai
import io

# Configuración de la página web
st.set_page_config(page_title="Asistente Médico PDF", page_icon="🩺", layout="centered")

st.title("🩺 Extractor de Recomendaciones Médicas")
st.write("Sube tu PDF médico, la IA extraerá las recomendaciones y te entregará un archivo Word listo.")

# 1. Entrada segura para la API Key de Gemini
api_key = st.text_input("Introduce tu API Key de Google Gemini:", type="password", 
                         help="Necesitas una clave de API de Google AI Studio para usar la IA.")

# 2. Subida del archivo PDF
uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo PDF", type="pdf")

if uploaded_file is not None:
    if not api_key:
        st.warning("⚠️ Por favor, introduce tu API Key de Gemini para poder procesar el archivo.")
    else:
        # Botón para iniciar el proceso
        if st.button("✨ Procesar y Generar Word"):
            with st.spinner("Leyendo el PDF y analizando con IA... Por favor espera."):
                try:
                    # Configurar la IA con la clave provista
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    
                    # Extraer el texto del PDF
                    texto_pdf = ""
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page in pdf.pages:
                            texto_pdf += page.extract_text() + "\n"
                    
                    if not texto_pdf.strip():
                        st.error("No se pudo extraer texto del PDF. Asegúrate de que no sea solo una imagen escaneada.")
                    else:
                        # Prompt estricto para evitar que la IA invente datos
                        prompt = f"""
                        Actúa como un asistente médico experto, sumamente preciso y ético. 
                        Tu tarea es leer el siguiente texto y extraer ÚNICAMENTE las recomendaciones médicas, 
                        planes de cuidado, tratamientos, dosis o advertencias que ya existan en el documento.
                        
                        CRUCIAL: No inventes nada, no asumas diagnósticos y no agregues información externa. 
                        Si el texto no contiene recomendaciones claras, indícalo.
                        Organiza el resultado con subtítulos claros y viñetas profesionales.
                        
                        Texto del documento:
                        {texto_pdf}
                        """
                        
                        # Llamar a la IA
                        respuesta = model.generate_content(prompt)
                        resultado_texto = respuesta.text
                        
                        # Mostrar vista previa en la web
                        st.subheader("📄 Vista previa de las recomendaciones:")
                        st.markdown(resultado_texto)
                        
                        # --- CREAR EL ARCHIVO WORD EN MEMORIA ---
                        doc = Document()
                        doc.add_heading('Recomendaciones Médicas Extraídas', 0)
                        
                        # Dividir el texto por líneas para intentar darle una estructura básica en Word
                        for linea in resultado_texto.split('\n'):
                            if linea.startswith('##'):
                                doc.add_heading(linea.replace('##', '').strip(), level=1)
                            elif linea.startswith('#'):
                                doc.add_heading(linea.replace('#', '').strip(), level=2)
                            elif linea.strip().startswith('*') or linea.strip().startswith('-'):
                                doc.add_paragraph(linea.strip(), style='List Bullet')
                            else:
                                if linea.strip():
                                    doc.add_paragraph(linea.strip())
                        
                        # Guardar en un buffer de bytes para descarga web
                        b_io = io.BytesIO()
                        doc.save(b_io)
                        b_io.seek(0)
                        
                        st.success("¡Procesamiento completado con éxito!")
                        
                        # Botón para descargar el Word
                        st.download_button(
                            label="📥 Descargar archivo de Word (.docx)",
                            data=b_io,
                            file_name="Recomendaciones_Medicas.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                except Exception as e:
                    st.error(f"Ocurrió un error durante el proceso: {e}")
