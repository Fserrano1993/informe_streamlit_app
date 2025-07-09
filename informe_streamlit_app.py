
import streamlit as st
from docx import Document
from PIL import Image
import os
from io import BytesIO
import base64

st.set_page_config(page_title="GeneradorTS", layout="centered")

st.title("📄 Generador de Informes TS")

# Subida de plantilla Word
plantilla_file = st.file_uploader("Selecciona plantilla Word (.docx)", type=["docx"])

# Subida de Excel (si se usa)
excel_file = st.file_uploader("Selecciona archivo Excel de pólizas (opcional)", type=["xlsx"])

# Área de texto para contenido
texto_encargo = st.text_area("Pega aquí el texto del encargo:", height=200)

# Subida de PDF o imagen catastro
catastro_file = st.file_uploader("Selecciona imagen/PDF de Catastro", type=["pdf", "png", "jpg", "jpeg"])

# Carpeta de fotos
fotos = st.file_uploader("Selecciona imágenes para el reportaje (máximo 6)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Botón de generación
if st.button("Generar informe"):
    if not plantilla_file or not texto_encargo:
        st.warning("Se requiere plantilla Word y texto de encargo.")
    else:
        doc = Document(plantilla_file)

        # Reemplazo de placeholders simples
        for p in doc.paragraphs:
            for run in p.runs:
                for key in ["{{EXPEDIENTE}}", "{{ASEGURADO}}", "{{DIR_CATASTRO}}", "{{PROVINCIA_CATASTRO}}"]:
                    if key in run.text:
                        run.text = run.text.replace(key, "VALOR")

        # Añadir imágenes al final
        if fotos:
            doc.add_page_break()
            doc.add_paragraph("Reportaje fotográfico", "Heading 1")

            table = doc.add_table(rows=3, cols=2)
            idx = 0
            for i in range(3):
                row_cells = table.rows[i].cells
                for j in range(2):
                    if idx < len(fotos):
                        img = Image.open(fotos[idx])
                        img_io = BytesIO()
                        img.save(img_io, format="PNG")
                        img_io.seek(0)
                        run = row_cells[j].paragraphs[0].add_run()
                        run.add_picture(img_io, width=docx.shared.Inches(2.5))
                        row_cells[j].add_paragraph(f"Imagen {idx+1}")
                        idx += 1

        # Guardar archivo en memoria
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        b64 = base64.b64encode(buffer.read()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="informe_generado.docx">📥 Descargar informe generado</a>'
        st.markdown(href, unsafe_allow_html=True)
