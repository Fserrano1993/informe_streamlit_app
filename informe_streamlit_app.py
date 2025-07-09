import streamlit as st
from io import BytesIO
from docx import Document
from docx.shared import Inches
from PIL import Image
from PyPDF2 import PdfReader
import re
import zipfile
import os

st.set_page_config(page_title="Generador de Informes TS", layout="centered")

st.title("📝 Generador de Informes TS")

# --- Entradas ---
tipo_modelo = st.radio("Selecciona plantilla", ["Base", "Jurídica"], horizontal=True)

encargo_method = st.radio("¿Cómo quieres introducir el encargo?", ["Pegar texto", "Subir archivo .txt"], horizontal=True)
txt_encargo = ""

if encargo_method == "Pegar texto":
    txt_encargo = st.text_area("Pega el texto del encargo aquí")
else:
    txt_file = st.file_uploader("Sube el archivo .txt del encargo", type=["txt"])
    if txt_file:
        txt_encargo = txt_file.read().decode("utf-8")

catastro_file = st.file_uploader("Sube el PDF o imagen del Catastro (opcional)", type=["pdf", "png", "jpg", "jpeg"])
zip_fotos = st.file_uploader("Sube un .zip con las fotos del siniestro (opcional)", type=["zip"])

# --- Funciones ---
def parse_encargo(text):
    rep = {}
    campos = {
        "EXPEDIENTE": r"Expediente: (.+)",
        "CLASE": r"Clase: (.+)",
        "FECHA_DE_OCURRENCIA": r"Fecha de Ocurrencia: (.+)",
        "FECHA_COMUNICACION": r"Fecha de Comunicacion: (.+)",
        "TRAMITADOR": r"Tramitador del expediente: (.+)",
        "GARANTIA": r"Garantia afectada: (.+)",
        "POLIZA": r"Nº Póliza: (.+)",
        "EFECTO": r"Efecto: (.+)",
        "ASEGURADO": r"Asegurado: (.+)",
        "DIR_CATASTRO": r"Lugar: (.+)",
        "D.P.": r"D\.P\.: (.+)",
        "LOCALIDAD": r"Localidad: (.+)",
        "PROVINCIA_CATASTRO": r"Provincia: (.+)"
    }
    for key, pat in campos.items():
        m = re.search(pat, text)
        rep[f"{{{{{key}}}}}"] = m.group(1).strip() if m else ""
    return rep

def replace_placeholders(doc, replacements):
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        inline[i].text = inline[i].text.replace(key, val)

def add_catastro_image(doc, file):
    if file.name.endswith(".pdf"):
        reader = PdfReader(file)
        page = reader.pages[0]
        xObject = page.images[0] if hasattr(page, "images") and page.images else None
        return
    else:
        image = Image.open(file)
        image_path = "/tmp/img_catastro.png"
        image.save(image_path)
        doc.add_picture(image_path, width=Inches(5.5))

def add_zip_images(doc, zip_data):
    with zipfile.ZipFile(zip_data) as z:
        img_files = [f for f in z.namelist() if f.lower().endswith((".png", ".jpg", ".jpeg"))]
        img_files.sort()
        row, col = 0, 0
        table = doc.add_table(rows=3, cols=2)
        for i, img_name in enumerate(img_files[:6]):
            data = z.read(img_name)
            img_path = f"/tmp/tmp_img_{i}.png"
            with open(img_path, "wb") as f:
                f.write(data)
            cell = table.cell(row, col)
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(img_path, width=Inches(2.5))
            col += 1
            if col >= 2:
                col = 0
                row += 1

# --- Generar informe ---
if st.button("📄 Generar Informe"):
    if not txt_encargo:
        st.warning("Por favor, introduce el texto del encargo.")
    else:
        st.info("Generando documento...")
        replacements = parse_encargo(txt_encargo)

        # Seleccionar plantilla
        template_name = "PLANTILLA BASE JURIDICO - V3 generador.docx" if tipo_modelo == "Jurídica" else "PLANTILLA BASE - V3 generador.docx"
        template_path = os.path.join(os.getcwd(), template_name)
        doc = Document(template_path)

        replace_placeholders(doc, replacements)

        if catastro_file:
            add_catastro_image(doc, catastro_file)

        if zip_fotos:
            doc.add_page_break()
            doc.add_paragraph("Reportaje fotográfico")
            add_zip_images(doc, zip_fotos)

        output = BytesIO()
        doc.save(output)
        st.success("✅ Informe generado correctamente.")
        st.download_button("📥 Descargar Informe Word", output.getvalue(), file_name=f"{replacements.get('{{EXPEDIENTE}}','Informe')}.docx")
