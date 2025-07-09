import streamlit as st
from docx import Document
import pandas as pd
from PyPDF2 import PdfReader
from PIL import Image
import io

st.set_page_config(page_title="Generador de Informes TS", layout="centered")
st.title("📝 Generador de Informes TS")

# Subir plantilla Word
st.subheader("Plantilla Word (.docx)")
plantilla_file = st.file_uploader("Cargar plantilla base", type="docx")

# Subir Excel de pólizas
st.subheader("Modelo de póliza (Excel)")
poliza_file = st.file_uploader("Cargar archivo .xlsx", type="xlsx")

# Pegar texto de encargo
st.subheader("Texto del encargo")
texto_encargo = st.text_area("Pega aquí el texto completo del encargo", height=250)

# Subir documento catastral
st.subheader("Documento del Catastro")
catastro_file = st.file_uploader("Cargar archivo catastral", type=["pdf", "png", "jpg", "jpeg"])

# Subir fotos para reportaje
st.subheader("Fotos para el reportaje fotográfico")
imagenes_files = st.file_uploader("Cargar imágenes", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Generar informe
if st.button("Generar informe"):
    if not plantilla_file:
        st.error("⚠️ Debes subir una plantilla Word.")
    elif not texto_encargo.strip():
        st.error("⚠️ Debes pegar el texto del encargo.")
    else:
        st.success("✅ Todo listo para generar informe.")
        st.write("ℹ️ Generación de informe aún no implementada.")
