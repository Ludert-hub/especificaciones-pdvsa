import os
import json
import streamlit as st
import pandas as pd
import openpyxl
from pypdf import PdfReader
from google import genai
from google.genai import types

st.set_page_config(page_title="Generador de Especificaciones PDVSA", page_icon="🏗️", layout="wide")

st.title("🏗️ Automatizador de Especificaciones Técnicas (PDVSA)")
st.markdown("Herramienta para transformar Análisis de Precios Unitarios (APU) en Especificaciones Técnicas oficiales con membrete de PDVSA, sin precios unitarios y verificadas paso a paso.")

# Barra lateral para configurar la API Key
st.sidebar.header("🔑 Configuración")
api_key_input = st.sidebar.text_input("Ingresa tu API Key de Google AI Studio", type="password")

if not api_key_input:
    st.warning("⚠️ Por favor ingresa tu API Key en la barra lateral para continuar.")
    st.stop()

# Configurar cliente de Gemini
client = genai.Client(api_key=api_key_input)

st.sidebar.divider()
st.sidebar.subheader("📂 Archivos del Proyecto")
pdf_file = st.sidebar.file_uploader("Sube el PDF de Análisis de Precios (ANALISIS.pdf)", type=["pdf"])
excel_molde = st.sidebar.file_uploader("Sube tu plantilla base (ESPECIFICACIONES TECNICAS.xlsx)", type=["xlsx"])

if pdf_file and excel_molde:
    st.success("✅ ¡Archivos cargados correctamente!")
    
    if st.button("🚀 Iniciar Extracción y Análisis de Partidas", type="primary"):
        with st.spinner("Leyendo documento PDF de APUs..."):
            reader = PdfReader(pdf_file)
            total_paginas = len(reader.pages)
            st.info(f"📄 Se detectaron {total_paginas} páginas en el documento.")
            
            # Ejemplo de lectura de la primera página para verificar
            texto_ejemplo = reader.pages[0].extract_text()
            st.subheader("Vista previa de la primera página leída:")
            st.text(texto_ejemplo[:400] + "...")
            
            st.info("💡 Tu aplicación base está conectada y lista. En el siguiente paso configuraremos el bucle de procesamiento partida por partida.")
else:
    st.info("👈 Por favor carga ambos archivos en la barra lateral para habilitar el sistema.")
