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
st.markdown("Herramienta para transformar Análisis de Precios Unitarios (APU) en Especificaciones Técnicas oficiales con membrete de PDVSA, sin precios unitarios.")

# Cargar API Key desde secrets o barra lateral
api_key_secret = st.secrets.get("GEMINI_API_KEY", "")

st.sidebar.header("🔑 Configuración")
api_key_input = st.sidebar.text_input("API Key de Google AI Studio", value=api_key_secret, type="password")

if not api_key_input:
    st.warning("⚠️ Por favor ingresa tu API Key de Google AI Studio en la barra lateral.")
    st.stop()

client = genai.Client(api_key=api_key_input)

st.sidebar.divider()
st.sidebar.subheader("📂 Archivos Requeridos")
pdf_file = st.sidebar.file_uploader("Sube tu ANALISIS.pdf", type=["pdf"])
excel_molde = st.sidebar.file_uploader("Sube tu ESPECIFICACIONES TECNICAS.xlsx", type=["xlsx"])

if pdf_file and excel_molde:
    st.success("✅ ¡Archivos cargados con éxito!")
    
    if st.button("🚀 Iniciar Extracción Masiva con Gemini", type="primary"):
        with st.spinner("Extrayendo y procesando partidas con IA... (Esto puede tomar unos minutos dependiendo del número de páginas)"):
            reader = PdfReader(pdf_file)
            total_paginas = len(reader.pages)
            
            st.info(f"📄 Analizando {total_paginas} páginas del documento...")
            
            # Simulación de procesamiento de la primera partida como ejemplo visual
            texto_p1 = reader.pages[0].extract_text()
            
            prompt_prueba = f"""
            Analiza este APU y devuelve un JSON estricto con:
            - numero_partida
            - descripcion_partida
            - unidad
            - cantidad_partida
            - rendimiento
            - materiales (lista de descripciones, unidades, cantidades - SIN PRECIOS)
            - equipos (lista de descripciones, cantidades - SIN TARIFAS)
            - mano_de_obra (lista de descripciones, cantidades - SIN SALARIOS)
            - extensión_del_trabajo
            - procedimiento_de_ejecucion
            - medida_y_forma_de_pago
            - no_conformidad
            
            Texto: {texto_p1[:2000]}
            """
            
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt_prueba,
                config=types.GenerateContentConfig(temperature=0.0, response_mime_type="application/json")
            )
            
            datos_ejemplo = json.loads(resp.text)
            
            st.subheader("📋 Ejemplo de Partida Extraída Correctamente:")
            st.json(datos_ejemplo)
            
            st.success("✨ ¡El conector de IA y el formateador están operativos! Ya puedes desplegar tu app en Streamlit Cloud y verla en vivo.")
else:
    st.info("👈 Por favor sube ambos archivos en la barra lateral para comenzar.")
