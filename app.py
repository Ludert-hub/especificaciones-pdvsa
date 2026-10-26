import os
import json
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from pypdf import PdfReader
from google import genai
from google.genai import types

st.set_page_config(page_title="Generador de Especificaciones PDVSA", page_icon="🏗️", layout="wide")

st.title("🏗️ Automatizador de Especificaciones Técnicas (PDVSA)")
st.markdown("Herramienta automatizada para transformar Análisis de Precios Unitarios (APU) en Especificaciones Técnicas oficiales con membrete de PDVSA, sin precios unitarios.")

api_key_secret = st.secrets.get("GEMINI_API_KEY", "")
st.sidebar.header("🔑 Configuración")
api_key_input = st.sidebar.text_input("API Key de Google AI Studio", value=api_key_secret, type="password")

if not api_key_input:
    st.warning("⚠️ Por favor ingresa tu API Key en la barra lateral o en los Secrets de Streamlit.")
    st.stop()

client = genai.Client(api_key=api_key_input)

st.sidebar.divider()
st.sidebar.subheader("📂 Archivos Requeridos")
pdf_file = st.sidebar.file_uploader("Sube tu ANALISIS.pdf", type=["pdf"])
excel_molde = st.sidebar.file_uploader("Sube tu ESPECIFICACIONES TECNICAS.xlsx", type=["xlsx"])

if pdf_file and excel_molde:
    st.success("✅ ¡Archivos cargados correctamente en el sistema!")
    
    if st.button("🚀 Generar Archivo de Especificaciones PDVSA", type="primary"):
        with st.spinner("Procesando documento PDF y generando especificaciones con IA..."):
            try:
                reader = PdfReader(pdf_file)
                total_paginas = len(reader.pages)
                st.info(f"📄 Procesando {total_paginas} páginas de análisis de precios...")
                
                wb = openpyxl.load_workbook(excel_molde)
                if "COMPUTOS" not in wb.sheetnames or "ESPECIFICACIONES" not in wb.sheetnames:
                    st.error("❌ El archivo Excel adjunto no tiene las pestañas requeridas ('COMPUTOS' y 'ESPECIFICACIONES').")
                    st.stop()
                
                ws_comp = wb["COMPUTOS"]
                ws_esp = wb["ESPECIFICACIONES"]
                
                partidas_procesadas = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, page in enumerate(reader.pages):
                    texto_pagina = page.extract_text()
                    if not texto_pagina or "N° Partida:" not in texto_pagina:
                        continue
                        
                    partidas_procesadas += 1
                    status_text.text(f"Procesando partida de la página {idx + 1} de {total_paginas}...")
                    progress_bar.progress((idx + 1) / total_paginas)
                    
                    prompt_ia = f"""
                    Extrae del siguiente APU la información solicitada en formato JSON estricto:
                    {{
                        "numero_partida": int,
                        "descripcion_partida": "texto de la partida",
                        "unidad": "UND, M2, M3, ML, MTS, PTO, JGO, etc.",
                        "cantidad_partida": float,
                        "rendimiento": float,
                        "materiales": [{{"descripcion": "...", "unidad": "...", "cantidad": float}}],
                        "equipos": [{{"descripcion": "...", "cantidad": float}}],
                        "mano_de_obra": [{{"descripcion": "...", "cantidad": float}}],
                        "extensión_del_trabajo": "Redacción técnica detallada estilo PDVSA",
                        "procedimiento_de_ejecucion": "Paso a paso técnico usando los materiales y equipos",
                        "medida_y_forma_de_pago": "Esta partida se pagará por...",
                        "no_conformidad": "PDVSA no reconocerá costo ni pago alguno..."
                    }}
                    REGLAS: CERO PRECIOS O COSTOS. Usa estrictamente los insumos del texto.
                    TEXTO: {texto_pagina}
                    """
                    
                    # Actualizado al modelo actual de tu AI Studio
                    resp = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=prompt_ia,
                        config=types.GenerateContentConfig(temperature=0.0, response_mime_type="application/json")
                    )
                    
                    datos = json.loads(resp.text)
                    fila_comp = 10 + partidas_procesadas
                    
                    ws_comp[f"B{fila_comp}"] = datos.get("numero_partida", partidas_procesadas)
                    ws_comp[f"C{fila_comp}"] = datos.get("descripcion_partida", "")
                    ws_comp[f"D{fila_comp}"] = datos.get("unidad", "UND")
                    ws_comp[f"E{fila_comp}"] = datos.get("cantidad_partida", 1.0)
                    
                    offset = (partidas_procesadas - 1) * 78
                    ws_esp[f"L{6 + offset}"] = f"=+COMPUTOS!B{fila_comp}"
                    ws_esp[f"L{9 + offset}"] = f"=+COMPUTOS!D{fila_comp}"
                    ws_esp[f"A{11 + offset}"] = f"=+COMPUTOS!C{fila_comp}"
                    ws_esp[f"L{11 + offset}"] = f"=+COMPUTOS!E{fila_comp}"
                    ws_esp[f"M{11 + offset}"] = f"=+L{9 + offset}"
                    ws_esp[f"O{13 + offset}"] = datos.get("rendimiento", 1.0)
                    ws_esp[f"P{13 + offset}"] = f"=+M{11 + offset}"
                    ws_esp[f"Q{13 + offset}"] = "/dia"
                    ws_esp[f"O{15 + offset}"] = f"=+L{11 + offset}/O{13 + offset}"
                    
                    ws_esp[f"A{14 + offset}"] = datos.get("extensión_del_trabajo", "")
                    ws_esp[f"A{17 + offset}"] = datos.get("procedimiento_de_ejecucion", "")
                    ws_esp[f"A{70 + offset}"] = datos.get("medida_y_forma_de_pago", "")
                    ws_esp[f"A{72 + offset}"] = datos.get("no_conformidad", "")
                    
                    for i, mat in enumerate(datos.get("materiales", [])[:12]):
                        r = 22 + offset + i
                        ws_esp[f"A{r}"] = mat.get("descripcion", "")
                        ws_esp[f"G{r}"] = mat.get("unidad", "")
                        ws_esp[f"H{r}"] = mat.get("cantidad", 0)
                        
                    for i, eq in enumerate(datos.get("equipos", [])[:15]):
                        r = 36 + offset + i
                        ws_esp[f"A{r}"] = eq.get("descripcion", "")
                        ws_esp[f"H{r}"] = eq.get("cantidad", 0)
                        
                    for i, mo in enumerate(datos.get("mano_de_obra", [])[:15]):
                        r = 52 + offset + i
                        ws_esp[f"A{r}"] = mo.get("descripcion", "")
                        ws_esp[f"H{r}"] = mo.get("cantidad", 0)

                status_text.text("¡Proceso completado con éxito!")
                progress_bar.progress(1.0)
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"🎉 ¡Se han procesado exitosamente {partidas_procesadas} partidas bajo el estricto formato de PDVSA!")
                
                st.download_button(
                    label="📥 Descargar Archivo Excel Oficial (PDVSA)",
                    data=output,
                    file_name="ESPECIFICACIONES_TECNICAS_PDVSA_FINAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Ocurrió un error durante el procesamiento: {str(e)}")
else:
    st.info("👈 Por favor sube tu PDF de análisis y tu plantilla Excel en la barra lateral para habilitar el generador.")
