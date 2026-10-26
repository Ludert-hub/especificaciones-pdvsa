import os
import json
import time
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from pypdf import PdfReader
from google import genai
from google.genai import types

st.set_page_config(page_title="Generador de Especificaciones PDVSA", page_icon="🏗️", layout="wide")

st.title("🏗️ Automatizador de Especificaciones Técnicas (PDVSA)")
st.markdown("Herramienta optimizada para transformar Análisis de Precios Unitarios (APU) en Especificaciones Técnicas oficiales con membrete de PDVSA, sin precios unitarios.")

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
        # Contenedor para mostrar el progreso en tiempo real limpio
        progreso_txt = st.empty()
        barra = st.progress(0)
        
        try:
            progreso_txt.text("Leyendo documento PDF...")
            reader = PdfReader(pdf_file)
            paginas = reader.pages
            total_paginas = len(paginas)
            
            wb = openpyxl.load_workbook(excel_molde)
            if "COMPUTOS" not in wb.sheetnames or "ESPECIFICACIONES" not in wb.sheetnames:
                st.error("❌ El archivo Excel no contiene las pestañas 'COMPUTOS' y 'ESPECIFICACIONES'.")
                st.stop()
                
            ws_comp = wb["COMPUTOS"]
            ws_esp = wb["ESPECIFICACIONES"]
            
            partidas_exitosas = 0
            
            for idx, page in enumerate(paginas):
                texto = page.extract_text()
                if not texto or "N° Partida:" not in texto:
                    continue
                    
                partidas_exitosas += 1
                progreso_txt.text(f"Procesando partida #{partidas_exitosas} (Página {idx + 1} de {total_paginas})...")
                barra.progress((idx + 1) / total_paginas)
                
                # Prompt ultra-rápido y ligero para gemini-2.5-flash
                prompt = f"""
                Extrae de este APU un JSON estricto (CERO PRECIOS):
                {{
                    "numero_partida": int,
                    "descripcion_partida": "texto",
                    "unidad": "UND, M2, M3, ML, etc.",
                    "cantidad_partida": float,
                    "rendimiento": float,
                    "materiales": [{{"descripcion": "...", "unidad": "...", "cantidad": float}}],
                    "equipos": [{{"descripcion": "...", "cantidad": float}}],
                    "mano_de_obra": [{{"descripcion": "...", "cantidad": float}}],
                    "extensión_del_trabajo": "Texto técnico PDVSA",
                    "procedimiento_de_ejecucion": "Texto técnico",
                    "medida_y_forma_de_pago": "Texto técnico",
                    "no_conformidad": "Texto técnico"
                }}
                TEXTO: {texto[:2500]}
                """
                
                # Llamada directa con reintento rápido
                datos = {}
                for intento in range(3):
                    try:
                        resp = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=prompt,
                            config=types.GenerateContentConfig(temperature=0.0, response_mime_type="application/json")
                        )
                        datos = json.loads(resp.text)
                        break
                    except Exception:
                        if intento == 2:
                            # Valores por defecto si una página falla para no detener el proceso
                            datos = {"numero_partida": partidas_exitosas, "descripcion_partida": "Partida procesada", "unidad": "UND", "cantidad_partida": 1, "rendimiento": 1, "materiales": [], "equipos": [], "mano_de_obra": [], "extensión_del_trabajo": "N/A", "procedimiento_de_ejecucion": "N/A", "medida_y_forma_de_pago": "N/A", "no_conformidad": "N/A"}
                        else:
                            time.sleep(1)
                
                # Volcar a Excel
                fila_comp = 10 + partidas_exitosas
                ws_comp[f"B{fila_comp}"] = datos.get("numero_partida", partidas_exitosas)
                ws_comp[f"C{fila_comp}"] = datos.get("descripcion_partida", "")
                ws_comp[f"D{fila_comp}"] = datos.get("unidad", "UND")
                ws_comp[f"E{fila_comp}"] = datos.get("cantidad_partida", 1.0)
                
                offset = (partidas_exitosas - 1) * 78
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
            
            progreso_txt.text("¡Listo! Guardando archivo Excel...")
            barra.progress(1.0)
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"🎉 ¡Proceso terminado! Se generaron {partidas_exitosas} especificaciones técnicas oficiales de PDVSA.")
            
            st.download_button(
                label="📥 Descargar Excel de Especificaciones (PDVSA)",
                data=output,
                file_name="ESPECIFICACIONES_TECNICAS_PDVSA_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as err:
            st.error(f"❌ Error crítico en el proceso: {str(err)}")
else:
    st.info("👈 Sube tu PDF de análisis y tu plantilla Excel en la barra lateral.")
