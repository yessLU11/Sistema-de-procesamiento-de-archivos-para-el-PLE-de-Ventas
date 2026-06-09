#2_📊_Validador_Documental.py
"""
CONVERTIDOR SIRE/PLE - SUNAT
Doble Formato + Archivos Grandes (hasta 500MB)
"""
import streamlit as st
import pandas as pd
import os
import logging
import datetime
import pathlib
import time
from config import APP_TITLE, APP_ICON, INPUT_DIR, OUTPUT_DIR, LOG_DIR
import gc
import io
from src.etl.processor import procesar_excel

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Validator Doc v1.0",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Validación Documental & Detección de Duplicados")
st.markdown("---")
st.info("Sube tu Excel (.xlsx) para normalizar códigos y detectar registros duplicados en documentos 01 y 03.")
# 1. ÁREA DE SUBIDA (UPLOAD)
archivo_subido = st.file_uploader("Sube tu Excel (.xlsx)", type=['xlsx'])

if archivo_subido:
    # Botón de ejecución
    if st.button("🚀 Ejecutar Validación"):
        
        start_time = time.perf_counter()
        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def progress_callback(value, message=''):
            try:
                progress_bar.progress(value)
            except Exception:
                pass
            if message:
                status_text.info(message)

        with st.spinner('Procesando documento...'):
            # Llamamos al módulo ETL alojado en src
            resultado = procesar_excel(archivo_subido, progress_callback=progress_callback)

        elapsed = time.perf_counter() - start_time
        progress_callback(1.0, f'Finalizado en {elapsed:.2f} segundos')

        if resultado['success']:
            # --- DASHBOARD DE MÉTRICAS ---
            st.subheader("📋 Resumen del Procesamiento")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Filas Leídas", resultado['total_registros'])
            col2.metric("Hojas Procesadas", len(resultado['hojas_procesadas']))
            col3.metric("Duplicados Detectados", resultado['duplicados_count'], delta_color="inverse")
            col4.metric("Tiempo total", f"{elapsed:.2f} s")

            # Mostrar errores o advertencias si existen
            if resultado['errores']:
                with st.expander("⚠️ Advertencias / Logs"):
                    for err in resultado['errores']:
                        st.warning(err)

            # --- VISUALIZACIÓN DE DUPLICADOS ---
            if resultado['duplicados_count'] > 0:
                st.subheader("🔍 Registros Duplicados encontrados")
                st.dataframe(resultado['df_duplicados'], use_container_width=True)
                
                # --- DESCARGA DE DUPLICADOS (En memoria dinámica para evitar bloqueos) ---
                buffer_duplicados = io.BytesIO()
                with pd.ExcelWriter(buffer_duplicados, engine='xlsxwriter') as writer:
                    resultado['df_duplicados'].to_excel(writer, index=False, sheet_name='Duplicados')
                
                # Rebobinar el búfer al inicio
                buffer_duplicados.seek(0)

                # Streamlit download button
                st.download_button(
                    label="📥 Descargar Archivo de Duplicados",
                    data=buffer_duplicados,
                    file_name=f"duplicados_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.success("✅ No se detectaron duplicados en documentos 01 y 03.")

        else:
            st.error(resultado['message'])