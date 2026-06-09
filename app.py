#app.py
import streamlit as st
import os
from config import APP_TITLE, APP_ICON, INPUT_DIR, OUTPUT_DIR, LOG_DIR

# Configuración global de la página
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Asegurar que los directorios base existan
for dir_path in [INPUT_DIR, OUTPUT_DIR, LOG_DIR]:
    os.makedirs(dir_path, exist_ok=True)

# Diseño de la página de inicio
st.title("🏦 PLE DE VENTAS - Portal de herramientas")
st.markdown("---")
st.markdown("""
##### Bienvenido al sistema de procesamiento de archivos para el PLE de Ventas. 
##### Aquí encontrarás dos herramientas principales para ayudarte a convertir y validar tus documentos de manera eficiente.
#####Descripción: Esta aplicación está diseñada para facilitar la gestión de tus archivos relacionados con el PLE de Ventas, permitiéndote convertir archivos SIRE a Excel y validar documentos para detectar posibles duplicados.
            
Por favor, selecciona una herramienta en el menú lateral de la izquierda:

1. **🔄 Convertidor SIRE/PLE:** Convierte tus archivos TXT masivos a formato Excel.
2. **📊 Validador Documental:** Sube tus Excels para normalizar códigos y detectar duplicados.
""")

st.info("👈 Usa el menú lateral para navegar entre las aplicaciones.")

