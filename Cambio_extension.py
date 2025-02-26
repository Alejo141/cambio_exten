import streamlit as st
import os
from io import BytesIO
from openpyxl import load_workbook

def cambiar_extension_xlsm_a_xlsx(archivo):
    """Carga un archivo .xlsm y lo guarda como .xlsx sin modificar su estructura"""
    output = BytesIO()
    wb = load_workbook(archivo)  # Cargar el archivo manteniendo su estructura
    wb.save(output)  # Guardarlo en memoria
    output.seek(0)
    return output

# Interfaz de Streamlit
st.title("Convertidor de XLSM a XLSX (sin modificar el archivo)")
st.write("Carga archivos `.xlsm` y simplemente cámbiales la extensión a `.xlsx`.")

archivos_subidos = st.file_uploader("Selecciona archivos", type=["xlsm"], accept_multiple_files=True)

if archivos_subidos:
    for archivo in archivos_subidos:
        nombre_sin_ext = os.path.splitext(archivo.name)[0]
        archivo_convertido = cambiar_extension_xlsm_a_xlsx(archivo)

        st.download_button(
            label=f"Descargar {nombre_sin_ext}.xlsx",
            data=archivo_convertido,
            file_name=f"{nombre_sin_ext}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
