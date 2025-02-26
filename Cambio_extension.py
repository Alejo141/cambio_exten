import os
import pandas as pd
import streamlit as st
from io import BytesIO

def convertir_xlsm_a_xlsx(archivo):
    """Convierte un archivo .xlsm a .xlsx sin macros"""
    df = pd.read_excel(archivo, sheet_name=None)  # Cargar todas las hojas
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for hoja, contenido in df.items():
            contenido.to_excel(writer, sheet_name=hoja, index=False)
    
    output.seek(0)
    return output

# Interfaz de Streamlit
st.title("Convertidor de archivos XLSM a XLSX")
st.write("Carga uno o varios archivos `.xlsm` y conviértelos a `.xlsx`.")

archivos_subidos = st.file_uploader("Selecciona archivos", type=["xlsm"], accept_multiple_files=True)

if archivos_subidos:
    for archivo in archivos_subidos:
        nombre_sin_ext = os.path.splitext(archivo.name)[0]
        archivo_convertido = convertir_xlsm_a_xlsx(archivo)

        st.download_button(
            label=f"Descargar {nombre_sin_ext}.xlsx",
            data=archivo_convertido,
            file_name=f"{nombre_sin_ext}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
