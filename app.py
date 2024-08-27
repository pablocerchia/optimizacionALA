# import streamlit as st
# import numpy as np
# import pandas as pd

# st.title("Separador excels")

# df = pd.read_excel(r'C:\Users\BB\Downloads\Optimizacion ALA\Separador de excels\Archivo formula Argentores.xlsb',
#                    sheet_name='AMERICA')


# df.drop(columns=df.columns[:3], axis=1, inplace=True)
# df.drop(index=0, inplace=True)

# new_header = df.iloc[0].tolist() 
# df = df[1:] 
# df.columns = new_header
# # df.columns = pd.io.parsers.base_parser.ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)

# # st.dataframe(df,use_container_width=True)
# st.write("hi")

import streamlit as st
import pandas as pd
import zipfile
import os

st.title("Separador excels")

anio=st.selectbox("Seleccionar año",("2024", "2025", "2026"))
mes=st.selectbox("Seleccionar mes",("01", "02", "03","04","05","06", "07", "08","09","10","11","12"))
dia="01"


# Subir el archivo Excel
uploaded_file = st.file_uploader("Subi el Excel de Argentores", type=["xlsb", "xlsx"])

if uploaded_file is not None:
    # Leer todas las hojas del Excel en un diccionario de DataFrames
    xls = pd.ExcelFile(uploaded_file)
    
    # Crear una carpeta temporal para guardar los archivos Excel procesados
    os.makedirs("temp_excels", exist_ok=True)
    
    # Iterar sobre cada hoja
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Aplicar el proceso de limpieza
        df.drop(columns=df.columns[:3], axis=1, inplace=True)
        df.drop(index=0, inplace=True)
        
        new_header = df.iloc[0].tolist() 
        df = df[1:] 
        df.columns = new_header
        
        sheet_name2= sheet_name.replace(" ", "_")
        # Guardar cada DataFrame en un archivo Excel separado
        output_filename = f"temp_excels/{anio}{mes}{dia}_{sheet_name2}_Business_Bureau_(Argentores).xlsx"
        df.to_excel(output_filename, index=False)
    
    # Crear un archivo ZIP con todos los Excel procesados
    zip_filename = "procesados.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, _, files in os.walk("temp_excels"):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)
    
    # Ofrecer el archivo ZIP para descarga
    with open(zip_filename, "rb") as fp:
        st.download_button(label="Descargar Excel Procesados", data=fp, file_name=zip_filename, mime="application/zip")
    
    # Limpiar archivos temporales
    os.remove(zip_filename)
    for file in os.listdir("temp_excels"):
        os.remove(os.path.join("temp_excels", file))
    os.rmdir("temp_excels")