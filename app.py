import streamlit as st
import conversorArchivos  as ca
from pathlib import Path



pathArchivosXML = Path(__file__).parent / "venv/XML/"
pathImages = Path(__file__).parent / "venv/images/"

st.image(pathImages+'banner_bantrab.png')
st.title(":blue[SISTEMA CONVERSIÓN XML FATCA]")


anio = st.text_input('Ingresa el año a declarar')

datos = st.file_uploader("Cargar archivo EXCEL:", type=['xlsx'])

if datos is not None:
    # Botón para llamar a las funciones
    if st.button("Ejecutar"):
        ca.csv_to_sql('DATOS_FATCA', datos)
        #ca.sql_to_xml('DATOS_FATCA',"C:\\Temp\\archivo_banco.xml",anio)
        #ca.xml_casa_bolsa("C:\\Temp\\archivo_casa_de_bolsa.xml",anio)
        #ca.xml_financiera("C:\\Temp\\archivo_financiera.xml",anio)
       
        archivo_banco=ca.sql_to_xml('DATOS_FATCA',pathArchivosXML+'archivo_banco.xml',anio)
        archivo_aseguradora=ca.xml_aseguradora(pathArchivosXML+'archivo_aseguradora.xml',anio) 
        archivo_casa_bolsa=ca.xml_casa_bolsa(pathArchivosXML+'archivo_casa_de_bolsa.xml',anio)
        archivo_financiara=ca.xml_financiera(pathArchivosXML+"archivo_financiera.xml",anio)

        ca.comprimir_archivos('Archivos XML',pathArchivosXML)

        st.success("Se han generado los archivos XML")

        with open ('archivo_banco.xml') as aseg:
            st.download_button('Descargar XML Banco', aseg,'archivo_banco.xml')
             
        with open ('archivo_aseguradora.xml') as aseg:
            st.download_button('Descargar XML Aseguradora', aseg,'archivo_aseguradora.xml')
        #st.success("Se han generado los archivos XML")
        
        