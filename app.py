import streamlit as st
import conversorArchivos  as ca
from pathlib import Path
import os



#pathArchivosXML = Path(__file__).parent / "venv/XML/"
#pathImages = Path(__file__).parent / "venv/images/"

st.image('images/banner_bantrab.png')
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
       
        archivo_banco=ca.sql_to_xml('DATOS_FATCA','XML/1MMZME.00000.LE.320',anio)
        archivo_aseguradora=ca.xml_aseguradora('XML/1MMZME.00003.ME.320.xml',anio) 
        archivo_casa_bolsa=ca.xml_casa_bolsa('XML/1MMZME.00004.ME.320.xml',anio)
        archivo_financiara=ca.xml_financiera("XML/1MMZME.00002.ME.320.xml",anio)
        
        ca.formatearXml("XML/1MMZME.00000.LE.320.xml")
        ca.formatearXml("XML/1MMZME.00003.ME.320.xml")
        ca.formatearXml("XML/1MMZME.00004.ME.320.xml")
        ca.formatearXml("XML/1MMZME.00002.ME.320.xml")

        carpeta_a_comprimir = 'XML'
        archivo_zip_salida = 'archivo_comprimido.zip'
        ca.comprimir_carpeta(carpeta_a_comprimir, archivo_zip_salida)

        st.success("Se han generado los archivos XML")
        
       

        with open ('XML/1MMZME.00000.LE.320.xml') as aseg:
            st.download_button('Descargar XML Banco', aseg,'XML/1MMZME.00000.LE.320.xml')
             
        with open ('XML/1MMZME.00003.ME.320.xml') as aseg:
            st.download_button('Descargar XML Aseguradora', aseg,'XML/1MMZME.00003.ME.320.xml')
            
        with open ('XML/1MMZME.00004.ME.320.xml') as aseg:
            st.download_button('Descargar XML Casa de bolsa', aseg,'XML/1MMZME.00004.ME.320.xml')
        
        with open ('XML/1MMZME.00002.ME.320.xml') as aseg:
            st.download_button('Descargar XML Financiera', aseg,'XML/1MMZME.00002.ME.320.xml')
            
        
        
        