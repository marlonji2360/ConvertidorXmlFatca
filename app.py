import streamlit as st
import conversorArchivos  as ca


st.image('banner_bantrab.png')
st.title(":blue[SISTEMA CONVERSIÓN XML FATCA]")


anio = st.text_input('Ingresa el año a declarar')

datos = st.file_uploader("Cargar archivo EXCEL:", type=['xlsx'])

if datos is not None:
    # Botón para llamar a las funciones
    if st.button("Ejecutar"):
        ca.csv_to_sql('DATOS_FATCA', datos)
        ca.sql_to_xml('DATOS_FATCA',"C:\\Temp\\archivo_banco.xml",anio)
        ca.xml_casa_bolsa("C:\\Temp\\archivo_casa_de_bolsa.xml",anio)
        ca.xml_financiera("C:\\Temp\\archivo_financiera.xml",anio)
        ca.xml_aseguradora("C:\\Temp\\archivo_aseguridad.xml",anio)
        st.success("Se han generado los archivos XML en la carpeta TEMP del disco local C:")
