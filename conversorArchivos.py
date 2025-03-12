import uuid
import datetime
import pandas as pd
import sqlite3 as sql
import zipfile
import os

def xml_aseguradora(output_file, anio):
    with open(output_file, 'w', encoding='utf-8') as archivo:
        # Escribir en el archivo.
        archivo.write('<?xml version="1.0" encoding="utf-8"?>\n')
        archivo.write('<ftc:FATCA_OECD xmlns:xsi="http:www.w3.org2001XMLSchema-instance" xsi:schemaLocation="urn:oecd:ties:fatca:v2" xmlns:iso="urn:oecd:ties:isofatcatypes:v1" xmlns:ftc="urn:oecd:ties:fatca:v2" xmlns:stf="urn:oecd:ties:stf:v4" xmlns:sfa="urn:oecd:ties:stffatcatypes:v2" version="2.0">\n')
        archivo.write('<ftc:MessageSpec>\n')
        archivo.write('<sfa:SendingCompanyIN>1MMZME.00003.ME.320</sfa:SendingCompanyIN>\n')
        archivo.write('<sfa:TransmittingCountry>GT</sfa:TransmittingCountry>\n')
        archivo.write('<sfa:ReceivingCountry>US</sfa:ReceivingCountry>\n')
        archivo.write('<sfa:MessageType>FATCA</sfa:MessageType>\n')
        archivo.write('<sfa:MessageRefId>'+str(uuid.uuid4())+'</sfa:MessageRefId>\n')
        archivo.write('<sfa:ReportingPeriod>'+anio+'-12-31</sfa:ReportingPeriod>\n')
        archivo.write('<sfa:Timestamp>'+str(datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))+'</sfa:Timestamp>\n')
        archivo.write('</ftc:MessageSpec>\n')

        archivo.write('<ftc:FATCA>\n')
        archivo.write('<ftc:ReportingFI>\n')
        archivo.write('<sfa:ResCountryCode>GT</sfa:ResCountryCode>\n')
        archivo.write('<sfa:TIN issuedBy="US">1MMZME.00003.ME.320</sfa:TIN>\n')
        archivo.write('<sfa:Name>ASEGURADORA DE LOS TRABAJADORES</sfa:Name>\n')
        archivo.write('<sfa:Address>\n')
        archivo.write('<sfa:CountryCode>GT</sfa:CountryCode>\n')
        archivo.write('<sfa:AddressFree>AVENIDA REFORMA 6 20 ZONA 9 GUATEMALA</sfa:AddressFree>\n')
        archivo.write('</sfa:Address>\n')
        archivo.write('<ftc:FilerCategory>FATCA601</ftc:FilerCategory>')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00003.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('</ftc:ReportingFI>\n')
        archivo.write('<ftc:ReportingGroup>\n')
        archivo.write('<ftc:NilReport>\n')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00003.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('<ftc:NoAccountToReport>yes</ftc:NoAccountToReport>\n')
        archivo.write('</ftc:NilReport>\n')
        archivo.write('</ftc:ReportingGroup>\n')
        archivo.write('</ftc:FATCA>\n')
        archivo.write('</ftc:FATCA_OECD>\n')
        return archivo

def xml_casa_bolsa(output_file, anio):
    # Abrir un archivo en modo de escritura ('w' para escritura).
    # Si el archivo no existe, se creará. Si ya existe, se sobrescribirá.
    with open(output_file, 'w', encoding='utf-8') as archivo:
        # Escribir en el archivo.
        archivo.write('<?xml version="1.0" encoding="utf-8"?>\n')
        archivo.write('<ftc:FATCA_OECD xmlns:xsi="http:www.w3.org2001XMLSchema-instance" xsi:schemaLocation="urn:oecd:ties:fatca:v2" xmlns:iso="urn:oecd:ties:isofatcatypes:v1" xmlns:ftc="urn:oecd:ties:fatca:v2" xmlns:stf="urn:oecd:ties:stf:v4" xmlns:sfa="urn:oecd:ties:stffatcatypes:v2" version="2.0">\n')
        archivo.write('<ftc:MessageSpec>\n')
        archivo.write('<sfa:SendingCompanyIN>1MMZME.00004.ME.320</sfa:SendingCompanyIN>\n')
        archivo.write('<sfa:TransmittingCountry>GT</sfa:TransmittingCountry>\n')
        archivo.write('<sfa:ReceivingCountry>US</sfa:ReceivingCountry>\n')
        archivo.write('<sfa:MessageType>FATCA</sfa:MessageType>\n')
        archivo.write('<sfa:MessageRefId>'+str(uuid.uuid4())+'</sfa:MessageRefId>\n')
        archivo.write('<sfa:ReportingPeriod>'+anio+'-12-31</sfa:ReportingPeriod>\n')
        archivo.write('<sfa:Timestamp>'+str(datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))+'</sfa:Timestamp>\n')
        archivo.write('</ftc:MessageSpec>\n')

        archivo.write('<ftc:FATCA>\n')
        archivo.write('<ftc:ReportingFI>\n')
        archivo.write('<sfa:ResCountryCode>GT</sfa:ResCountryCode>\n')
        archivo.write('<sfa:TIN issuedBy="US">1MMZME.00004.ME.320</sfa:TIN>\n')
        archivo.write('<sfa:Name>CASA DE BOLSA DE LOS TRABAJADORES</sfa:Name>\n')
        archivo.write('<sfa:Address>\n')
        archivo.write('<sfa:CountryCode>GT</sfa:CountryCode>\n')
        archivo.write('<sfa:AddressFree>AVENIDA REFORMA 6 20 ZONA 9 GUATEMALA</sfa:AddressFree>\n')
        archivo.write('</sfa:Address>\n')
        archivo.write('<ftc:FilerCategory>FATCA601</ftc:FilerCategory>')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00004.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('</ftc:ReportingFI>\n')
        archivo.write('<ftc:ReportingGroup>\n')
        archivo.write('<ftc:NilReport>\n')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00004.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('<ftc:NoAccountToReport>yes</ftc:NoAccountToReport>\n')
        archivo.write('</ftc:NilReport>\n')
        archivo.write('</ftc:ReportingGroup>\n')
        archivo.write('</ftc:FATCA>\n')
        archivo.write('</ftc:FATCA_OECD>\n')

def xml_financiera(output_file, anio):
    # Abrir un archivo en modo de escritura ('w' para escritura).
    # Si el archivo no existe, se creará. Si ya existe, se sobrescribirá.
    with open(output_file, 'w', encoding='utf-8') as archivo:
        # Escribir en el archivo.
        archivo.write('<?xml version="1.0" encoding="utf-8"?>\n')
        archivo.write('<ftc:FATCA_OECD xmlns:xsi="http:www.w3.org2001XMLSchema-instance" xsi:schemaLocation="urn:oecd:ties:fatca:v2" xmlns:iso="urn:oecd:ties:isofatcatypes:v1" xmlns:ftc="urn:oecd:ties:fatca:v2" xmlns:stf="urn:oecd:ties:stf:v4" xmlns:sfa="urn:oecd:ties:stffatcatypes:v2" version="2.0">\n')
        archivo.write('<ftc:MessageSpec>\n')
        archivo.write('<sfa:SendingCompanyIN>1MMZME.00002.ME.320</sfa:SendingCompanyIN>\n')
        archivo.write('<sfa:TransmittingCountry>GT</sfa:TransmittingCountry>\n')
        archivo.write('<sfa:ReceivingCountry>US</sfa:ReceivingCountry>\n')
        archivo.write('<sfa:MessageType>FATCA</sfa:MessageType>\n')
        archivo.write('<sfa:MessageRefId>'+str(uuid.uuid4())+'</sfa:MessageRefId>\n')
        archivo.write('<sfa:ReportingPeriod>'+anio+'-12-31</sfa:ReportingPeriod>\n')
        archivo.write('<sfa:Timestamp>'+str(datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))+'</sfa:Timestamp>\n')
        archivo.write('</ftc:MessageSpec>\n')

        archivo.write('<ftc:FATCA>\n')
        archivo.write('<ftc:ReportingFI>\n')
        archivo.write('<sfa:ResCountryCode>GT</sfa:ResCountryCode>\n')
        archivo.write('<sfa:TIN issuedBy="US">1MMZME.00002.ME.320</sfa:TIN>\n')
        archivo.write('<sfa:Name>FINANCIERA DE LOS TRABAJADORES</sfa:Name>\n')
        archivo.write('<sfa:Address>\n')
        archivo.write('<sfa:CountryCode>GT</sfa:CountryCode>\n')
        archivo.write('<sfa:AddressFree>AVENIDA REFORMA 6 20 ZONA 9 GUATEMALA</sfa:AddressFree>\n')
        archivo.write('</sfa:Address>\n')
        archivo.write('<ftc:FilerCategory>FATCA601</ftc:FilerCategory>')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00002.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('</ftc:ReportingFI>\n')
        archivo.write('<ftc:ReportingGroup>\n')
        archivo.write('<ftc:NilReport>\n')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00002.ME.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('<ftc:NoAccountToReport>yes</ftc:NoAccountToReport>\n')
        archivo.write('</ftc:NilReport>\n')
        archivo.write('</ftc:ReportingGroup>\n')
        archivo.write('</ftc:FATCA>\n')
        archivo.write('</ftc:FATCA_OECD>\n')
        


def csv_to_sql(table, input_file):
    # Conexión a la base de datos SQL Server
    connection = sql.connect('FATCA.db')
    cursor = connection.cursor()

    archivo_xmlx = input_file
    sheet_name = 'Hoja1'
    
    # Lectura del archivo Excel utilizando pandas
    data = pd.read_excel(archivo_xmlx, sheet_name=sheet_name)
    

    query = f"DELETE FROM {table}"
    cursor.execute(query)

   # Iterar sobre cada fila del DataFrame y ejecutar una consulta de inserción para insertar los datos en la tabla
    for index, row in data.iterrows():
        # Reemplazar los valores nulos con un valor predeterminado (por ejemplo, None)
        row = [value if not pd.isnull(value) else None for value in row]
        # Crear la consulta de inserción
        query = f"INSERT INTO {table} (NO, ID, CODIGO_CLIENTE, NOMBRE_COMPLETO, PAIS, TIN, LARGO, PRIMER_NOMBRE, SEGUNDO_NOMBRE, PRIMER_APELLIDO, SEGUNDO_APELLIDO, PAIS_DIRECCION, DIRECCION, NUMERO, CODIGO_POSTAL, CIUDAD, ESTADO, FECHA_NACIMIENTO, CAPITAL, INTERESES, TOTAL_GENERAL) VALUES (?, ?, ?, UPPER(?), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ROUND(?,2), ROUND(?,2), ROUND(?,2))"  # Ajusta las columnas según tu tabla
        
        # Ejecutar la consulta de inserción con los valores de la fila actual
        cursor.execute(query, tuple(row))
        connection.commit() 

def sql_to_xml(table, output_file, anio):
    # Conexión a la base de datos MySQL
    # Conexión a la base de datos SQL Server
    connection = sql.connect('FATCA.db')
    cursor = connection.cursor()

    # Consulta SQL para recuperar los datos de la tabla
    query = f"SELECT * FROM {table}"
    cursor.execute(query)
    rows = cursor.fetchall()
# Abrir un archivo en modo de escritura ('w' para escritura).
# Si el archivo no existe, se creará. Si ya existe, se sobrescribirá.
    with open(output_file, 'w', encoding='utf-8') as archivo:
        # Escribir en el archivo.
        archivo.write('<?xml version="1.0" encoding="utf-8"?>\n')
        archivo.write('<ftc:FATCA_OECD xmlns:xsi="http:www.w3.org2001XMLSchema-instance" xsi:schemaLocation="urn:oecd:ties:fatca:v2" xmlns:iso="urn:oecd:ties:isofatcatypes:v1" xmlns:ftc="urn:oecd:ties:fatca:v2" xmlns:stf="urn:oecd:ties:stf:v4" xmlns:sfa="urn:oecd:ties:stffatcatypes:v2" version="2.0">\n')
        archivo.write('<ftc:MessageSpec>\n')
        archivo.write('<sfa:SendingCompanyIN>1MMZME.00000.LE.320</sfa:SendingCompanyIN>\n')
        archivo.write('<sfa:TransmittingCountry>GT</sfa:TransmittingCountry>\n')
        archivo.write('<sfa:ReceivingCountry>US</sfa:ReceivingCountry>\n')
        archivo.write('<sfa:MessageType>FATCA</sfa:MessageType>\n')
        archivo.write('<sfa:MessageRefId>'+str(uuid.uuid4())+'</sfa:MessageRefId>\n')
        archivo.write('<sfa:ReportingPeriod>'+anio+'-12-31</sfa:ReportingPeriod>\n')
        archivo.write('<sfa:Timestamp>'+str(datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))+'</sfa:Timestamp>\n')
        archivo.write('</ftc:MessageSpec>\n')

        archivo.write('<ftc:FATCA>\n')
        archivo.write('<ftc:ReportingFI>\n')
        archivo.write('<sfa:ResCountryCode>GT</sfa:ResCountryCode>\n')
        archivo.write('<sfa:TIN issuedBy="US">1MMZME.00000.LE.320</sfa:TIN>\n')
        archivo.write('<sfa:Name>BANCO DE LOS TRABAJADORES</sfa:Name>\n')
        archivo.write('<sfa:Address>\n')
        archivo.write('<sfa:CountryCode>GT</sfa:CountryCode>\n')
        archivo.write('<sfa:AddressFree>AVENIDA REFORMA 6 20 ZONA 9 GUATEMALA</sfa:AddressFree>\n')
        archivo.write('</sfa:Address>\n')
        archivo.write('<ftc:FilerCategory>FATCA601</ftc:FilerCategory>\n')
        archivo.write('<ftc:DocSpec>\n')
        archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
        archivo.write('<ftc:DocRefId>1MMZME.00000.LE.320.'+str(uuid.uuid4())+'</ftc:DocRefId>\n')
        archivo.write('</ftc:DocSpec>\n')
        archivo.write('</ftc:ReportingFI>\n')
        archivo.write('<ftc:ReportingGroup>\n')

        # Itera sobre las filas recuperadas y agrega los datos al XML
        for row in rows:
            archivo.write('<ftc:AccountReport>\n')
            archivo.write('<ftc:DocSpec>\n')
            archivo.write('<ftc:DocTypeIndic>FATCA1</ftc:DocTypeIndic>\n')
            archivo.write('<ftc:DocRefId>'+str(row[1])+'</ftc:DocRefId>\n')
            archivo.write('</ftc:DocSpec>\n')
            archivo.write('<ftc:AccountNumber>'+str(row[2])+'</ftc:AccountNumber>\n')
            archivo.write('<ftc:AccountClosed>false</ftc:AccountClosed>\n')
            archivo.write('<ftc:AccountHolder>\n')
            archivo.write('<ftc:Individual>\n')
            archivo.write('<sfa:ResCountryCode xmlns:sfa="urn:oecd:ties:stffatcatypes:v2">'+str(row[4])+'</sfa:ResCountryCode>\n')
            archivo.write('<sfa:TIN issuedBy="US" xmlns:sfa="urn:oecd:ties:stffatcatypes:v2">'+str(row[5])+'</sfa:TIN>\n')
            archivo.write('<sfa:Name xmlns:sfa="urn:oecd:ties:stffatcatypes:v2">\n')
            archivo.write('<sfa:FirstName>'+str(row[7])+'</sfa:FirstName>\n')
            archivo.write('<sfa:MiddleName>'+str(row[8])+'</sfa:MiddleName>\n')
            archivo.write('<sfa:LastName>'+str(row[9])+'</sfa:LastName>\n')
            archivo.write('</sfa:Name>\n')
            archivo.write('<sfa:Address xmlns:sfa="urn:oecd:ties:stffatcatypes:v2">\n')
            archivo.write('<sfa:CountryCode>'+str(row[11])+'</sfa:CountryCode>\n')
            archivo.write('<sfa:AddressFix>\n')
            archivo.write('<sfa:Street>'+str(row[12])+'</sfa:Street>\n')
            archivo.write('<sfa:BuildingIdentifier>'+str(row[13])+'</sfa:BuildingIdentifier>\n')
            archivo.write('<sfa:PostCode>'+str(row[14])+'</sfa:PostCode>\n')
            archivo.write('<sfa:City>'+str(row[15])+'</sfa:City>\n')
            archivo.write('<sfa:CountrySubentity>'+str(row[16])+'</sfa:CountrySubentity>\n')
            archivo.write('</sfa:AddressFix>\n')
            archivo.write('</sfa:Address>\n')
            archivo.write('<sfa:BirthInfo xmlns:sfa="urn:oecd:ties:stffatcatypes:v2">\n')
            archivo.write('<sfa:BirthDate>'+str(row[17])+'</sfa:BirthDate>\n')
            archivo.write('</sfa:BirthInfo>\n')
            archivo.write('</ftc:Individual>\n')
            archivo.write('</ftc:AccountHolder>\n')
            archivo.write('<ftc:AccountBalance currCode="USD">'+str(round(float(row[18]),2))+'</ftc:AccountBalance>\n')
            archivo.write('<ftc:Payment>\n')
            archivo.write('<ftc:Type>FATCA502</ftc:Type>\n')
            archivo.write('<ftc:PaymentAmnt currCode="USD">'+str(round(float(row[19]),2))+'</ftc:PaymentAmnt>\n')
            archivo.write('</ftc:Payment>\n')
            archivo.write('</ftc:AccountReport>\n')
        
    
        archivo.write('</ftc:ReportingGroup>\n')
        archivo.write('</ftc:FATCA>\n')
        archivo.write('</ftc:FATCA_OECD>')
        # Cierra la conexión a la base de datos
        connection.close()
        return archivo
    
def comprimir_carpeta(carpeta, archivo_zip):
    with zipfile.ZipFile(archivo_zip, 'w') as zipf:
        for carpeta_raiz, _, archivos in os.walk(carpeta):
            for archivo in archivos:
                ruta_completa = os.path.join(carpeta_raiz, archivo)
                ruta_relativa = os.path.relpath(ruta_completa, carpeta)
                zipf.write(ruta_completa, ruta_relativa)

# Llama a la función para insertar
#csv_to_sql('DATOS_FATCA', 'C:\\Users\\marlon_jimenez\\Downloads\\FATCA 2023.xlsx')

# Llama a la función para convertir
#sql_to_xml('NBCUMPLI032\MSSQLSERVERLOCAL', 'FATCA', 'DATOS_FATCA', 'archivo_banco.xml')

# Llama a la función de xml para Casa de Bolsa
#xml_casa_bolsa('archivo_casa_de_bolsa.xml')

# Llama a la función de xml para Casa de Bolsa
#xml_financiera('archivo_financiera.xml')
    