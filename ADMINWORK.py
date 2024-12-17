import time
import os
import re
import math
import glob
import requests
import zipfile
import PyPDF2
import openpyxl
import zipfile
import threading
import shutil
import pdfplumber
import numpy as np
import BDD_A as bdd
import pandas as pd
import MODDECON as mdc
from PyPDF2 import PdfWriter, PdfReader
from icecream import ic
from DOWN import INSUMOS
from Multiherramienta import *
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from exchangelib import Credentials, Account, DELEGATE, Configuration
from exchangelib.errors import ErrorItemNotFound
import win32com.client as win32 #pywin32
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (TimeoutException, ElementClickInterceptedException, 
    NoSuchElementException, UnexpectedAlertPresentException, NoSuchFrameException, 
    InvalidArgumentException, StaleElementReferenceException, UnexpectedAlertPresentException)



#------------------------------------------------Programas base de Hoja viajera--------------------

def HV_manager(threads, headless=True):
    ruta_descargas = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\'

    # Función para revisar tabla de hoja viajera
    def check_tabla_hv(driver):
        pagina_seguimiento = "https://hojaviajeradigital.javer.com.mx:9260/#/seguimiento"
        # Visitar página y asegurar estár en página
        while True:
            try:
                time.sleep(1)
                url_actual = driver.current_url

                # Checar que estamos en la url correcta
                if url_actual != pagina_seguimiento:
                    acceso_hv(driver)
                    seguimiento = "//button[@ng-reflect-router-link='/seguimiento']"
                    swdw(driver, 5, 0, seguimiento).click()
                # De estar correcto, rompemos el bucle
                else:
                    break
            except Exception as e:
                print(e)
                continue

        # Filtramos la tabla
        while True:
            solicitante = "(//input[@name='solicitante'])[2]"
            swdw(driver, 5, 0, solicitante).clear()
            before_xpath = "//span[contains(@class,'mat-select-placeholder ng-tns-c110-29')]"
            time.sleep(1)
            swdw(driver, 5, 0, solicitante).send_keys("FPRADO" + Keys.ENTER)
            time.sleep(1)
            fecha_desde = swdw(driver, 1, 0, "//input[@ng-reflect-name='desdeFechaCreacion']")
            fecha_desde.click()
            time.sleep(1)
            fecha_desde.send_keys("01012019")
            time.sleep(1)
            swdw(driver, 2, 1, "mat-select-6").send_keys("200" + Keys.ENTER + Keys.ENTER)
            time.sleep(1)
            break

    # Función para descargar tabla de hoja viajera
    def descarga_tabla_hv(driver):
        # Ingresa en el protal de hoja viajera y descarga todas las tablas, y archivos zip
        while True:
            check_tabla_hv(driver)
            descargar = '//mat-icon[text()="download"]'
            # enviamos datos de solicitante
            try:
                swdw(driver, 2, 0, descargar).click()
                break
            except StaleElementReferenceException:
                continue
        # cerramos el driver y esperamos
        time.sleep(3)

    # Función para encontrar el archivo más reciente
    def buscar_archivo_reciente():
        lista_archivos = glob.glob("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\Reporte_Hojas_Viajeras*.csv")
        if not lista_archivos:
            return None
        archivo_mas_reciente = max(lista_archivos, key=os.path.getctime)
        return archivo_mas_reciente

    # Función para eliminar los archivos antiguos
    def eliminar_archivos_antiguos():
        lista_archivos = glob.glob("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\Reporte_Hojas_Viajeras*.csv")
        if not lista_archivos:
            return
        archivo_mas_reciente = max(lista_archivos, key=os.path.getctime)
        for archivo in lista_archivos:
            if archivo != archivo_mas_reciente:
                os.remove(archivo)
    
    # Descargador de hoja viajera
    def HV_downloader():        
        csv_files = glob.glob(os.path.join(ruta_descargas, 'Reporte_Hojas_Viajeras*.csv'))
        # Obtener la fecha y hora actual
        now = datetime.now()
        # Inicializar la variable para mantener el archivo CSV más reciente
        most_recent_csv = None
        # Bucle para asegurar que exista archivo
        while True:
            archivo_mas_reciente = buscar_archivo_reciente()
            print('archivo_mas_reciente is', archivo_mas_reciente) 
            if archivo_mas_reciente is None:
                driver = create_driver(headless=headless)
                descarga_tabla_hv(driver)
                driver.quit()
                organize_firefox()
            else:
            # Verificar si el archivo más reciente tiene más de 60 minutos
                tiempo_actual = time.time()
                tiempo_archivo_mas_reciente = os.path.getctime(archivo_mas_reciente)
                if (tiempo_actual - tiempo_archivo_mas_reciente) > 120:
                    driver = create_driver(headless=headless)
                    descarga_tabla_hv(driver)
                    eliminar_archivos_antiguos()
                    archivo_mas_reciente = buscar_archivo_reciente()
                    driver.quit()
                    organize_firefox()
                else:
                    eliminar_archivos_antiguos()
                    archivo_mas_reciente = buscar_archivo_reciente()
            if archivo_mas_reciente is not None:
                break
        # Convertimos en dataframe el mas reciente
        hv_df = pd.read_csv(archivo_mas_reciente,
                    header=0,            # Use the first row as header
                    index_col=False,     # Ensure no column is used as the index
                    sep=',',             # Set the delimiter to comma
                    quotechar='"',       # Specify the quote character
                    skipinitialspace=True # Skip any spaces after the delimiter
                    )
        ic(hv_df)
        return hv_df

    def hoja_d(driver):
        # Definiciones xpath
        usuario = JAVER_ID.USER
        password = JAVER_ID.PASSWORD
        seguimiento = "//button[@ng-reflect-router-link='/seguimiento']"

        while True:
            row, df = take_first_row()
            if len(row) == 0:
                break
            hoja = row['Folio']

            check_folder = os.path.exists(os.path.join(ruta_descargas, hoja))
            check_pdf = os.path.exists(os.path.join(ruta_descargas, hoja + ".pdf"))
            check_zip = os.path.exists(os.path.join(ruta_descargas, hoja + ".zip"))
            
            if check_folder and check_zip:
                continue

            try:
                # Nos aseguramos de estár en seguimiento
                check_tabla_hv(driver)
                # Ajusta el zoom al 75% (0.75)
                zoom_percentage = 0.75
                driver.execute_script(f"document.body.style.zoom='{zoom_percentage}'")
                try:
                    time.sleep(1)
                    df = beautiful_table(driver, name="mat-table cdk-table")
                    fila = df.loc[df['Folio'] == hoja]
                    indice = fila.index[0] + 1
                    print(f'N: {indice}, {hoja}| {row['Tipo Hoja']}, {row['Rubro']}, {row['Conjunto']}, {row['Contratista']}')
                    xpath_hv = f"//table[contains(@class,'mat-table cdk-table')]/tbody[1]/tr[{indice}]/td[11]/a[1]/button[1]/span[1]/mat-icon[1]"
                    swdw(driver, 2, 0, xpath_hv).click()
                except IndexError:
                    print('Falló:', hoja)
                    for i, row in df.iterrows():
                        print(i, row['Folio'])
                        continue

                ventanas = driver.window_handles
                driver.switch_to.window(ventanas[-1])

                # Ingresamos a cada hoja viajera
                descarga_pdf = "//mat-icon[text()='picture_as_pdf']"
                descarga_zip = "//mat-icon[text()='download']"

                # Descargamos el zip y el pdf
                print(f'¿{hoja} pdf existe? {check_pdf}')
                if not check_pdf:
                    swdw(driver, 5, 0, descarga_pdf)
                    time.sleep(1)
                    swdw(driver, 5, 0, descarga_pdf).click()
                    swdw(driver, 5, 0, descarga_pdf).click()
                    time.sleep(2)

                # Descargamos el zip y el pdf
                print(f'¿{hoja} zip existe? {check_zip}')
                if not check_zip:
                    swdw(driver, 5, 0, descarga_zip)
                    time.sleep(1)
                    print(f'Try to download {descarga_zip}')
                    swdw(driver, 5, 0, descarga_zip).click()
                    time.sleep(9)

                driver.close()  # Cierra la pestaña actual
                driver.switch_to.window(ventanas[0])

            except (TimeoutException, StaleElementReferenceException) as E:
                print(E)
                time.sleep(0.2)

        driver.quit()
        organize_firefox()

    def HV_compile_folders():
        # Hacemos el os walk
        for root, dirs, files in os.walk(ruta_descargas):
            for file in files:
                if file.upper().startswith('HV') and file.lower().endswith('.zip'):
                    nombre, extension = os.path.splitext(file)
                    nombre_pdf = nombre + '.pdf'
                    try:
                        with zipfile.ZipFile(os.path.join(ruta_descargas, file), 'a') as archivo_zip:
                            archivo_zip.write(os.path.join(ruta_descargas, nombre_pdf), arcname=nombre_pdf)
                            archivo_zip.extractall(ruta_descargas + nombre + "\\")
                    except FileNotFoundError as E:
                        continue
                    # Eliminar pdf
                    os.remove(os.path.join(ruta_descargas, nombre_pdf))

    # ACABAN DEFINICIONES EXTRAS --------------------------------------------------------

    # Obtenemos el dataframe mas reciente de hojas viajeras
    hv_df = HV_downloader()
    hojas_id = hv_df
    # Limpiamos la lista de aquellos elemntos previamente descargados
    folios = os.listdir(ruta_descargas)
    for folio in folios:
        file = os.path.join(ruta_descargas, folio)
        if folio.upper().startswith("HV") and os.path.isdir(file) and os.path.exists(os.path.join(file + ".zip")):
            print('Existe:', folio)
            hojas_id = hojas_id.loc[hojas_id['Folio'] != folio] 
    save_flash_memory(hojas_id)        

    # dividimos la lista en varios threads
    chunk_size = min(len(hojas_id), threads)
    if chunk_size != 0:
        threads = []

        # Ejecutamos cada thread
        for x in range(chunk_size):
            driver = create_driver(driver_type='firefox', headless=headless, download_folder="C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\")
            t = threading.Thread(target=hoja_d, args=(driver, ))
            t.start()
            threads.append(t)
        for t in threads:
            t.join()

    #Compilamos archivos descargados
    HV_compile_folders()


    # Limpiamos carpetas y archivos ya concluidos
    carpetas = os.listdir(ruta_descargas)
    for carpeta in carpetas:
        # Ruta de carpeta
        carpeta_eliminar = os.path.join(ruta_descargas, carpeta)
        # Separamos nombre y extensión
        name, ext = os.path.splitext(carpeta)
        # Revisamos que el nombre del archivo empiece con HV y no esté el nombre en la lista
        if carpeta.upper().startswith("HV") and name not in hv_df['Folio'].values:
            # Si es carpeta la eliminamos con todo lo interno
            if os.path.isdir(carpeta_eliminar):
                try:
                    shutil.rmtree(carpeta_eliminar)
                except Exception as e:
                    print("Error al eliminar carpeta: ", e)
            # Si es archivo borramos el archivo
            elif carpeta.lower().endswith('.zip') or carpeta.lower().endswith('.pdf'):
                print(carpeta, carpeta_eliminar)
                os.remove(carpeta_eliminar)

    print('Éxito')


#-----------------------------------------------------------Cuestiones del IMSS--------------------

def SIROC_update():

    # -----------------------------------------Definiciones------------------------------------------------------------------------------------------
    def extract_bdd_xlsx():

        # Definiciones de rutas
        path_siroc = "C:\\users\\fprado\\REPORTES\\CONCENTRADO_J.xlsx"
        xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"

        # Registros de siroc acutales
        df_actual = pd.read_excel(xlsx_root, sheet_name="QRO")
        df_s_actual = pd.read_excel(xlsx_root, sheet_name="x")
        # Dataframe de todos los contratos generados
        df_origin = pd.read_excel(path_siroc, sheet_name="CON")

        # Arreglos a dataframes
        columnas_eliminar = ["UEN-Plaza", "Tipo de contrato", "Tipo de obra", "Fecha creación", "Fecha terminación",
                            "Registro patronal", "Numero de trabajadores", "Contratante", "Municipio firma", 
                            "Municipio(fracc)", "Tipo documento", "Numero de modificatorio", "Aplica fianza", 
                            "Aplica FG", "Días FG", "%FG", "Monto Anticipo"]
        df_origin = df_origin.drop(columnas_eliminar, axis=1)
        df_origin = df_origin.iloc[:, 1:]
        df_origin = df_origin[["Contratista", "Fraccionamiento", "Conjunto", "Contrato legal", 
                                "Descripción de trabajo", "Importe contrato", "Superficie", 
                                "Fecha inicio", "Usuario creador", 'Estado el contrato actual', 'Aplica SIROC']]

        # Filtrar aprobados
        filtro = df_origin['Estado el contrato actual'] == 'APROBADO'
        df_origin = df_origin.loc[filtro]
        df_origin = df_origin.drop('Estado el contrato actual', axis=1)

        # Separar alica SIROC de posible error
        filtro = df_origin['Aplica SIROC'] == 'NO'
        df_posible_siroc = df_origin.loc[filtro]
        df_origin = df_origin.loc[~filtro]
        df_posible_siroc.drop('Aplica SIROC', axis=1, inplace=True)
        df_origin.drop('Aplica SIROC', axis=1, inplace=True)
        df_origin.drop_duplicates(subset='Contrato legal', inplace=True)

        # De los posibles tomamos todos aquellos con monto 
        filtro = df_posible_siroc['Importe contrato'] > (250000 * 1.16)
        df_posible_siroc = df_posible_siroc.loc[filtro]

        # Utilizar merge para combinar los DataFrames basándose en la columna 'Contrato legal'
        df_merged = pd.merge(df_actual, df_origin[['Contrato legal', 'Contratista']], on='Contrato legal', how='left')
        df_merged.drop(['Contratista_x'], axis=1, inplace=True)
        df_merged.rename(columns={'Contratista_y': 'Contratista'}, inplace=True)
        df_actual = df_merged
        df_actual = df_actual[["Contratista", "Fraccionamiento", "Conjunto", "Contrato legal", 
                            "Descripción de trabajo", "Importe contrato", "Superficie", "Fecha inicio", 
                            "Usuario creador", "Javer Registro", "Fecha AMMA", "Contratista Registro"]]

        # Agregamos nuevas columnas
        df_origin['Javer Registro'] = "FALTA"
        df_origin['Fecha AMMA'] = "PENDIENTE"
        df_origin['Contratista Registro'] = "FALTA"

        # Checamos ya existentes
        df_origin_depurado = df_origin[~df_origin['Contrato legal'].isin(df_actual['Contrato legal'])]
        df_siroc_depurado = df_posible_siroc[~df_posible_siroc['Contrato legal'].isin(df_s_actual['Contrato legal'])]

        # Concatenamos viejos y nuevos
        df_origin = pd.concat([df_actual, df_origin_depurado], ignore_index=True)
        df_origin = df_origin.sort_values(by=['Contratista', 'Fecha inicio'])
        df_siroc = pd.concat([df_s_actual, df_siroc_depurado], ignore_index=True)

        # Guardamos los cambios
        writer = pd.ExcelWriter(xlsx_root, engine='xlsxwriter')
        df_origin.to_excel(writer, sheet_name='QRO', index=False)
        df_posible_siroc.to_excel(writer, sheet_name="x", index=False)
        writer.close()

    def download_pdf_siroc_mail():

        download_path = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"

        # Convertir la lista en dataframe
        xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"
        df_siroc = pd.read_excel(xlsx_root, sheet_name="QRO")
        filtro = df_siroc['Javer Registro'].isna() | (df_siroc['Javer Registro'] == 'FALTA')
        df_falta_s = df_siroc[filtro]
        busqueda = df_falta_s['Contrato legal']

        # Inicializar objeto Outlook
        outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        # Obtener la carpeta de correo
        ìnbox = outlook.GetDefaultFolder(6)
        in_imss = ìnbox.Folders("IMSS")
        in_sirocs = in_imss.Folders("SIROCS")

        #obtener los mensajes
        messages = in_sirocs.items

        for message in messages:
            # Obtener el cuerpo del mensaje
            body_content = message.body
            # Recorrer la serie busqueda
            for folio in busqueda:                
                ultimos_digitos = re.search((r'(?:.*-)(\d+)$'), folio).group(1)            # Verificar si el folio está en el cuerpo del mensaje
                if str(ultimos_digitos) in body_content:
                    # Si el folio se encuentra en el cuerpo del mensaje, descargar el archivo PDF adjunto
                    for attachment in message.Attachments:
                        if attachment.FileName.endswith("pdf"):
                            fecha_envio = message.SentOn.strftime("%d-%m-%Y")
                            attachment.SaveAsFile(download_path + "c_" + folio + "_d_" + fecha_envio +  "_.pdf")
                            print(folio, attachment.FileName)


        print("Éxito descargando mensajes")

    def get_siroc_date():
        # Creamos el dataframe de almacenamiento
        df = pd.DataFrame(columns=['Contrato', 'Fecha'])

        # Convertir la lista en dataframe
        xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"
        df_siroc = pd.read_excel(xlsx_root, sheet_name="QRO")
        df_falta_s = df_siroc[df_siroc['Fecha AMMA'].isna() | (df_siroc['Fecha AMMA'] == '')]
        # filtro = df_siroc['Fecha AMMA'].isnan()
        # df_falta_s = df_siroc[filtro]
        busqueda = df_falta_s['Contrato legal']

        # Inicializar objeto Outlook
        outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        # Obtener la carpeta de correo
        ìnbox = outlook.GetDefaultFolder(6)
        in_imss = ìnbox.Folders("IMSS")
        in_sirocs = in_imss.Folders("SIROCS")

        # obtener los mensajes
        messages = in_sirocs.items
        for message in messages:
            body_content = message.body
            for contrato in busqueda:
                folio = re.search((r'(?:.*-)(\d+)$'), contrato).group(1)
                # Verificar si el folio está en el cuerpo del mensaje
                if str(folio) in body_content and message.Attachments:
                    # Obtener la fecha de envío del contrato
                    fecha_envio = message.SentOn.strftime("%d/%m/%Y")
                    df = pd.concat([df, pd.DataFrame({'Contrato': [contrato], 'Fecha': [fecha_envio]})], ignore_index=True)
        return df

    def get_registro_siroc():

        # Definiciones iniciales
        download_path = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"
        df_registro_obra = pd.DataFrame(columns = ['Contrato', 'Registro', 'Fecha'])

        # Paso archivo por archivo
        for root, dirs, files in os.walk(download_path):
            for file in files:
                if file.lower().startswith('c_') and file.lower().endswith('.pdf'):

                    try:
                        # Lector de pdfs
                        pdf_reader = PyPDF2.PdfReader(os.path.join(root, file))
                        page = pdf_reader.pages[0]
                        text = page.extract_text()
                        content = text
                        # Extracción de datos con regular expression
                        contrato = re.search(r'c_(.*?)_', file).group(1)
                        fecha = re.search(r'd_(.*?)_', file).group(1)

                        try:
                            resultado = re.search(r'registro de obra:([A-Za-z]\d{7})\|', content, re.DOTALL).group(1)
                        except AttributeError:
                            resultado = re.search(r'registro de obra.*?([A-Za-z]\d{7})', content, re.DOTALL).group(1)
                        print(contrato, fecha, resultado)
                        # Almacenamiento en dataframe
                        data = {'Contrato': contrato, 'Registro': resultado, 'Fecha': fecha}
                        new_data = pd.DataFrame([data])
                        df_registro_obra = pd.concat([df_registro_obra, new_data])

                    except AttributeError as E:
                        print(E)
                        print("------------------E R R O R------------------------")
                        print(content)
                        print("---------------------------------------------------")
                        print("")
                        print("")
                        continue

        # Función de limpieza para aplicar a cada valor de la columna
        def limpiar_valor(valor):
            if isinstance(valor, str):
                valor = valor.strip()  # Eliminar caracteres no deseados al principio y al final
            return valor

        # Aplicar la función de limpieza a cada columna del DataFrame
        df_registro_obra = df_registro_obra.applymap(limpiar_valor)
        return df_registro_obra


    def write_registro_obra():
        # Definiciones iniciales
        xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"
        df_registro_obra = get_registro_siroc() # Dataframe de registro de obra
        df_bdd = pd.read_excel(xlsx_root, sheet_name="QRO") # Libro de excel con información
        df_x = pd.read_excel(xlsx_root, sheet_name="x") 
        print(df_registro_obra)

        # Combinar los DataFrames en función de las columnas "Contrato legal" y "Contrato"
        df_merged = pd.merge(df_bdd, df_registro_obra, left_on='Contrato legal', right_on='Contrato', how='left')
        # Actualizar los valores de la columna "Javer Registro" con los valores de la columna "Registro"
        df_merged['Javer Registro'] = df_merged['Registro'].fillna(df_merged['Javer Registro'])
        df_merged['Fecha AMMA'] = df_merged['Fecha'].fillna(df_merged['Fecha AMMA'])

        # Eliminar las columnas innecesarias
        df_merged = df_merged.drop(['Contrato', 'Registro', 'Fecha'], axis=1)
        # Agregamos el calculo de fecha
        df_merged['Fecha AMMA'] = df_merged['Fecha AMMA'].str.replace('/', '-')
        df_merged['Fecha AMMA'] = pd.to_datetime(df_merged['Fecha AMMA'], format='%d-%m-%Y', errors='ignore')

        def asignar_dias_pasados(row):
            if row['Contratista Registro'] == 'FALTA' and row['Javer Registro'] != 'FALTA':
                return (datetime.now() - pd.to_datetime(row['Fecha AMMA'], format='%d-%m-%Y')).days
            else:
                return 'En cumplimiento'

        # Aplicar la función personalizada a cada fila usando apply()
        df_merged['Dias pasados'] = df_merged.apply(asignar_dias_pasados, axis=1)
        reorder = ['Contratista', 'Contrato legal', 'Descripción de trabajo', 'Importe contrato', 'Superficie','Javer Registro', 'Fecha AMMA', 'Contratista Registro', 'Dias pasados', 'Usuario creador', 'Fraccionamiento', 'Conjunto', 'Fecha inicio']
        df_merged = df_merged.reindex(columns=reorder)

        # Guardar todo
        writer = pd.ExcelWriter(xlsx_root, engine='xlsxwriter')
        df_merged.to_excel(writer, sheet_name='QRO', index=False)
        df_x.to_excel(writer, sheet_name="x", index=False)
        writer.close()
        print("Éxito escribiendo valores en Registro")

    # Creación de respaldo y lectura de archivo pasado
    xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"
    fecha_hoy = datetime.now()
    fecha_hoy = fecha_hoy.strftime("%m%d%H%M%S")
    fecha_hoy = int(fecha_hoy)
    fecha_hoy = hex(fecha_hoy)
    fecha_hoy = fecha_hoy.lstrip("0x")
    respaldo_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_r_" + fecha_hoy + ".xlsx"
    
    archivos_cumplen_condicion = []
    for root, dirs, files in os.walk("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\"):
        for file in files:
            if file.startswith('N_SIROC_r'):
                ruta_completa = os.path.join(root, file)
                archivos_cumplen_condicion.append(ruta_completa)

    # Ordenar los archivos por fecha de modificación, separación de 3 mas recientes y archivos a eliminar
    archivos_cumplen_condicion.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    archivos_recientes = archivos_cumplen_condicion[:3]
    archivos_no_recientes = [archivo for archivo in archivos_cumplen_condicion if archivo not in archivos_recientes]

    # Imprimir los archivos conservados
    for archivo in archivos_no_recientes:
        print(archivo)
        os.remove(archivo)

    # Definiciones de dataframes
    df_actual = pd.read_excel(xlsx_root, sheet_name="QRO")
    df_s_actual = pd.read_excel(xlsx_root, sheet_name="x")
    
    # Creamos respaldo, por si las moscas
    writer = pd.ExcelWriter(respaldo_root, engine='xlsxwriter')
    df_actual.to_excel(writer, sheet_name='QRO', index=False)
    df_s_actual.to_excel(writer, sheet_name="x", index=False)
    writer.close()

    # Aplicaciones
    extract_bdd_xlsx()
    download_pdf_siroc_mail()
    write_registro_obra()

    xlsx_root = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\N_SIROC_2023.xlsx"
    df_bdd = pd.read_excel(xlsx_root, sheet_name="QRO") # Libro de excel con información
    df_x = pd.read_excel(xlsx_root, sheet_name="x")
    df_dates = get_siroc_date()

    # Combinar los DataFrames en función de las columnas "Contrato legal" y "Contrato"
    df_merged = pd.merge(df_bdd, df_dates, left_on='Contrato legal', right_on='Contrato', how='left')
    df_merged['Fecha AMMA'] = df_merged['Fecha AMMA'].fillna(df_merged['Fecha'])

    # Eliminar las columnas innecesarias
    df_merged = df_merged.drop(['Contrato', 'Fecha'], axis=1)
    
    # Agregamos el calculo de fecha
    df_merged['Fecha AMMA'] = df_merged['Fecha AMMA'].str.replace('/', '-')
    df_merged['Fecha AMMA'] = pd.to_datetime(df_merged['Fecha AMMA'], format='%d-%m-%Y', errors='ignore')

    def asignar_dias_pasados(row):
        if row['Contratista Registro'] == 'FALTA' and row['Javer Registro'] != 'FALTA':
            return (datetime.now() - pd.to_datetime(row['Fecha AMMA'], format='%d-%m-%Y')).days
        elif row['Javer Registro'] == 'FALTA':
            return 'Pendiente'
        else:
            return 'Recibido'

    # Aplicar la función personalizada a cada fila usando apply()
    df_merged['Dias pasados'] = df_merged.apply(asignar_dias_pasados, axis=1)

    # Creamos respaldo, por si las moscas
    writer = pd.ExcelWriter(xlsx_root, engine='xlsxwriter')
    df_merged.to_excel(writer, sheet_name='QRO', index=False)
    df_s_actual.to_excel(writer, sheet_name="x", index=False)
    writer.close()

    download_path = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"

    # Paso archivo por archivo
    for root, dirs, files in os.walk(download_path):
        for file in files:
            if file.lower().startswith('c_') and file.lower().endswith('.pdf'):
                archivo = os.path.join(root, file)
                os.remove(archivo)


def viejo_finiquitador(ubic, mod=True, fin=True, abandono=False):
    
    def select_document(elemento):
        while True:
            try:
                select = Select(select_element)
                select.select_by_visible_text(elemento)
                break
            except StaleElementReferenceException:
                swdw(driver, 3, 1, 'typeFile').click()
                swdw(driver, 3, 1, 'typeFile').send_keys(elemento + Keys.ENTER)
                break
            except TimeoutException:
                continue

    driver = create_driver(driver_type='firefox', headless=False, download_folder="C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\")
    
    # Dataframe y lista de contratos
    df = pd.read_excel(ubic)
    contratos = df['Contrato']
    errores = []

    # Iteración de contratos
    for contrato in contratos:

        try:
            # Obtenemos el valor ID del dataframe
            c_id = df.loc[df['Contrato'] == contrato, 'conjunto_ID'].values[0]

            # Partes del hipervínculo

            link = f"https://{JAVER_ID.USER}:{JAVER_ID.PASSWORD}@portal.javer.net/juridico/Paginas/Biblioteca.aspx?ContratoID={str(c_id)}"
            
            # Ingresar a página y verificación de información

            while True:
                try:
                    ic(link, contrato, c_id)        
                    # Acceso a la página
                    driver.get(link)
                    check = swdw(driver, 5, 1, "titleContrato").text

                except StaleElementReferenceException:
                    time.sleep(1)

                except TimeoutException:
                    continue

                # Revisamos compatibilidad
                if check == contrato:
                    break

            # Definimos selector de nuevos archivos
            select_element = driver.find_element(By.ID, "typeFile")
            # Obtenemos la fecha y monto del dataframe
            fecha = df.loc[df['Contrato'] == contrato, 'Fecha'].values[0]
            monto = df.loc[df['Contrato'] == contrato, 'Total_estimado'].values[0]

            # Extraer los existentes
            table = mdc.beautiful_table(driver, element="id", name="contentDS")
            # Eliminar caracteres no numéricos excepto puntos decimales y comas
            table['Monto'] = table['Monto'].str.replace('[^0-9.,]', '', regex=True)
            
            # table['Monto'] = table['Monto'].str.replace('[^\d.,]', '', regex=True)

            # Reemplazar comas por puntos
            table['Monto'] = table['Monto'].str.replace(',', '')
            # Convertir la columna en números
            table['Monto'] = table['Monto'].astype(float)
            # Obtenemos finiquitos y modificatorios existentes        
            mods = table[table['Archivo'].str.contains('CM')]
            fins = table[table['Archivo'].str.contains('CF')]

            # Atender las necesidades del objeto mod, fin, abandono
            hay_mod, hay_fin = False, False

            if not mods.empty:
                if monto in mods['Monto'].values:
                    hay_mod = True
                    print('ya hay mod')

            if not fins.empty:
                hay_fin = True
                print('ya hay fin')

            # Si no hay nada que hacer salimos
            if hay_fin and not abandono:
                df = df.drop(df[df['Contrato'] == contrato].index)
                df.to_excel(ubic, index=False)
                continue

            # Modificatorios ------------------------------------
            if mod and not hay_fin and not hay_mod:

                select_document('Modificatorio')
                is_fecha = swdw(driver, 4, 1, "isfecha")
                valid = pd.to_datetime(fecha, errors='coerce')
                validator = not pd.isna(valid)
                ic(valid, validator)

                # Cambiamos la fecha de modificatorio
                if fecha is not None and validator:

                    # INGRESA LOS DATOS DEL MODIFICATORIO O REALIZAR
                    MES = ['Nel', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
                    try: 
                        DIG1 = int(fecha.strftime("%m"))
                        DIA = str(int(fecha.strftime("%d")))
                        ANNO = fecha.strftime("%Y")
                    except (AttributeError, ValueError):
                        try:
                            # SI LA FECHA VIENE EN FORMATO TEXTO SE CAMBIA A FORMATO FECHA
                            fecha = str(fecha)
                            fecha = fecha[:10]
                            ic(fecha)
                            fecha = datetime.strptime(fecha, '%d/%m/%Y')
                            DIG1 = int(fecha.strftime("%m"))
                            DIA = str(int(fecha.strftime("%d")))
                            ANNO = fecha.strftime("%Y")
                        except (AttributeError, ValueError):
                            fecha = str(fecha)
                            fecha = fecha[:10]
                            ic(fecha)
                            fecha = datetime.strptime(fecha, '%Y-%m-%d')
                            DIG1 = int(fecha.strftime("%m"))
                            DIA = str(int(fecha.strftime("%d")))
                            ANNO = fecha.strftime("%Y")
                    
                    while True:
                        try:
                            # Damos click a la fecha
                            is_fecha.click()
                            break
                        except (ElementClickInterceptedException, StaleElementReferenceException):
                            continue
                    
                    # esperamos carga del calendario y le damos click
                    calendar = swdw(driver, 2, 0, "//img[contains(@class, 'ui-datepicker-trigger')]")
                    calendar.click()

                    # esperamos carga de botón de mes y seleccionamos el mes
                    swdw(driver, 2, 0, "//select[@class='ui-datepicker-month']")

                    MONTH = Select(driver.find_element(By.XPATH, "//select[@class='ui-datepicker-month']"))
                    MONTH.select_by_visible_text(MES[DIG1])

                    # esperamos carga de botón de año y seleccionamos año
                    swdw(driver, 2, 0, "//select[@class='ui-datepicker-year']")
                    YEAR = Select(driver.find_element(By.XPATH, "//select[@class='ui-datepicker-year']"))
                    YEAR.select_by_visible_text(ANNO)

                    # damos click al día
                    boton_dia = '//a[text()="' + DIA + '"]'
                    swdw(driver, 5, 0, boton_dia).click()

                # Cambiamos el monto en el modificatorio
                if monto is not None:
                    try:
                        swdw(driver, 5, 1, 'ismonto').click()
                    except ElementClickInterceptedException:
                        time.sleep(1)
                        WebDriverWait(driver, 10).until(EC.invisibility_of_element((By.CLASS_NAME, 'loading')))
                        swdw(driver, 5, 1, 'ismonto').click()

                    monto_input = swdw(driver, 5, 1, 'montoIG').send_keys(str(monto) + Keys.TAB)

                # Obtenemos la fecha de último modificatorio
                fecha_str = swdw(driver, 5, 0, "//span[@class='UltimoModificatorio']").text
                fecha_firma = datetime.strptime(fecha_str, "%d/%m/%Y")

                # Obtenemos la fecha promedio
                fecha_promedio = fecha_firma - timedelta(days=50)

                # Convertimos y enviamos fecha de firma
                fecha_para_firma = fecha_promedio.strftime("%d/%m/%Y")
                swdw(driver, 10, 1, 'firmaCM').send_keys(fecha_para_firma)

                while True:
                    try:
                        # Terminamos saliendo y regresando
                        try: 
                            swdw(driver, 10, 1, 'sig').click()
                            swdw(driver, 10, 1, 'sig').click()
                            swdw(driver, 10, 1, 'acc').click()
                            time.sleep(5)
                            swdw(driver, 10, 1, 'close')
                        except UnexpectedAlertPresentException:
                            time.sleep(5)
                            driver.refresh()

                        while True:
                            try:
                                swdw(driver, 1, 1, "titleContrato")
                                break
                            except TimeoutException:
                                swdw(driver, 10, 0, "//h3[text()='X']").click()
                        break
                    except ElementClickInterceptedException:
                        while True:
                            try:
                                swdw(driver, 1, 1, "message")
                                swdw(driver, 10, 0, "//h3[text()='X']").click()
                            except TimeoutException:
                                swdw(driver, 1, 1, 'firmaCM')
                                break
                        fecha_promedio = fecha_promedio - timedelta(days=20)
                        fecha_para_firma = fecha_promedio.strftime("%d/%m/%Y")
                        print(fecha_para_firma)
                        swdw(driver, 10, 1, 'firmaCM').clear()
                        swdw(driver, 10, 1, 'firmaCM').send_keys(fecha_para_firma + Keys.TAB)
                        time.sleep(1)
                        continue


            # Finiquitos ----------------------------------------
            if fin and not hay_fin:

                select_document('Finiquito')

                # damos avanzar hasta regresar al contrato
                while True:
                    try:
                        swdw(driver, 10, 1, 'sig').click()
                        break
                    except (ElementClickInterceptedException, StaleElementReferenceException):
                        continue
                swdw(driver, 10, 1, 'acc').click()
                swdw(driver, 10, 1, 'close')

                while True:
                    try:
                        swdw(driver, 1, 1, "titleContrato")
                        break
                    except TimeoutException:
                        swdw(driver, 10, 0, "//h3[text()='X']").click()

            # Abandono de obra ---------------------------------
            if abandono:
                # Seleccionamos el abandono de obra
                select_document('Abandono de obra')

                # Damos click al botón de abandono con obra ejecutada
                while True:
                    try:
                        boton_abandono = "//label[text()='Acta de Abandono Con Obra Ejecutada']"
                        swdw(driver, 10, 0, boton_abandono).click()
                        swdw(driver, 5, 1, "txtPresupuestoEjercido").send_keys(monto_mod)
                        time.sleep(0.5)
                        swdw(driver, 5, 1, "txtMotivoAbandono").send_keys("Contratista no firmó los convenios finiquito")
                        time.sleep(0.5)
                        swdw(driver, 5, 1, 'txtRetenido').send_keys(monto_fg)
                        time.sleep(0.5)

                        # Adjuntamos el soporte de correo txtPresupuestoEjercido
                        ruta_archivos = "C:\\Users\\fprado\\REPORTES\\Abandono\\Correos\\"
                        contratista = df.loc[df['Contratos'] == contrato, 'Contratista'].values[0]
                        hypervinculo_archivo = ruta_archivos + contratista + ".msg"
                        archivo = swdw(driver, 10, 1, "fileCorreoF")
                        archivo.send_keys(hypervinculo_archivo)

                        # Adjuntamos el soporte de estado de cuenta txtPresupuestoEjercido
                        ruta_archivos = "C:\\Users\\fprado\\REPORTES\\Abandono\\Compilados\\"
                        hypervinculo_archivo = ruta_archivos + contratista + " - " + contrato + ".pdf"
                        archivo = swdw(driver, 10, 1, "fileEstadoCuentaF")
                        archivo.send_keys(hypervinculo_archivo)
                        time.sleep(0.2)
                    except (ElementClickInterceptedException, StaleElementReferenceException):
                        continue
                    except UnexpectedAlertPresentException:
                        print("ya con carta de abandono")
                        break
                    except InvalidArgumentException:
                        ram_mem = []
                        print(contratista, contrato)
                        ram_mem.append(contratista)
                        ram_mem.append(contrato)
                        errores.append(ram_mem)
                        break
                    try:
                        swdw(driver, 5, 1, "btnEnviar")
                        # Encontrar el elemento que deseas hacer clic (supongamos que tiene el ID "elemento-a-hacer-clic")
                        elemento = driver.find_element(By.ID, 'btnEnviar')
                        # Desplazarse hasta el elemento utilizando JavaScript
                        driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
                        time.sleep(0.5)                   
                        swdw(driver, 5, 1, "btnEnviar").send_keys(Keys.ENTER)
                    except(TimeoutException, ElementClickInterceptedException, StaleElementReferenceException):
                        break
                    break
                    driver.refresh()
                df_errores = pd.DataFrame(errores, columns=[contratista, contrato])
                df_errores.to_excel('C:\\Users\\fprado\\REPORTES\\Abandono\\df_errores.xlsx')

        except TimeoutException:
            continue
    driver.quit()
    organize_firefox()


def juridico_downloader(root_in, root_out):

    # se convierten en data frame los archivos
    lista_por_finiquitar = pd.read_excel(root_in)

    if os.exists(root_out):
        lista_imprimidos = pd.read_excel(root_out)
    else:
        lista_imprimidos = pd.DataFrame(columns=['Contrato', 'Tipo', 'Contratista', 'Ruta'])
    
    # tomamos cada finiquito de la lista finiquitar
    print(lista_por_finiquitar.columns)
    for finiquito, contratista in lista_por_finiquitar.values:
        
        elementos = []

        # buscamos el finquito dentro de la lista reducida de existentes
        modificatorios = bdd_mod[bdd_mod['contrato'] == finiquito]['título']
        elementos.append(finiquito + '_cf')

        for modificatorio in modificatorios:
            elementos.append(modificatorio)

        for elemento in elementos:
            driver.get("http://fprado@javer.com.mx:" + ju_password + "@portal.javer.net/juridico/contratos_juridico/" + finiquito + "/" + elemento + ".aspx")
            try:
                WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//img[@src='/_layouts/15/javercontratos/img/logojaver_2020.png']")))
            except TimeoutException:
                try:
                    WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//img[@src='/_layouts/15/javercontratos/img/logo_javer.png']")))
                except TimeoutException:
                    print(contratista, elemento, " -no existe en sistema")
                    no_existe_sistema.append("http://fprado@javer.com.mx:" + ju_password + "@portal.javer.net/juridico/contratos_juridico/" + finiquito + "/" + elemento + ".aspx")
                    break

            pyautogui.moveTo(683, 384)
            pyautogui.click()
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'p')
            time.sleep(0.5)
            pyautogui.press('g')
            time.sleep(0.1)
            pyautogui.press('g')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0,5)
            pyautogui.write(contratista + " - " + elemento, interval=0.0001)
            print(contratista + "-" + elemento)
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(0.40)
            pyautogui.press('enter')
            time.sleep(0.40)
            pyautogui.press('enter')
            time.sleep(0.40)
            pyautogui.press('enter')
    print('Se acabó, estos no existen:')
    print(no_existe_sistema)


def organizar_estados_cuenta():
    # Rutas de folderes
    download_path = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"
    root_place = "C:\\Users\\fprado\\REPORTES\\Abandono\\Estado_cuenta\\"

    # Oswalk para cambiar nombre a los ya descargados
    for root, dirs, files in os.walk(download_path):
        for file in files:
            if file.lower().startswith('xxmcan') and file.lower().endswith('.pdf'):

                file_path = os.path.join(root, file)
                # Abrimos el pdf y extraemos el contrato
                with open(file_path, 'rb') as file:
                    reader = PyPDF2.PdfFileReader(file)
                    first_page = reader.pages[0]  # Obtener la primera página
                    pdf_text = first_page.extractText()

                # Buscar el texto entre "Contrato " y " - "
                match = re.search(r'Contrato (.+?) -', pdf_text)
                if match:
                    try:
                        #renombramos los archivos
                        contrato_text = str(match.group(1)) + ".pdf"
                        new_name = os.path.join(root_place, contrato_text)
                        os.rename(file_path, new_name)
                    # omitimos duplicados
                    except FileExistsError:
                        time.sleep(0.1)


def compilar_estados_cuenta():
    root_place = "C:\\Users\\fprado\\REPORTES\\Abandono\\Estado_cuenta\\"

    # Tomamos dataframe de referencia para archivos
    df = pd.read_excel("C:\\Users\\fprado\\REPORTES\\estado.xlsx", dtype={'Frentes': object})
    contratos = df['Contrato']
    file_reports = []
    file_headers = ['Contratista', 'Contrato', 'Root']

    # Buscamos todos los archivos
    for root, dirs, files in os.walk(root_place):
        for file in files:

            # Los archivos pdfs existentes y dentro de la lista de contratos
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                file_name, file_extension = os.path.splitext(file)
                if file_name in contratos.values:
                    data_file = []

                    # Obtenemos cada dato del dataframe basado en el contrato 
                    contratista =  df.loc[df['Contrato_x'] == file_name, 'Contratista_x'].values[0]
                    legal =  df.loc[df['Contrato_x'] == file_name, 'Contrato Legal'].values[0]

                    # Guardamos los datos en una lista
                    data_file.append(contratista)
                    data_file.append(legal)
                    data_file.append(file_path)                    

            # Guardamos la lista en otra lista
            file_reports.append(data_file)
    file_df = pd.DataFrame(file_reports, columns=file_headers)

    # sacamos de la lista creada los contratistas y contratos coincidentes
    contratistas = file_df['Contratista'].drop_duplicates()
    for contratista in contratistas:
        contratos = file_df[file_df['Contratista'] ==  contratista]['Contrato'].drop_duplicates()
        for contrato in contratos:
            
            # Obtenemos los archivos corespondientes a cada contrato
            files = file_df[file_df['Contrato'] == contrato]['Root']
            print(files)
            output = "C:\\Users\\fprado\\REPORTES\\Abandono\\Compilados\\"

            # Si hay mas de un archivo por contrato los mezclamos
            if len(files) > 1:
                merger = PyPDF2.PdfFileMerger()
                filename = f"{contratista} - {contrato}.pdf"
                output_filename = output + filename
                for file in files:
                    merger.append(file)
                merger.write(output_filename)
                merger.close()

            # En caso de solo ser uno lo guardamos
            else:
                try:
                    old_filename = files.iloc[0]
                    filename = f"{contratista} - {contrato}.pdf"
                    new_filename = output + filename
                    os.rename(old_filename,  new_filename)
                except FileExistsError:
                    continue


def busca_correo_archivos(file_in, file_out, extra_word):

    # Dataframe desde file_in excel
    df = pd.read_excel(file_in, dtype={'Frentes': object})
    contratos = df['ID'].drop_duplicates()

    # Inicializar objeto Outlook, in a carpeta SCA y obtener los mensajes
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.GetDefaultFolder(6)
    sca = inbox.Folders('CAO').Folders("SCA")
    messages = sca.items

    for mail in messages:
        subject = mail.Subject
        attachments = mail.Attachments
        n_subject = subject.lower()
        
        # revisamos que tenga escaneo
        if attachments.Count > 0:
            for attachment in attachments:

                # Le quitamos la extension al nombre
                c_adjunto = attachment.FileName.lower()
                match = re.search(r'^(.*)\.*$', c_adjunto)
                if match:
                    n_adjunto = match.group(1)
                else:
                    n_adjunto = c_adjunto

                # revisamos folio por folio que existan en escaneo
                for n_folio in contratos:
                    folio = str(n_folio)
                    if (folio in n_subject or folio in n_adjunto) and (extra_word in n_subject or extra_word in n_adjunto):
                        attachment.SaveAsFile(file_out + n_adjunto)


def multithread_executor(df, application, var, headless=False, threads=2):

    df.sort_values(by='Conjunto_x')
    chunks = np.array_split(df, threads)
    print(chunks)
    threads = []

    for chunk in chunks:
        driver = create_driver(driver_type='firefox', headless=headless, download_folder="C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\")
        t = threading.Thread(target=application, args=(driver, chunk, var,))
        t.start()
        threads.append(t)

    for t in threads:
        t.join()


def df_sint():
    df = pd.read_excel("C:\\Users\\fprado\\REPORTES\\Acta de abandono.xlsx", dtype={'Frente': object})
    descargados = []
    for root, dirs, files in os.walk("C:\\Users\\fprado\\REPORTES\\Abandono\\Estado_cuenta"):
        for file in files:
            file_name, file_extension = os.path.splitext(file)
            descargados.append(file_name)
    filtro = ~df['Contrato_x'].isin(descargados)
    df = df[filtro]
    return(df)


def fecha_ultima_estimacion(threads, root_in, root_out, headless=True):

    # Revisamo existencia de la salida de datos
    flash_root = get_here(name="flash_memory.xlsx")

    def update_flash_memory():
        lock.acquire()
        # Obtenemos el archivo contenedor de finiquitos
        para_finiquitar = pd.read_excel(root_in)
        df_out = pd.read_excel(root_out)

        # Revisamos que datos tienen ya fecha obtenida
        contratos_existentes = df_out['Contrato'].unique()
        existen = para_finiquitar['Contrato'].isin(contratos_existentes)
        contratos_sin_fecha = para_finiquitar[~existen]

        # Filtra las filas donde Total_estimado es diferente de 0 y reasigna el DataFrame
        contratos_sin_fecha = contratos_sin_fecha.loc[contratos_sin_fecha['Total_estimado'] != 0]
        contratos_sin_fecha['Contrato'].drop_duplicates(inplace=True)

        for i, row in contratos_sin_fecha.iterrows():
            print(row['Contrato'])

        contratos_sin_fecha.to_excel(flash_root, index=False)
        lock.release()

    def get_first_and_save():
        lock.acquire()
        memory_root = get_here(name="flash_memory.xlsx")
        df = pd.read_excel(memory_root)
        try:
            row = df.iloc[0]
            len_df = len(df)
            df.drop(df.index[0], inplace=True)
        except IndexError:
            row = []
            len_df = 0
        df.to_excel(memory_root, index=False)
        lock.release()
        return row, len_df

    def ir_reporte_estimación(driver):
        while True:
            try:
                checker = '//h2[@class="x7c" and text()="Reporte de estimacion para contratista"]'
                swdw(driver, 1, 0, checker)
                break
            except TimeoutException:
                go_sistema_reportes(driver)
                busca_reporte_sistema_de_reportes(driver, "Reporte de estimacion para contratista", "colSwitcherExecute")

    # Programa para multithreading:
    def busca_ultima_fecha(in_df, out, driver):
        acceder_oracle(driver)
        barrier.wait()

        while True:
            # Iteramos contrato por contrato
            row, len_df = get_first_and_save()
            if len_df == 0:
                print(0)
                update_flash_memory()
                row, len_df = get_first_and_save()
                if len_df == 0:
                    break

            ir_reporte_estimación(driver)

            # Definiciones del dataframe
            data = []
            contrato = row['Contrato']
            contratista = row['Contratista']
            conjunto = row['Conjunto']
            org = row['Organización']
            frente = conjunto[4:6]
            print(contrato, contratista, conjunto)

            # Definiciones XPATHS de menú
            x_proyecto = "//span[@id='PROJECT_ID__xc_0']/a[1]"
            x_frente = "//span[@id='FRONT_ID__xc_0']/a[1]"
            x_conjunto = "//span[@id='FRONT_BUILD_SET_ID__xc_0']/a[1]"
            x_proveedor = "//span[@id='VENDOR_ID__xc_0']/a[1]"
            x_orden = "//span[@id='CONTRACT_OP_HDR_ID__xc_0']/a[1]"
            
            # Vaciado de datos en formulario
            stf(driver, 0, x_proyecto, org)
            stf(driver, 0, x_frente, frente)
            stf(driver, 0, x_conjunto, conjunto)
            stf(driver, 0, x_proveedor, contratista)
            time.sleep(0.5)

            # Hacemos click en el elemento que querémos abrir en frame
            try:
                swdw(driver, 4, 0, x_orden).click()
            except ElementClickInterceptedException:
                driver.switch_to.window(driver.window_handles[0])
                driver.switch_to.default_content()
                swdw(driver, 4, 0, x_orden).click()

            while True:
                try:
                    # Esperamos a dos ventanas
                    while True:
                        ventanas = driver.window_handles
                        cantidad_ventanas = len(ventanas)
                        if cantidad_ventanas == 2:
                            break
                        else:
                            time.sleep(0.5)
                            continue

                    # Esperamos a frame
                    while True:
                        try:
                            driver.switch_to.window(driver.window_handles[1])
                            driver.switch_to.frame(0)
                            break
                        except NoSuchFrameException:
                            time.sleep(0.5)
                            continue

                    swdw(driver, 5, 0, "//input[@title='Término de Búsqueda']").clear()
                    swdw(driver, 2, 0, "//input[@title='Término de Búsqueda']").send_keys("%" + Keys.TAB + Keys.ENTER)
                    fecha = swdw(driver, 3, 1, "N1:displayColumn4:0").text
                    driver.close()
                    break
                except StaleElementReferenceException:
                    continue
                except TimeoutException:
                    fecha = "No hay fecha"
                    driver.close()
                    break

            # Salimos del frame
            driver.switch_to.window(driver.window_handles[0])
            driver.switch_to.default_content()

            # Guardar
            row_df = []
            data.append(contrato)
            data.append(fecha)
            row_df.append(data)

            # Sección crítica: Guardar en el archivo de Excel
            lock.acquire()
            bdd = pd.read_excel(out)
            df_out = pd.DataFrame(row_df, columns=['Contrato', 'Fecha'])
            output = pd.concat([bdd, df_out])
            output.to_excel(out, index=False)
            lock.release()

        driver.quit()
        organize_firefox()

    # Código fuera de la definición ------------------------


    if os.path.exists(root_out):
        df_out = pd.read_excel(root_out)
    else:
        df_out = pd.DataFrame(columns=['Contrato', 'Fecha'])
        df_out.to_excel(root_out, index=False)

    # Crear un objeto Lock para la sección crítica
    lock = threading.Lock()
    update_flash_memory()
    df_in = pd.read_excel(flash_root)

    # Si el largo es diferente
    if len(df_in) < threads:
        threads = len(df_in)
    if len(df_in) != 0:
        chunks = np.array_split(df_in, threads)
        barrier = threading.Barrier(threads)
        thread_list = []

        # Crear y ejecutar los hilos
        for chunk in chunks:
            driver = create_driver(headless=headless)
            thread = threading.Thread(target=busca_ultima_fecha, args=(chunk, root_out, driver))
            thread.start()
            thread_list.append(thread)

        # Esperar a que todos los hilos terminen
        for thread in thread_list:
            thread.join()




def merger_dates(root_in):
    # Definiciones de bases de datos y dataframes
    root_df_fechas = 'C:\\Users\\fprado\\REPORTES\\merg_finiquitos.xlsx'
    bdd_fechas = pd.read_excel(root_df_fechas)
    bdd_contratos_id = pd.read_excel("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\EJJ.xlsm", sheet_name="C", usecols=['Contrato', 'conjunto_ID'])
    bdd_contratos_fecha_inicial = pd.read_excel("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\EJJ.xlsm", sheet_name="N", usecols=['Contrato', 'Fecha inicio'])
    df_finiquitos = pd.read_excel(root_in)

    # filtramos aquellos contratos que no fueron estimados, los dejamos aparte
    filtro_sin_estimar = df_finiquitos.loc[df_finiquitos['Total_estimado'] < 10]
    contratos_sin_estimar = filtro_sin_estimar['Contrato']
    fecha_contrato_sin_estimar = pd.merge(contratos_sin_estimar, bdd_contratos_fecha_inicial, how='left', on='Contrato')
    fecha_contrato_sin_estimar['Fecha inicio'] = pd.to_datetime(fecha_contrato_sin_estimar['Fecha inicio'])
    fecha_contrato_sin_estimar['Fecha inicio'] = fecha_contrato_sin_estimar['Fecha inicio'] + timedelta(days=1)
    fecha_contrato_sin_estimar.rename(columns={'Fecha inicio': 'Fecha'}, inplace=True)

    # Agregamos los datos requeridos
    df_con_fechas = pd.merge(df_finiquitos, bdd_fechas, how='left', on='Contrato')
    df_con_id = pd.merge(df_con_fechas, bdd_contratos_id, how='left', on='Contrato')
    
    # Filtramos lo contratos que ya tienen fecha
    filtro_con_estimacion = df_con_id.loc[df_con_id['Total_estimado'] != 0]
    columnas_deseadas = ['Contrato', 'Fecha']
    fecha_contrato_con_estimacion = filtro_con_estimacion[columnas_deseadas]
    fechas_de_contratos = pd.concat([fecha_contrato_con_estimacion, fecha_contrato_sin_estimar], ignore_index=True)

    #  Eliminamos la columna fecha para hacer merge con todas las fechas
    df_con_id = df_con_id.drop(['Fecha'], axis=1)
    df_final = pd.merge(df_con_id, fechas_de_contratos, how='left', on='Contrato')

    df_final.drop_duplicates(inplace=True)
    df_final.to_excel("C:\\Users\\fprado\\REPORTES\\finiquitador.xlsx", index=False)


def archivo_finiquitador(file):
    # Compilamos la fuente para crear un reporte fidedigno de los finiquitos
    ec_df = pd.read_excel(file, dtype={'Frente': object})
    c_df = pd.read_excel("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\EJJ.xlsm", sheet_name="C")

    # Enlistamos todos los contratos legales en el dataframe
    contratos = ec_df['Contrato Legal'].drop_duplicates()
    contratistas_correos = []
    hdf = ['Contratista', 'Conjunto', 'Contratos', 'Estimado', 'FG', 'Fecha', 'ID', 'URL']
    ndf =[]

    # Buscamos todos los archivos
    for root, dirs, files in os.walk("C:\\Users\\fprado\\REPORTES\\Abandono\\Correos\\"):
        for file in files:
            # Los archivos pdfs existentes y dentro de la lista de contratos
            if file.lower().endswith('.msg'):
                file_name, file_extension = os.path.splitext(file)
                contratistas_correos.append(file_name)                

    # iteramos contrato por contrato
    for contrato in contratos:
        # Comprobamos congruencia
        contratistas = ec_df[ec_df['Contrato Legal'] == contrato]['Contratista_x']
        conjunto = ec_df[ec_df['Contrato Legal'] == contrato]['Conjunto_x']
        estimado = ec_df[ec_df['Contrato Legal'] == contrato]['Estimado']
        fondo_g = ec_df[ec_df['Contrato Legal'] == contrato]['Penalizado FG']
        congruencia = contratistas.all() and conjunto.all()

        if congruencia:
            # Datos del contrato
            contratista = contratistas.drop_duplicates().values[0]
            
            if contratista in contratistas_correos:
                data_file = []
                conjunto = conjunto.drop_duplicates().values[0]
                Total_estimado = estimado.sum()
                Total_FG = fondo_g.sum()
                # Composición del url
                idd = str(c_df[c_df['Contrato'] == contrato]['conjunto_ID'].values[0])
                url = 'http://portal.javer.net/juridico/Paginas/Biblioteca.aspx?ContratoID=' + idd

                print(contratista, conjunto, contrato, Total_estimado, url)
                data_file.append(contratista)
                data_file.append(conjunto)
                data_file.append(contrato)
                data_file.append(Total_estimado)
                data_file.append(Total_FG)
                data_file.append("")
                data_file.append(idd)
                data_file.append(url)

        else:
            print("Something is very wrong")
        ndf.append(data_file)

    file_df = pd.DataFrame(ndf, columns=hdf)
    file_df.drop_duplicates(subset='Contrato', inplace=True)
    print(file_df)
    file_df.to_excel("C:\\Users\\fprado\\REPORTES\\A_FIN.xlsx")

#------------------------------------------------------creación de contratos-----------------------

def fill_tasks_mdc(insumos=False, vivienda=False):
    root = "C:\\Users\\fprado\\REPORTES\\Programas\\PROYECTO_X\\BDD\\BDD_Insumos_IV.xlsx"
    # Inputs
    dfs = []
    driver = create_driver(headless=False)
    df = pd.read_excel(root)
    go_contract(driver)
    secuencias_preferencias = {
        "U": ['03', '04', '06', '05'],
        "I": ['04', '03', '06', '05'],
        "P": ['06', '03', '04', '05'],
        "E": ['05', '03', '04', '06']
        }
    print(df)

    # Instrucciones de espera en modulo de construcción
    if insumos:
        try:
            INSUMOS(driver)
        except Exception as e:
            print(e)
    
    # Exigimos la letra hasta el cansancio
    while True:
        opciones = {
            "U": "Urbanización",
            "I": "Infraestructura",
            "P": "Plataformas",
            "E": "Equipamiento"
            }
        letra = window_input("Ingreso de Letra", "Selecciona la etapa", opciones)
        letra = letra.upper()

        # Verificar si la letra ingresada está en el diccionario
        if letra in secuencias_preferencias:
            secuencia = secuencias_preferencias[letra]
            print("La secuencia de preferencias para la letra {} es: {}".format(letra, secuencia))
            break

        else:
            print("La letra ingresada no tiene una secuencia de preferencias asociada.")

    # Filtrar las filas que cumplan la condición y guardarlas en lista
    for mi_variable in secuencia:
        condicion = df['Código'].str[2:4] == mi_variable
        resultados_filtrados = df[condicion]
        dfs.append(resultados_filtrados)

    if not vivienda:
        while True:
            try:
                print('Waiting for Asignar button')
                xpath = '//*[@id="Assign"]'
                swdw(driver, 1800, 0, xpath)
                swdw(driver, 2, 1, "ResultsTable:RowsToDisplay:0").send_keys(str(100) + Keys.ENTER)
                time.sleep(3)
                fill_list = beautiful_table(driver)
                history = pd.DataFrame(columns=['Descripción', 'Código'])

                for indice, task in fill_list.iterrows():
                    print(len(fill_list))
                    check_box = task['Select']
                    Tarea_completa = task['Descripción'].split("INCLUYE")
                    Tarea = Tarea_completa[0]
                    Unidad = task['Unidad de medida'].upper()
                    actividad = task['Actividad']
                    print(Tarea, Unidad, actividad)

                    try:
                        index = actividad.split('LOV:')
                        wait_id = 'ResultsTable:DescActividad:' + index[-1]
                        swdw(driver, 1, 1, wait_id)
                        continue

                    except TimeoutException as e:
                        print(e)

                    if Tarea in history['Descripción'].values:
                        elección = history.loc[history['Descripción'] == Tarea, 'Código'].values[0]
                        time.sleep(1)

                    else:
                        while True:
                            resultados = 10
                            valores = []

                            for sample_df in dfs:
                                filtro = sample_df['Unidad'].str.lower() == Unidad.lower()
                                df_filtrada = sample_df[filtro]
                                valores.append(encontrar_coincidencias(df_filtrada, "Descripción", Tarea, top_n=resultados))
                                resultados = 3
                            resultado = pd.concat(valores, ignore_index=True)
                            codigo_descripcion_dict = resultado.set_index('Código')['Descripción'].to_dict()
                            elección = window_input("Elección de insumos", Tarea, codigo_descripcion_dict)
                            nuevos_valores = [{'Descripción': Tarea, 'Código': elección}]
                            history_new = pd.DataFrame(nuevos_valores)
                            history = pd.concat([history, history_new], ignore_index=True)

                            if len(codigo_descripcion_dict) == 0:
                                print("Diccionario vacio")
                                print(resultado)
                                print(dfs[0], dfs[1], dfs[2])
                                continue

                                try:
                                    index = actividad.split('LOV:')
                                    wait_id = 'ResultsTable:DescActividad:' + index[-1]
                                    swdw(driver, 6, 1, wait_id)
                                    break
                                except:
                                    continue

                            else:
                                break

                    print(elección)
                    while True:
                        try:
                            swdw(driver, 3, 1, actividad).click()
                            swdw(driver, 3, 1, actividad).send_keys(elección + Keys.TAB)
                        except:
                            print("Something wrong")
                            time.sleep(0.1)

                        try:
                            index = actividad.split('LOV:')
                            wait_id = 'ResultsTable:DescActividad:' + index[-1]
                            swdw(driver, 6, 1, wait_id)
                            break
                        except:
                            continue

                    swdw(driver, 2, 1, check_box).click()
                break
            except TimeoutException as e:
                print(e)
                continue


def excrute_tasks(headless=False):
    driver = create_driver(headless=headless)
    go_in_mdc(driver, )



def delete_juridicos(ubic):
    
    def select_document(elemento):
        while True:
            try:
                select = Select(select_element)
                select.select_by_visible_text(elemento)
                break
            except StaleElementReferenceException:
                swdw(driver, 3, 1, 'typeFile').click()
                swdw(driver, 3, 1, 'typeFile').send_keys(elemento + Keys.ENTER)
                break
            except TimeoutException:
                continue

    driver = create_driver(driver_type='firefox', headless=False, download_folder="C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\")
    # Dataframe y lista de contratos
    df = pd.read_excel(ubic)
    contratos = df['Contrato']
    print(contratos)
    siroc = []

    # Iteración de contratos
    for contrato in contratos:
        try:
            # Obtenemos el valor ID del dataframe
            c_id = df.loc[df['Contrato'] == contrato, 'conjunto_ID'].values[0]

            # Partes del hipervínculo
            intro = "https://fprado@javer.com.mx:"
            med = JAVER_ID.PASSWORD
            outro =  "@portal.javer.net/juridico/Paginas/Biblioteca.aspx?ContratoID="
            link = intro + med + outro + str(c_id)
            print(contrato, c_id)

            # Ingresar a página y verificación de información
            while True:
                try:
                    # Acceso a la página
                    driver.get(link)
                    check = swdw(driver, 3, 1, "titleContrato").text

                except StaleElementReferenceException:
                    time.sleep(1)

                except TimeoutException:
                    break

                # Revisamos compatibilidad
                if check == contrato:
                    break

            # Definimos eliminador
            delete_button = swdw(driver, 2, 1, "btnEliminarContrato")

            # Extraer los existentes
            table = beautiful_table(driver, element="id", name="contentDS")
            # Obtenemos finiquitos y modificatorios existentes        
            mods = table[table['Archivo'].str.contains('CM')]
            fins = table[table['Archivo'].str.contains('CF')]

            if not fins.empty:
                hay_fin = True
                print('ya hay fin')
                continue

            # SIROC
            try:
                swdw(driver, 1, 0, "//table[@class='tblArchivosSIROC']//table/tbody[1]/tr[1]/td[1]/a[1]")
                c_fila = df.loc[df['Contrato'] == contrato]
                siroc.append(c_fila)
            except TimeoutException:
                print(contrato, "sin Siroc")

            if len(siroc) != 0:
                # Concatenamos todas las filas
                siroc_concat = pd.concat(siroc)
                # Guardamos el archivo
                siroc_concat.to_excel("siroc.xlsx", index=False)

            delete_button.click()
            # Imprimimos el texto de la alerta
            
            time.sleep(2)
            Alert(driver).accept()
            continue

        except Exception as e:
            print(e)
            continue
    driver.quit()
    organize_firefox()


def delete_fin_and_mods_juridicos(ubic):

    driver = create_driver(driver_type='firefox', headless=False, download_folder="C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\")
    df = pd.read_excel(ubic)
    contratos = df['Contrato']
    separador = '-' * 100

    # Iteración de contratos
    for contrato in contratos:
        try:
            # Obtenemos el valor ID del dataframe
            hyperlink = df.loc[df['Contrato'] == contrato, 'Hipervínculo'].values[0]

            # Partes del hipervínculo
            intro = "https://fprado@javer.com.mx:"
            med = JAVER_ID.PASSWORD
            link = intro + med + "@" + hyperlink[7:]
            print(separador)

            # Ingresar a página y verificación de información
            while True:
                try:
                    # Acceso a la página
                    driver.get(link)
                    ic(link)
                    check = swdw(driver, 4, 1, "titleContrato").text

                except StaleElementReferenceException:
                    time.sleep(1)

                except TimeoutException:
                    break

                # Revisamos compatibilidad
                if check == contrato:
                    ic(check, contrato)
                    break

            # Extraer los existentes
            time.sleep(1)
            table = beautiful_table(driver, element="id", name="contentDS")
            # Obtenemos finiquitos y modificatorios existentes        
            mods = table[table['Archivo'].str.contains('CM')]
            fins = table[table['Archivo'].str.contains('CF')]

            for i, fin in fins.iterrows():
                check = swdw(driver, 4, 1, "titleContrato").text
                file_fin = fin['Archivo']
                index_fin = table.index[table['Archivo'] == file_fin].values[0]
                index_one = str(int(index_fin) + 2)
                xpath_fin = f"//table[@id='contentDS']/tbody[1]/tr[{index_one}]/td[6]/select[1]"
                easy_select(driver, xpath_fin, "Eliminar", tipo='xpath')
                Alert(driver).accept()
                time.sleep(1)
                Alert(driver).accept()
                time.sleep(2)

            for i, mod in mods.iterrows():
                check = swdw(driver, 4, 1, "titleContrato").text
                file_mod = mod['Archivo']
                index_mod = table.index[table['Archivo'] == file_mod].values[0]
                index_one = str(int(index_mod) + 2)
                xpath_mod = f"//table[@id='contentDS']/tbody[1]/tr[{index_one}]/td[6]/select[1]"
                easy_select(driver, xpath_mod, "Eliminar", tipo='xpath')
                Alert(driver).accept()
                time.sleep(1)
                Alert(driver).accept()
                time.sleep(2)

            print(separador)
            print("\n")
        except Exception as e:
            print(e)
            continue
    driver.quit()
    organize_firefox()


def nuevo_finiquitador(driver, df, mod=True, fin=True):

    def add_doc(driver, tipe_doc):

        add_doc = swdw(driver, 2, 1, 'AddDoc')
        add_doc.click()
        time.sleep(0.5)

        easy_select(driver, 'DocType', tipe_doc)
        time.sleep(1)
        easy_select(driver, 'Signed', "No")
        time.sleep(0.5)

        aceptar = swdw(driver, 2, 1, "Aceptar")
        aceptar.click()


    def modificator(driver, df, mod=True):
        flag = False

        save_flash_memory(df)
        while True:
            if not flag:
                row, df = take_first_row()
            if df.empty:
                organize_firefox()
                break

            flag = False
            contrato = row['Contrato']
            conjunto = row['Conjunto']
            total_estimado = row['Total_estimado']
            print('\n', '-'*100)
            print(contrato, conjunto)

            try:
                get_contract(driver, conjunto, contrato)
            except IndexError:
                print("Index Error")

            fecha = row['Fecha']
            print(fecha)
            if isinstance(fecha, str):
                fecha_ultima = datetime.strptime(fecha, "%Y-%m-%d %H:%M:%S.%f")
            elif isinstance(fecha, datetime):
                fecha_ultima = fecha
                print(fecha)
            else:
                ic(fecha, type(fecha))
                fecha_ultima = fecha.to_pydatetime()

            dia = str(fecha_ultima.strftime("%d"))
            mes = str(fecha_ultima.strftime("%m"))
            anno = str(fecha_ultima.strftime("%Y"))
            fecha_modificatorio = dia + "/" + mes + "/" + anno

            go_doc = swdw(driver, 1, 1, 'LegalContractsLink')
            go_doc.click()

            list_of_contracts = beautiful_table(driver)
            modificatorio = list_of_contracts['Descripción'].str.contains('modificatorio').any()
            anticipada = list_of_contracts['Descripción'].str.contains('terminación anticipada').any()
            finiquito =  list_of_contracts['Descripción'].str.contains('finiquito').any()
            ic(list_of_contracts, modificatorio, anticipada, finiquito)

            last_days = list_of_contracts[
                    list_of_contracts['Fecha terminación'].notna() & (list_of_contracts['Fecha terminación'] != '')
                    ]
            last_days = last_days['Fecha terminación'].iloc[-1]
            fecha_terminacion = datetime.strptime(last_days, '%d-%m-%Y %H:%M:%S')

            if fecha_terminacion < fecha_ultima or total_estimado <= 1: 
                mod = True
                dia = str(fecha_terminacion.strftime("%d"))
                mes = str(fecha_terminacion.strftime("%m"))
                anno = str(fecha_terminacion.strftime("%Y"))
                last_date = fecha_terminacion
            else:
                mod = False
                dia = str(fecha_ultima.strftime("%d"))
                mes = str(fecha_ultima.strftime("%m"))
                anno = str(fecha_ultima.strftime("%Y"))
                last_date = fecha_ultima
            ic(mod)

            fecha_firma = dia + "/" + mes + "/" + anno
            importe_inicial = swdw(driver, 2, 1, 'ImporteContrato').text
            importe_inicial = float(importe_inicial.replace("$", "").replace(",", ""))
            porcentaje = (total_estimado / importe_inicial) * 100
            if porcentaje > 100:
                porcentaje = 100
            porcentaje = str(round(porcentaje, 0))
            total_estimado = str(total_estimado)   

            if finiquito or anticipada:
                print('_'*20, 'Este está cerrado','_'*20, '\n', '-'*100, '\n')
                swdw(driver, 1, 1, "XXMCAN_CONTRACT_SEARCH").click()
                continue

            elif modificatorio:
                modificatorios = list_of_contracts[list_of_contracts['Descripción'].str.contains('modificatorio', na=False)]
                modificatorio = modificatorios.iloc[-1]
                last_date_mod = datetime.strptime(modificatorio['Fecha terminación'], '%d-%m-%Y %H:%M:%S')
                last_amount_mod_str = modificatorio['Monto']
                last_amount_mod = float(last_amount_mod_str.replace("$", "").replace(",", ""))
                last_amount_mod = str(last_amount_mod)
                mismo_monto = last_amount_mod == total_estimado
                misma_fecha = last_date_mod == last_date
                ic(last_amount_mod, total_estimado)
                ic(last_date_mod, last_date)

                if mismo_monto and misma_fecha:
                    print("Modificatorio repetido")
                    swdw(driver, 2, 1, "XXMCAN_CONTRACT_SEARCH").click()
                    continue

            if mod:
                add_doc(driver, 'Convenio modificatorio')
                date_id = 'Date'
            else:
                add_doc(driver, 'Convenio de terminación de obra anticipada')
                date_id = 'TermDate'

            input_date = swdw(driver, 1, 1, date_id)
            input_date.send_keys(fecha_modificatorio + Keys.TAB)
            time.sleep(0.5)

            if not mod:
                advanced_perc = swdw(driver, 1, 1, 'AdvancedPerc')
                advanced_perc.send_keys(porcentaje + Keys.TAB)
                time.sleep(1)

            input_monto = swdw(driver, 1, 1, 'Amount')
            input_monto.send_keys(total_estimado + Keys.TAB)
            time.sleep(1)

            if mod:
                input_fecha_modificatorio = swdw(driver, 1, 1, 'FechaConvenio')
                input_fecha_modificatorio.send_keys(fecha_firma + Keys.TAB)
                time.sleep(0.5)

                swdw(driver, 1, 1, 'Budget').click()
                time.sleep(0.5)
                swdw(driver, 1, 1, 'DetailedItems').click()
                time.sleep(0.5)


            swdw(driver, 2, 1, 'Apply').click()
            time.sleep(3)

            # list_of_contracts = beautiful_table(driver)
            # modificatorio = list_of_contracts['Descripción'].str.contains('modificatorio').any()
            # finiquito =  list_of_contracts['Descripción'].str.contains('finiquito').any()
            try:
                swdw(driver, 2, 1, "XXMCAN_CONTRACT_SEARCH").click()
                print('|'*20, "Este fué exitoso", '|'*20,'\n', '-'*100, '\n')
            except TimeoutException:
                print('*'*20, 'ESTE FALLÓ','*'*20, '\n', '-'*100, '\n')
                add_df = pd.concat([df, row], ignore_index=True)
                flash_memory = get_here(folder="cache", name="flash_memory_mth.xlsx")
                add_df.to_excel(flash_memory, index=False)
                acceder_oracle(driver)
                go_contract(driver)


    def finiquitator(driver, df):

        flag = False
        save_flash_memory(df)
        print('\n'*2, '|'*50, 'Empiezan finiquitos', '|'*50, '\n'*2)
        while True:
            if not flag:
                row, df = take_first_row()
            
            if df.empty:
                organize_firefox()
                break

            # try:
            flag = False

            print('-'*100, '\n', row['Contrato'], row['Conjunto'])
            conjunto = row['Conjunto']
            contrato = row['Contrato']
            ultima_modificacion = datetime.strptime(row['Ultima_modificación'], '%d/%m/%Y')
            fecha_de_terminación = datetime.strptime(row['Fecha_de_terminación'], '%d/%m/%Y')

            try:
                get_contract(driver, conjunto, contrato)
            except IndexError:
                break

            go_doc = swdw(driver, 1, 1, 'LegalContractsLink')
            go_doc.click()

            list_of_contracts = beautiful_table(driver)
            finiquito =  list_of_contracts['Descripción'].str.contains('finiquito').any()
            modificatorio = list_of_contracts['Descripción'].str.contains('modificatorio').any()
            anticipada = list_of_contracts['Descripción'].str.contains('terminación anticipada').any()
            ic(list_of_contracts, finiquito, modificatorio, anticipada)


            if finiquito or anticipada:
                print(' '*15, f'Ya está terminado {ic(finiquito, anticipada)}')
                swdw(driver, 1, 1, "XXMCAN_CONTRACT_SEARCH").click()
                continue

            elif not finiquito and modificatorio:
                print(' '*15, f'Finiquito elaborado {ic(finiquito, anticipada)}')
                add_doc(driver, 'Convenio finiquito')
                swdw(driver, 2, 0, "//button[@title='Sí']").click()
                time.sleep(3)

            else:
                print('*'*15, f'ERROR {ic(finiquito, anticipada)}', '*'*15)
            swdw(driver, 2, 1, 'XXMCAN_CONTRACT_SEARCH').click()
            swdw(driver, 2, 1, 'XXMCAN_CONTRACT_SEARCH').click()
            print('-'*100, '\n')


    if mod:
        modificator(driver, df)
    if fin:
        finiquitator(driver, df)
    driver.quit()



def download_scans():

    path_download = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\scan\\"

    # Inicializar objeto Outlook, in a carpeta SCA y obtener los mensajes
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
    sca = outlook.Folders.Item("fprado@javer.com.mx").Folders.Item("Bandeja de entrada").Folders.Item('CAO').Folders.Item("SCA")
    done = outlook.Folders.Item("fprado@javer.com.mx").Folders.Item("Bandeja de entrada").Folders.Item('CAO').Folders.Item("SCA-DONE")
    messages = sca.Items

    for mail in messages:
        subject = mail.Subject
        attachments = mail.Attachments
        n_subject = subject.upper()
        
        # revisamos que tenga escaneo
        if attachments.Count > 0:
            for attachment in attachments:
                if n_subject.startswith("MESSAGE FROM"):
                    # Le quitamos la extension al nombre
                    c_adjunto = attachment.FileName.upper()
                    match = re.search(r'^(.*)\.*$', c_adjunto)
                    if match:
                        n_adjunto = match.group(1)
                    else:
                        n_adjunto = c_adjunto
                else:
                    n_adjunto = n_subject + ".pdf"
                
                # Handle existing filenames with a counter
                file_path = os.path.join(path_download, n_adjunto)
                counter = 0
                while os.path.exists(file_path):
                    counter += 1
                    name, extention = os.path.splitext(n_adjunto)
                    file_path = os.path.join(path_download, f"{name}_{counter}{extention}")

                print(file_path)
                # Save attachment
                try:
                    attachment.SaveAsFile(file_path)

                except Exception as e:
                    print(f"Error saving attachment: {e}")

                mail.UnRead = False
                try:
                    mail.Move(done)
                except:
                    continue


def download_finalizados_doc2sign():

    path_download = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\DOC2SIGN\\FIRMADOS ZIP\\"

    # Inicializar objeto Outlook, in a carpeta SCA y obtener los mensajes
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
    finalizados = outlook.Folders.Item("fprado@javer.com.mx").Folders.Item("Bandeja de entrada").Folders.Item('Doc2Sign').Folders.Item("Finalizados")
    messages = finalizados.Items

    for mail in messages:
        subject = mail.Subject
        upper_subject = subject.upper()

        if mail.Unread:

            # Patrón para extraer el nombre del archivo PDF del asunto
            pattern = r'EL DOCUMENTO (.+?)\.PDF'
            file_name_match = re.search(pattern, upper_subject)

            if file_name_match:
                file_name = file_name_match.group(1) + ".zip"
                
                # Obtener el cuerpo del correo en formato HTML
                if mail.BodyFormat == 2:  # HTML
                    # Patrón para encontrar el hipervínculo con el texto "descargar documento"
                    body_html = mail.HTMLBody
                    link_pattern = r'<a\s+href="([^"]+)"[^>]*>descargar documento</a>'
                    link_match = re.search(link_pattern, body_html, re.IGNORECASE)
                    
                    if link_match:
                        # vamos al hipervínculo de descarga
                        download_link = link_match.group(1)
                        response = requests.get(download_link, stream=True)
                        time.sleep(1)

                        if response.status_code == 200:
                            new_filename = os.path.join(path_download, file_name)

                            with open(new_filename, 'wb') as file:
                                # Escribe el contenido del archivo
                                for chunk in response.iter_content(chunk_size=8192):
                                    file.write(chunk)

                            # Marcamos el correo como leído
                            print("Listo, ya se descargó el archivo")
                            mail.UnRead = False

                    else:
                        print(f"No se encontró el enlace de descarga.\n{body_html}")
                else:
                    print(f"El cuerpo del correo no está en formato HTML.\n{mail.BodyFormat}")
            else:
                print('No hubo coincidencias en el asunto')
                ic(file_name_match, upper_subject)


def download_files_from_contractS(root_table_conjunto_contrato, list_of_files):

    # creamos un web browser bot con un programa creado previamente
    driver = create_driver(headless=False)

    # Leemos el excel como un pandas data frame 
    df = pd.read_excel(root_table_conjunto_contrato)
    
    # Iteramos cada una de las filas del data frame
    for i, row in df.iterrows():

        # De cada fila obtenemos la variable contrato y conjunto
        contrato = row['Contrato']
        conjunto = row['Conjunto']
        
        # Usamos un programa creado con anterioridad, para ir a cada contrato con el bot
        get_contract(driver, conjunto, contrato)

        # Creamos una lista que contiene los reportes disponibles
        list_of_reports = ('Alcance de contrato por paquete',
                            'Alcance detallado de contrato',
                            'Comparacion de actividades de contrato vs presupuesto',
                            'Estado de cuenta',
                            'Explosion de insumos de contrato por categoria',
                            'Reporte de calificacion de contratista')
        
        # Creamos una lista de coincidencias entre ambas listas
        coincidencias = list(set(list_of_files) & set(list_of_reports))

        # Si existen coincidencias procedemos con acciones para los reportes
        if coincidencias:

            # Hacemos que el bot se dirija a los reportes en el contrato usando programa simple de webdriverwait
            swdw(driver, 3, 1, 'ReportLink').click()

            # Iteramos para cada reporte en la lista de coincidencias
            for reporte in coincidencias:

                ic(reporte)
                while True:
                    try:
                        # Convertimos en dataframe la tabla de la página web usando beautiful soup, en un programa previo
                        time.sleep(1)
                        reports_table = beautiful_table(driver)
                        col_switcher = reports_table['colSwitcherExecute'].loc[reports_table['Descripción'] == reporte].values[0]
                        swdw(driver, 2, 1, col_switcher).click()
                        # Actualizamos la tabla de reportes
                        swdw(driver, 2, 1, "buttonRefreshReport").click()
                        time.sleep(2)
                        reports_table = beautiful_table(driver)
                        col_pdf = reports_table['colPdf'].loc[reports_table['Descripción'] == reporte].values[0]
                        ic(reporte, col_pdf)
                        if col_pdf != None:
                            break

                    except (TimeoutException, StaleElementReferenceException, AttributeError):
                        continue

                    except Exception as e:
                        print(e)



            for reporte in coincidencias:

                ic(reporte)
                while True:
                    try:
                        time.sleep(1)
                        # Actualizamos la tabla de reportes
                        reports_table = beautiful_table(driver)
                        # Extraemos la ubicación del botón para descargar el reporte
                        col_pdf = reports_table['colPdf'].loc[reports_table['Descripción'] == reporte].values[0]
                        col_pdf.startswith("TableReportsRN:PDF1:")
                        break
                    except Exception as e:
                        swdw(driver, 2, 1, "buttonRefreshReport").click()
                        continue

                while True:
                    try:
                        # Actualizamos la tabla de reportes
                        time.sleep(1.5)
                        swdw(driver, 2, 1, "buttonRefreshReport").click()
                        time.sleep(1.5)
                        reports_table = beautiful_table(driver)
                        # Extraemos la ubicación del botón para descargar el reporte
                        col_pdf = reports_table['colPdf'].loc[reports_table['Descripción'] == reporte].values[0]
                        ic(reports_table['Descripción'].loc[reports_table['Descripción'] == reporte].values[0])
                        ic(reporte, col_pdf)
                        # Una vez listo el archivo para descargar damos click en el botón de pdf
                        swdw(driver, 2, 1, col_pdf).click()
                        time.sleep(1)
                        swdw(driver, 2, 1, col_pdf).click()
                        break

                    except Exception as e:
                        time.sleep(1)                      

        # Verificar si algún documento de contrato están en la lista
        have_finiquito = "Finiquito" in list_of_files
        have_modificatorio = "Modificatorio" in list_of_files
        have_contrato = "Contrato" in list_of_files

        ic(have_finiquito, have_contrato, have_modificatorio)

        # Si se dá alguna de las 3 condiciones
        if have_finiquito or have_modificatorio or have_contrato:

            # Damos click a la pestaña de Contratos Legales
            swdw(driver, 2, 1, "LegalContractsLink").click()
            time.sleep(1)


            # Adquirimos la tabla de documentos
            while True:
                try:
                    swdw(driver, 2, 0, "//h3[@class='x44' and text()='Documentos legales']")
                    contract_table = beautiful_table(driver)
                    ic(contract_table)
                    break

                except AttributeError:
                    swdw(driver, 2, 1, "LegalContractsLink").click()
                    time.sleep(1)

                except TimeoutException:
                    time.sleep(1)
                    swdw(driver, 2, 1, "LegalContractsLink").click()



            if have_finiquito:

                try:
                    # Extraemos el último elemento de la tabla
                    row = contract_table.tail(1)

                    # Adquirimos, el nombre del documento, creamos un xpath y damos click en el documento
                    documento = row.iloc[0]['Documento']
                    xpath = f"//*[@title='{documento}']"
                    swdw(driver, 2, 0, xpath).click()
                    time.sleep(5)

                except Exception as e:
                    print(e) 

            if have_contrato:

                try:
                    # Extraemos el primer elemento de la tabla
                    row = contract_table.head(1)

                    # Adquirimos, el nombre del documento, creamos un xpath y damos click en el documento
                    documento = row.iloc[0]['Documento']

                    xpath = f"//*[@title='{documento}']"
                    swdw(driver, 2, 0, xpath).click()
                    time.sleep(5)

                except Exception as e:
                    print(e) 

            if have_modificatorio:

                try:
                    #Extraemos todas las filas intermedias
                    rows = contract_table.iloc[1:-1]

                    # Iteramos las filas
                    for index, row in rows.iterrows():

                        # Adquirimos, el nombre del documento, creamos un xpath y damos click en el documento
                        documento = row.iloc[0]['Documento']
                        xpath = f"//*[@title='{documento}']"
                        swdw(driver, 2, 0, xpath).click()
                        time.sleep(5)

                except Exception as e:
                    print(e)                
        

        swdw(driver, 2, 1, 'XXMCAN_CONTRACT_SEARCH').click()
        swdw(driver, 2, 1, 'XXMCAN_CONTRACT_SEARCH').click()

    driver.quit()
    organize_firefox()  


def el_nuevo_finiquitador(ruta, mod=True, fin=True):
    df_fin = pd.read_excel(ruta)
    ic(df_fin)
    df_fin.sort_values(by='Conjunto', ascending=True, inplace=True)
    driver = create_driver(headless=False)
    nuevo_finiquitador(driver, df_fin, mod=mod, fin=fin)


def delete_files_from_contract(root_table_conjunto_contrato, number_files):

    # creamos un web browser bot con un programa creado previamente
    driver = create_driver(headless=False)

    # Leemos el excel como un pandas data frame 
    df = pd.read_excel(root_table_conjunto_contrato)
    
    # Iteramos cada una de las filas del data frame
    for i, row in df.iterrows():

        # De cada fila obtenemos la variable contrato y conjunto
        contrato = row['Contrato']
        conjunto = row['Conjunto']
        
        # Usamos un programa creado con anterioridad, para ir a cada contrato con el bot
        get_contract_contracts(driver, conjunto, contrato)

        contracts = beautiful_table(driver)
        # finiquito =  list_of_contracts['Descripción'].str.contains('finiquito').any()
        # modificatorio = list_of_contracts['Descripción'].str.contains('modificatorio').any()

        files = contracts.iloc[number_files]

        for file in files:
            contracts = beautiful_table(driver)
            files = contracts.iloc[number_files]
            ic(files)
            botones = files[files['Eliminar'].str.contains(':Habilitar', na=False)]
            ic(botones)
            if not botones.empty:
                eliminar = botones['Eliminar'].iloc[0]
                ic(eliminar)
                swdw(driver, 1, 1, eliminar).click()
                swdw(driver, 1, 0, "//button[@title='Sí']").click()
                time.sleep(1)
            else:
                print('No hubo continuidad')
        swdw(driver, 1, 1, 'Return').click()
    driver.quit()
            




def rename_downloaded_contracts():

    ruta = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\*.pdf"
    pdf_files_downloaded = glob.glob(ruta) + glob.glob(ruta.replace(".pdf", ".PDF"))
    contracts_dict = {}  # Diccionario para almacenar los archivos por contrato

    for pdf_file in pdf_files_downloaded:
        try:
            open_pdf = pdfplumber.open(pdf_file)
        except Exception as e:
            print(e)
            continue

        # Extraemos texto de todas las páginas
        with open_pdf as pdf:
            try:
                pdf_text = ''
                for page in pdf.pages:
                    pdf_text += page.extract_text() + "\n"
            except Exception as e:
                print(e)
                continue

        # Encontramos el contrato del texto
        patron_contrato = r"QRO-\w{3}-\w{3}-\d{6}"
        found_contrato = re.search(patron_contrato, pdf_text)

        if found_contrato:

            try:
                nuevo_nombre_base = found_contrato.group(0)

                patron_documento = r"Datoslegalesenanexoque"
                found_documento = re.search(patron_documento, pdf_text)

                patron_adc = r"Alcance detallado de contrato"
                found_adc = re.search(patron_adc, pdf_text)

                patron_acp = r"Alcance de contrato por paquete"
                found_acp = re.search(patron_acp, pdf_text)

                patron_ei = r"Explosion de insumos de contrato por categoria"
                found_ei = re.search(patron_ei, pdf_text)


                if found_adc:
                    tipo_documento = "ADC"

                elif found_acp:
                    tipo_documento = "ACP"

                elif found_ei:
                    tipo_documento = "EI"


                elif found_documento:

                    patron_conjunto = r"E\d{2}-\d{2}-\w\d{2}-\d{2}(?:-\d{3})?"
                    patron_contratista = r"CONTRATISTA\s*(.*)\nAPODERADO"

                    # Corregir el uso del patrón adecuado
                    found_conjunto = re.search(patron_conjunto, pdf_text)
                    found_contratista = re.search(patron_contratista, pdf_text)

                    if found_conjunto:
                        conjunto = found_conjunto.group(0)
                    else:
                        conjunto = "Desconocido"
                        print("\n", pdf_text[:800], "\n")

                    if found_contratista:
                        contratista = found_contratista.group(1).strip()  # Capturar el nombre del contratista y eliminar espacios
                    else:
                        patron_contratista = r"PRESTADOR\s*(.*)\nAPODERADO"
                        found_contratista = re.search(patron_contratista, pdf_text)
                        if found_contratista:
                            contratista = found_contratista.group(1).strip()
                        else:
                            contratista = "Desconocido"
                            print("\n", pdf_text[:800], "\n")

                else:
                    print("No founds")
                    ic(pdf_file, nuevo_nombre_base)
                    ic(pdf_text)
                    continue
                
                if found_documento:
                    nuevo_nombre = f"{nuevo_nombre_base} - {conjunto} - {contratista} - C.pdf"

                else:
                    nuevo_nombre = f"{nuevo_nombre_base} - {tipo_documento}.pdf"

                nuevo_ruta = os.path.join(os.path.dirname(pdf_file), nuevo_nombre)            
                
                # Renombrar el archivo
                try:
                    os.rename(pdf_file, nuevo_ruta)
                    print(f"Archivo renombrado a: {nuevo_nombre}")

                except OSError as e:
                    if e.winerror == 183:  # WinError 183: El archivo ya existe
                        print(f"El archivo {pdf_file} ya existe. Eliminando...")
                        os.remove(pdf_file)  # Elimina el archivo si ya existe
                    else:
                        print(f"Error al renombrar {pdf_file}: {e}")


            except Exception as e:
                print(e)


    ruta_1 = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\CONTRATOS\\"
    pdf_files_downloaded = glob.glob(ruta)

    # Diccionario para agrupar archivos por patrón de contrato
    patrones_archivos = {}

    # Expresión regular para extraer el patrón de contrato
    patron_contrato = r"QRO-\w{3}-\w{3}-\d{6}"

    # Agrupar archivos por su patrón de contrato
    for pdf_file in pdf_files_downloaded:
        match = re.search(patron_contrato, os.path.basename(pdf_file))
        if match:
            contrato = match.group(0)
            if contrato not in patrones_archivos:
                patrones_archivos[contrato] = []
            patrones_archivos[contrato].append(pdf_file)

    # Orden y nombres para la combinación
    orden_prioridad = [" - C", " - ACP", " - ADC", " - EI"]

    # Combinar PDFs por patrón de contrato
    for contrato, archivos in patrones_archivos.items():
        if len(archivos) >= 3:
            print(contrato, "'\n", archivos)
            # Filtrar el archivo que tiene "- C" para usar su nombre
            archivo_c = next((archivo for archivo in archivos if "- C" in archivo), None)

            # Ordenar archivos según la prioridad
            archivos_ordenados = sorted(archivos, key=lambda x: [orden_prioridad.index(x[-8:-4]) if x[-8:-4] in orden_prioridad else len(orden_prioridad)])

            # Crear un nuevo PDF
            writer = PdfWriter()
            
            for archivo in archivos_ordenados:
                reader = PdfReader(archivo)
                for page in reader.pages:
                    writer.add_page(page)

            # Si se encontró el archivo "- C", usar su nombre para el nuevo archivo
            if archivo_c:
                nuevo_nombre = os.path.basename(archivo_c)  # Obtener el nombre del archivo "- C"
                nuevo_ruta = os.path.join(ruta_1, nuevo_nombre)  # Guardar en la misma carpeta

                # Guardar el nuevo PDF combinado
                with open(nuevo_ruta, "wb") as f_out:
                    writer.write(f_out)

                print(f"Combinado: {nuevo_ruta}")
            else:
                print(f"No se encontró un archivo '- C' para el contrato: {contrato}")

            for archivo in archivos_ordenados:
                os.remove(archivo)
            

def upload_signed_doc2sign():
    driver = create_driver(headless=False)
    path_zip = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\DOC2SIGN\\FIRMADOS ZIP\\"
    path_cancelados = path_zip.replace('FIRMADOS ZIP', 'CANCELADOS')
    path_to_heavy = path_zip.replace('FIRMADOS ZIP', 'TO HEAVY')
    path_cargados = path_to_heavy = path_zip.replace('FIRMADOS ZIP', 'CARGADOS')
    path_none = path_zip.replace('FIRMADOS ZIP', 'NONE')

    zip_files = glob.glob(path_zip + "*.zip")
    for zip_file in zip_files:
        folder, file_extention = os.path.split(zip_file)
        file, extention = os.path.splitext(file_extention)

        file_size = round(os.path.getsize(zip_file) / (1024 * 1024), 3)
        if file_size > 10:
            new_path = os.path.join(path_to_heavy, file_extention)
            os.rename(zip_file, new_path)
            print(f'File to heavy: {file_extention}, size; {file_size}')
            continue


        patron_contrato = r"QRO-\w{3}-\w{3}-\d{5}\d?"
        match_contrato = re.search(patron_contrato, file)

        patron_conjunto = r"E\d{2}-\d{2}-\w\d{2}-\d{2}(?:-\d{3})?"
        match_conjunto = re.search(patron_conjunto, file)

        if match_conjunto and match_contrato:
            contrato = match_contrato.group(0)  
            conjunto = match_conjunto.group(0)

            rest = file.replace(contrato, '')
            rest = rest.replace(conjunto, '')
            rest = rest.upper()

            patron_c = r"\sC$"
            patron_cf =r" CF"
            patron_cm =r" CM"
            patron_adendum =r" ADENDUM"

            match_c = re.search(patron_c, rest)
            match_cf = re.search(patron_cf, rest)
            match_cm = re.search(patron_cm, rest)
            match_adendum = re.search(patron_adendum, rest)

            if match_c:
                file_type = 'contrato inicial'
                rest = rest[:-1]
            elif match_cf:
                file_type = 'finiquito'
                rest = rest.replace(match_cf.group(0), '')
            elif match_cm:
                file_type = 'modificatorio'
                rest = rest.replace(match_cm.group(0), '')
            elif match_adendum:
                file_type = 'adendum'
                rest = rest.replace(match_adendum.group(0), '')
            else:
                file_type = False

            rest = rest.replace('-', '')
            rest = rest.strip()
            contratista = re.sub(r'\s+', ' ', rest)


            if not file_type:
                new_path = os.path.join(path_none, file_extention)
                os.rename(zip_file, new_path)
                print(f'File without type: {file_extention}')
                continue

            else:
                doc_n = contrato + ' - ' + file_type
                new_path = os.path.join(path_cargados, file_extention)
                print(f'{conjunto} - {contrato} - {contratista}')
                charge_document_contract(driver, conjunto, contrato, zip_file, doc_n, file_type)
                os.rename(zip_file, new_path)

        else:
            new_path = os.path.join(path_none, file_extention)
            os.rename(zip_file, new_path)
            print(f'File without type: {file_extention}')
            continue



#--------------------------------------------------------------------

# upload_signed_doc2sign()
# download_scans()
# download_finalizados_doc2sign()
# cambiar_extension_carpeta("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\scan\\")

#--------------------------------------------------------------------

# fill_tasks_mdc(insumos=True, vivienda=True)
# investigar_contratos("C:\\Users\\fprado\\REPORTES\\Abandono\\viejos.xlsx")

#----------------------------Ejecutor--------------------------------

# HV_manager(threads=4, headless=False)
mdc.update_contract(headless=False)
SIROC_update()

#---------------------------Finiquitar------------------------------

# actualizar_BDD_de_mis_fraccionamientos()
 
# fi = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\Libro1.xlsx" 
# fo = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\2018 UJN\\"
# ew = "C"
# busca_correo_archivos(fi, fo, ew)
# organizar_estados_cuenta()
# compilar_estads_cuenta()

# delete_juridicos("C:\\Users\\fprado\\REPORTES\\Contratos_inexistentes.xlsx")
# delete_fin_and_mods_juridicos("C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\acdc.xlsx")

root_table_conjunto_contrato = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\Libro1.xlsx"
list_of_files = ('Explosion de insumos de contrato por categoria', 'Alcance detallado de contrato', 'Contrato')
# download_files_from_contractS(root_table_conjunto_contrato, list_of_files)
# rename_downloaded_contracts()

# fecha_ultima_estimacion(4, 'C:\\Users\\fprado\\REPORTES\\para_finiquitar.xlsx', 'C:\\Users\\fprado\\REPORTES\\merg_finiquitos.xlsx', headless=False)
# merger_dates('C:\\Users\\fprado\\REPORTES\\para_finiquitar.xlsx')
# archivo_finiquitador("C:\\Users\\fprado\\REPORTES\\Acta de abandono.xlsx")
# viejo_finiquitador("C:\\Users\\fprado\\REPORTES\\finiquitador.xlsx", mod=True, fin=True, abandono=False)
# el_nuevo_finiquitador("C:\\Users\\fprado\\REPORTES\\finiquitador.xlsx")

# apagar_pc()
#'Alcance de contrato por paquete'

path_1 = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\Libro1.xlsx'
# delete_files_from_contract(path_1, slice(-2, None))