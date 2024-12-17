import time
import ast
import os
import re
import math
import logging
import threading
import traceback
import glob
import datetime
from icecream import ic
from datetime import datetime as dt
import pandas as pd
import numpy as np
import BDD_A as BDD
from Multiherramienta import *
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (UnexpectedAlertPresentException, 
    ElementClickInterceptedException, StaleElementReferenceException, NoSuchFrameException, 
    NoSuchElementException, TimeoutException, NoSuchWindowException, NoAlertPresentException)
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

HERE = get_here()

#-----------------------------------------------Buscadores-----------------------------------------

def update_contract(headless=True):
    # Descargar el reporte de contratos
    ruta_descarga = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"

    def download_report_contracts():
        driver = create_driver(headless=headless)
        go_sistema_reportes(driver)
        busca_reporte_sistema_de_reportes(driver, 'Reporte de contratos legales [P]', 'colSwitcherExecute')
        xpath_origen = "/html/body/form/span[2]/div/div[3]/div[1]/div[2]/table[2]/\
                    tbody/tr[4]/td/table/tbody/tr/td/div/div[2]/table/tbody/tr/td[2]/\
                    table/tbody/tr[1]/td[3]/span/a/img"
        stf(driver, 0, xpath_origen, "Actuales")
        swdw(driver, 1, 1, 'ButtonExecute_uixr').click()
        busca_reporte_sistema_de_reportes(driver, 'Reporte de contratos legales [P]', 'colXls')
        counter = 0
        while True:
            if counter == 25:
                swdw(driver, 2, 1, "TableReportsRN:XLS1img:0").click()
                counter = -5
                time.sleep(3)
            
            file_name = "XXMCAN_XML_Report_Publisher*.xls"
            file_root = os.path.join(ruta_descarga, file_name)
            archivos = glob.glob(file_root)

            file_part = "XXMCAN_XML_Report_Publisher*.xls.part"
            part_root = os.path.join(ruta_descarga, file_part)
            parts = glob.glob(part_root)

            if not archivos or parts:
                counter += 1
                time.sleep(3)
                continue
            else:
                time.sleep(2)
                break

        print("success downloading contracts")
        driver.quit()
    
    # Convertir el dataframe
    def dataframe_contratos():
        # inputs
        ruta_reportes = "C:\\Users\\fprado\\REPORTES\\"

        # this_day = datetime.date.today()
        # this_day = str(this_day)
        file_name = "XXMCAN_XML_Report_Publisher*.xls"
        file_root = os.path.join(ruta_descarga, file_name)

        archivos = glob.glob(file_root)
        print(archivos)
        archivo_mas_reciente = max(archivos, key=os.path.getmtime)

        df_reporte_contratos = pd.read_html(archivo_mas_reciente)[2]
        df_reporte_contratos = df_reporte_contratos[1:]
        df_reporte_contratos.columns = df_reporte_contratos.iloc[0]
        df_reporte_contratos = df_reporte_contratos[1:]
        print(df_reporte_contratos)
        df_sirocs = df_reporte_contratos.loc[df_reporte_contratos['Tipo documento'] == "SIROC"]
        df_contratos = df_reporte_contratos.loc[df_reporte_contratos['Tipo documento'] != "SIROC"]
        with pd.ExcelWriter(ruta_reportes + 'CONCENTRADO_J.xlsx') as writer:
            df_sirocs.to_excel(writer, sheet_name='SIR')
            df_contratos.to_excel(writer, sheet_name='CON')
        os.remove(archivo_mas_reciente)
        print("Success changing DB contracts")

    # Funciones
    download_report_contracts()
    time.sleep(3)
    dataframe_contratos()

def lookup_contract(driver, org, frente, conjunto):
# Revisa dentro de lista de contratos    
    while True:
        # Usando herramientas para enviar datos en frames, mandamos la info de cada contrato
        try:
            search = "//table[@id='CamposRN']//table/tbody[1]/tr[13]/td[3]/button[1]"
            
            # Buscamos organización
            swdw(driver, 2, 0, "//img[@title='Buscar: Organización']")
            stf(driver, 0, "//img[@title='Buscar: Organización']", org)
            time.sleep(0.5)
            
            # Buscamos frentes
            swdw(driver, 2, 0, "//img[@title='Buscar: Frente']")
            stf(driver, 0, "//img[@title='Buscar: Frente']", frente)
            time.sleep(0.5)
            
            # Buscamos conjunto
            swdw(driver, 2, 0, "//img[@title='Buscar: Conjunto de Construcción']")
            stf(driver, 0, "//img[@title='Buscar: Conjunto de Construcción']", conjunto)
            time.sleep(0.5)
            
            # Dar click en buscar
            swdw(driver, 2, 0, search).click()

            # Ampliamos a 100
            select_element = driver.find_element(By.ID, "ResultsTable:ResultsDisplayed:0")
            select = Select(select_element)
            select.select_by_visible_text("100")

            swdw(driver, 2, 1, "ResultsTable:Mostrar:0")
            break
            
        except (TimeoutException, StaleElementReferenceException):
            time.sleep(0.5)


#----------------------------Bases De Datos--------------------------------------------------------


def frentes_existentes(bdd, headless=True, threads=4):
    # Descarga los frentes existentes

    # Creamos un dataframe para sacar los datos
    df_out = pd.DataFrame(columns=['Unicode', 'Organizaciones','ID', 'Frentes', 'Descripción', 'Estado'])
    df_out.to_excel(HERE + '\\BDD\\Frentes_try.xlsx', index=False)

    # Hacemos una lista con las organizaciones a descargar
    dispatcher = pd.read_excel(bdd)
    dispatcher = dispatcher.drop_duplicates()
    dispatcher = dispatcher['Organizaciones'].values.tolist()

    barrier = threading.Barrier(threads)
    lock = threading.Lock()
    thread_list = []

    def core_frentes_existentes(driver):
        '''INGRESA CONJUNTO POR CONJUNTO Y OBTIENE LA INFORMACIÓN DE LOS CONJUNTOS Y SU ESTATUS'''
        # Vamos a frentes
        acceder_oracle(driver)
        go_front(driver)

        while True:
            if len(dispatcher) == 0:
                driver.quit()
                organize_firefox()
                break

            # Atendemos cada elemento de la lista
            try:
                lock.acquire()
                org = dispatcher.pop(0)
                lock.release()
            except:
                org = []
            concat = []

            while True:
                while True:
                    # Se revisa que se no haya aparecido un error inesperado de oracle y se pueda ir al menu principal
                    try:
                        swdw(driver, 7, 1, "FrontsTable:ResultsDisplayed:0")
                        go_org(driver, org)
                        break
                    # En caso de no poder detectar la ruta a main menu, reiniciamos el explorador
                    except TimeoutException:
                        go_front(driver)
                        time.sleep(1)

                while True:
                    time.sleep(0.5)
                    # Pasamos a beautifulsoup
                    html = driver.page_source
                    soup = BeautifulSoup(html, 'html.parser')

                    # Encontrar el elemento span con id "FrontsTable"
                    fronts_table = soup.find('span', {'id': 'FrontsTable'})
                    table_x1o = fronts_table.find('table', {'class': 'x1o'})

                    # Encontrar todos los elementos tr dentro del elemento span con id "FrontsTable"
                    trs = table_x1o.find_all('tr')
                    trs.pop(0)

                    # Imprimir los contenidos de los elementos tr encontrados
                    concatena = []
                    for tr in trs:
                        try:
                            # Encontramos todos los td
                            td_list = tr.find_all('td')

                            # A cada TD lo pulimos y adquirimos la info necesaria
                            id_bott_a = td_list[0].find_all('a')
                            id_bott = id_bott_a[0].get('id')
                            front_code = str(td_list[1].text)
                            front_description = td_list[2].text
                            front_state = td_list[3].text
                            proyect = td_list[5].text
                            unicod = proyect + " - " + front_code
                            print(id_bott, front_code, front_description, front_state, proyect)

                            # Juntamos todos los tds filtrados en un dataframe
                            data = {'Unicode': unicod,'Organizaciones': proyect, 'ID': id_bott, 'Frentes': str(front_code), 'Descripción': front_description, 'Estado': front_state}
                            concatena.append(data)

                        except (IndexError, AttributeError):
                            break

                    new_data = pd.DataFrame(concatena)
                    concat.append(new_data)
                    break

                # Guardamos el dataframe nuevo
                df_new = pd.concat(concat)
                lock.acquire()
                df_out = pd.read_excel(HERE + '\\BDD\\Frentes_try.xlsx')
                df_out = pd.concat([df_out, df_new])
                df_out['Frentes'] = df_out['Frentes'].astype(str).str.zfill(2)
                df_out.to_excel(HERE + '\\BDD\\Frentes_try.xlsx', index=False )
                lock.release()
                break    

    for thread in range(threads):
        driver = create_driver(headless=headless)   
        thread = threading.Thread(target=core_frentes_existentes, args=(driver,))
        thread.start()
        thread_list.append(thread)

    for thread in thread_list:
        thread.join()

    os.remove(HERE + '\\BDD\\BDD_Frentes_existentes.xlsx')
    os.rename(HERE + '\\BDD\\Frentes_try.xlsx', HERE + '\\BDD\\BDD_Frentes_existentes.xlsx')
    print('Listo')


def estatus_conjuntos(root, headless=True, threads=6):
    '''INGRESA CONJUNTO POR CONJUNTO Y OBTIENE LA INFORMACIÓN DE LOS CONJUNTOS Y SU ESTATUS'''

    # Separación de organizaciones
    dispatcher = pd.read_excel(root, dtype={'Frentes': object})

    # Crear lista de guardado
    df_conjunto = pd.DataFrame(columns = ['Detalles', 'Conjunto', 'Descripción', 'Estado', 'No de Elementos', 'Creado Por'])
    df_conjunto.to_excel(HERE + '\\BDD\\Conjuntos_try.xlsx', index=False )
    barrier = threading.Barrier(threads)
    lock = threading.Lock()
    thread_list = []
    dispatcher.to_excel(HERE + "\\dispatcher.xlsx", index=False)

    def core_estatus_conjuntos(driver):
        acceder_oracle(driver)
        go_front(driver)

        # Iteramos organizaciones
        while True:
            try:
                # Revisamos el avance
                lock.acquire()
                dispatcher = pd.read_excel(HERE + "\\dispatcher.xlsx", dtype={'Frentes': object})
                lock.release()

                # Cuando se acabe la lista se termina el programa y se sale
                if len(dispatcher) == 0:
                    driver.quit()
                    organize_firefox()
                    break

                # Atendemos el primer elemento de la lista lo extraemos y guardamos la lista
                lock.acquire()
                row = dispatcher.iloc[0]
                dispatcher = dispatcher.iloc[1:]
                dispatcher.to_excel(HERE + "\\dispatcher.xlsx", index=False)
                lock.release()
                
                # Extraemos de la primera fila los elementos necesarios para correr el programa
                IDD = row['ID']
                org = row['Organizaciones']
                SPAN_ID = "FrontsTable:Code:" + IDD.split("Details:", 1)[1]
                frente = row['Frentes'].zfill(2)
                # Creamos un espacio para concatenar todo el frente
                concat = []
                
                # Aseguramos estar en la lista de frentes, en la organización buscada, con frentes expandidos
                while True:
                    try:
                        swdw(driver, 2, 0, "//table[@class='x1o']/tbody[1]/tr[2]/td[1]/a[1]/img[1]")
                        break
                    # En caso de no poder detectar la ruta a main menu, reiniciamos el explorador
                    except (TimeoutException):
                        try:
                            go_org(driver, org)
                        except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException):
                            go_front(driver)
                            go_org(driver, org)

                # Aseguramos estar dentro del frente buscado
                while True:
                    try:
                        swdw(driver, 2, 0, "//h2[text()='Frente']")
                        break
                    except TimeoutException:
                        try:
                            SPAN = swdw(driver, 2, 1, SPAN_ID).text
                            if SPAN == frente:
                                swdw(driver, 8, 1, IDD).click()
                            else:
                                print("Algo salió mal al dar click en el frente")
                                break
                        except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException):
                            time.sleep(0.2)
                            continue

                # Damos click a los conjuntos dentro del frente
                while True:
                    checker = "//h2[text()='Conjuntos de Construcción']"
                    try:
                        swdw(driver, 2, 0, checker)
                        break                    
                    except TimeoutException:
                        swdw(driver, 2, 1, 'SetsLink')  
                        swdw(driver, 2, 1, 'SetsLink').click()

                # Extraemos todos los selectores, para ver cuantas tablas con conjuntos existen
                while True:
                    # Pasamos a BeatifoulSoup
                    html = driver.page_source
                    soup = BeautifulSoup(html, 'html.parser')
                    # Obtenemos todas las páginas con conjuntos para seleccionarlas
                    selector_soup = soup.find('select', {'class': 'x8'})

                    # Si la página está vacia no habrá siquiera selectores, asi que se asigna una bandera
                    if selector_soup is not None:
                        selectors = selector_soup.find_all('option')
                        selection = []
                        for selector in selectors:
                            selection.append(selector.text)
                    #  se asigna la bandera en caso de estar vacia
                    else:
                        selection = [1]

                    # Para cada selector iremos a dicha página
                    for sel in selection:
                        while True:
                            # si la bandera no está, se va a la página; si está, se omite el paso
                            if sel != 1:
                                siguiente = "//select[@title='Seleccionar juego de registros']"
                                checker = "SetsTable:ShowDetails:0"

                                try:
                                    # ubicar el elemento select
                                    time.sleep(0.1)
                                    select_element = driver.find_element(By.XPATH, siguiente)
                                    # crear un objeto Select a partir del elemento select
                                    select = Select(select_element)
                                    # seleccionar la opción por su texto visible
                                    select.select_by_visible_text(sel)
                                    # revisamos que ya esté cargado
                                    swdw(driver, 8, 1, checker)
                                    break
                                except (StaleElementReferenceException, ElementClickInterceptedException, NoSuchElementException):
                                    time.sleep(0.5)
                            else:
                                break

                        # Ya dentro de los conjuntos extraemos todos los elementos de la tabla
                        data_table = beautiful_table(driver)
                        data_table = data_table.rename(columns={'Código': 'Conjunto'})
                        concat.append(data_table)

                    df_new = pd.concat(concat)

                    # Guardamos y salimos
                    lock.acquire()
                    df_out = pd.read_excel(HERE + '\\BDD\\Conjuntos_try.xlsx')
                    df_out = pd.concat([df_out, df_new])
                    df_out.to_excel(HERE + '\\BDD\\Conjuntos_try.xlsx', index=False)
                    ic(len(dispatcher), df_out.shape[0])
                    lock.release()
                    break

                # Regresamos a la página principal para buscar frentes
                while True:
                    try: 
                        swdw(driver, 2, 1, "Search" )
                        break
                    except TimeoutException:
                        try:
                            swdw(driver, 8, 1, "Return").click()
                        except StaleElementReferenceException:
                            swdw(driver, 8, 1, "Return").click()
                        except TimeoutException:
                            go_front(driver)
                            go_org(driver, org)
                            time.sleep(2)
    
            except (StaleElementReferenceException):
                continue

    for thread in range(threads):
        driver = create_driver(headless=headless)
        thread = threading.Thread(target=core_estatus_conjuntos, args=(driver,))
        thread.start()
        thread_list.append(thread)

    for thread in thread_list:
        thread.join()

    os.remove(HERE + '\\BDD\\BDD_Estatus_en_conjuntos.xlsx')
    os.rename(HERE + '\\BDD\\Conjuntos_try.xlsx', HERE + '\\BDD\\BDD_Estatus_en_conjuntos.xlsx')
    os.remove(HERE + "\\dispatcher.xlsx")

    "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\REPORTES\\"
    print('Listo')
    

def generador_de_reportes(headless=True, threads=4, dias=3):
    '''
    Genera reportes en modulo de construcción de acuerdo a la lista y reportes
    '''
    # Lista de rutas de reportes
    ruta_codigo_reportes = get_here(name="codigos_reportes.xlsx", folder="BDD")
    ruta_frentes_existentes = get_here(name="BDD_Frentes_existentes.xlsx", folder="BDD")
    ruta_reportes_faltantes = get_here(name="reportes_faltantes.xlsx", folder="BDD")

    # Definición para actualizar los reportes faltantes de descargar
    def get_reportes_faltantes():
        lock.acquire()
        LISTA_RUTAS_REPORTES = [["JAV_MC_CAO_QRO", "Reportes Generales", "Reporte Estado de Cuenta de Contratos"],
                                ["JAV_MC_CAO_QRO", "Reportes Generales", "Reporte de Finiquito de Obra"],
                                ["JAV_MC_CAO_QRO", "Reportes Generales", "JVR-Reportes Transacciones de Inventario por Frente"]]

        # Revisión de reportes ya descargados y eliminación de los que estén expirados
        try:
            # Cargamos la lista de reportes ya guardados
            ruta_codigo_reportes = get_here(name="codigos_reportes.xlsx", folder="BDD")
            reportes_generados = pd.read_excel(ruta_codigo_reportes, dtype={'Frentes': object})
            # Rectificamos los valores del dataframe a modo fecha
            reportes_generados['Fechas'] = pd.to_datetime(reportes_generados['Fechas'])
            # Obtenemos la fecha actual y generamos la diferencia de la fecha del reporte creado con la de hoy.
            fecha_actual = dt.now()
            diferencia_en_días = (fecha_actual - reportes_generados['Fechas']).dt.days
            # Eliminamos todos los datos que superen los días establecidos cómo máximos
            reportes_generados = reportes_generados[diferencia_en_días <= dias].reset_index(drop=True)
            # Guardamos los reportes a conservar
            reportes_generados.to_excel(ruta_codigo_reportes, index=False)
        except FileNotFoundError:
            print("no se encontró")
            reportes_generados = pd.DataFrame(columns=['Reportes', 'Unicode', 'Organizaciones', 'Frentes', 'Cod_reportes', 'Códigos', 'Fechas'])


        # Obtenemos el dataframe con todos los frentes
        frentes_dataframe = pd.read_excel(ruta_frentes_existentes, dtype={'Frentes': object})

        # Creamos un dataframe con todos los reportes a descargar y sus organizaciones y frentes
        lista_reportes = []
        for ruta in LISTA_RUTAS_REPORTES:
            reporte = ruta[-1]
            # De los reportes ya generados filtramos de acuerdo al reporte actual 
            filtro_1 = reportes_generados['Reportes'] == reporte
            reportes_coincidentes = reportes_generados.loc[filtro_1]
            # Filtrar frentes_dataframe según los valores de Unicode no existentes en reportes_coincidentes
            filtro_2 = ~frentes_dataframe['Unicode'].isin(reportes_coincidentes['Unicode'])
            faltantes = frentes_dataframe.loc[filtro_2].copy()
            if len(faltantes) != 0:
                # Agregamos el valor de la ruta
                faltantes.loc[:, 'Cod_reportes'] = [ruta] * len(faltantes)
                # Agregamos el valor del reporte
                faltantes.loc[:, 'Reportes'] = reporte
                # Guardamos para concatenar df
            lista_reportes.append(faltantes)
        # Concatenamos dataframes
        df = pd.concat(lista_reportes)
        # Quitamos columnas sobrantes
        df.drop(columns=['ID', 'Descripción', 'Estado'], inplace=True)
        # Agregamos colmnas de Códigos y fechas
        if len(df) != 0:
            df.loc[:, 'Códigos'] = 'Falta'
            df.loc[:, 'Fechas'] = ''
        df = pd.concat([reportes_generados, df])
        df.to_excel(ruta_codigo_reportes, index=False)
        df = df.loc[df['Códigos'] == 'Falta']
        df.to_excel(ruta_reportes_faltantes, index=False)
        lock.release()
        return df

    def update_faltantes():
        lock.acquire()
        reportes_faltantes = pd.read_excel(ruta_reportes_faltantes, dtype={'Frentes': object})

        try:
            row = reportes_faltantes.iloc[0]
            reportes_faltantes.drop(reportes_faltantes.index[0], inplace=True)
            org = row['Organizaciones']
            frente = str(row['Frentes']).zfill(2)
            reporte_str = row['Cod_reportes']
            reporte = ast.literal_eval(reporte_str)
            unicod = row['Unicode']

        except IndexError:
            row = []
            org = ''
            frente = ''
            reporte = ''
            unicod = ''

        reportes_faltantes.to_excel(ruta_reportes_faltantes, index=False)   
        lock.release()

        # Generación de organizaciones generales para ir una por una en el dataframe
        return org, frente, reporte, unicod, row, reportes_faltantes

    def send_and_check(driver, checker, send, to_send):
        # Hacemos loop hasta asegurarnos que se ingresaron correctamente proyecto y frente
        while True: #PROYECTO
            try:
                swdw(driver, 2, 1, "Fndcpparamlink")
            except TimeoutException:
                break

            try:
                # Limpiamos lo que pueda estár en el campo proyecto y mandamos el texto
                swdw(driver, 2, 1, to_send).clear()
                swdw(driver, 2, 1, to_send).send_keys(send + Keys.TAB)
                # El comprobador demuestra que se ingreso el dato de forma correcta
                xpath_comprobador = checker
                comprobador = swdw(driver, 8, 0, xpath_comprobador).get_attribute('class')

                if comprobador == "x2v":
                    break

            except TimeoutException:
                continue

    def navegador_de_reportes(driver, org, frente, reporte):

        # Definiciones XPATH de botones
        CONTINUAR = "//table[@id='CPTrainFooterRG']/tbody[1]/tr[1]/td[8]/button[1]"
        EJECUTAR = "//table[@id='CPTrainFooterRG']/tbody[1]/tr[1]/td[10]/button[1]"
        ACEPTAR = "(//table[@class='x6w']/following::table)[9]/tbody[1]/tr[1]/td[2]/button[1]"

        # Definiciones ID de busquedas
        PROYECTO = "N330"
        FRENTE = "N331"
        TASA_I = "N333"

        # En caso de no poder detectar la ruta a main menu, reiniciamos el explorador
        try:
            swdw(driver, 2, 0, "//table[@class='x6w']/tbody[1]/tr[1]/td[3]/a[1]")
        except (StaleElementReferenceException, TimeoutException) as error:
            print("Stale or timeout", error)
            acceder_oracle(driver)

        # Acceder al reporte y el vaciado de datos
        new_mainmenu(driver, reporte)
        swdw(driver, 8, 0, CONTINUAR).click()

        # Definiciones de datos
        xpath_programa = "//table[@id='Fndcpprogramnamedisplay__xc_']/tbody[1]/tr[1]/td[3]/span[1]"
        nombre_programa = swdw(driver, 8, 0, xpath_programa).text

        # Hacemos loop hasta asegurarnos que se ingresaron correctamente proyecto y frente
        xpath_check_proyecto = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[2]/td[2]/span[1]"
        send_and_check(driver, xpath_check_proyecto, org, PROYECTO)

        xpath_check_frentes = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[4]/td[2]/span[1]"
        send_and_check(driver, xpath_check_frentes, str(frente), FRENTE)

        # si el programa es finiquito de obra, requiere de iva 16 por lo que es un dato extra para el reporte
        if nombre_programa == "XXMCAN - Finiquito de Obra":
            xpath_16 = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[4]/td[2]/span[1]"
            send_and_check(driver, xpath_16, "16", TASA_I)

        # Aseguramos llegar a ejecutar
        while True:
            try:
                swdw(driver, 2, 0, CONTINUAR).click()
                swdw(driver, 8, 0, EJECUTAR)
                break
            except (TimeoutException, StaleElementReferenceException) as error:
                print(error)
                continue

        lock.acquire()
        # Extraemos el ID de descarga y la fecha
        while True:
            try:     
                swdw(driver, 8, 0, EJECUTAR).click()
                INFORM = "(//div[@class='x79']//table)[2]/tbody[1]/tr[1]/td[1]/span[1]"
                id_descarga = int(re.search(r'tud es (.*?)$', swdw(driver, 8, 0, INFORM).text).group(1))
                now = dt.now()
                swdw(driver, 2, 0, ACEPTAR).click()
                break
            except (TimeoutException, StaleElementReferenceException) as error:
                print(error)
                continue
        lock.release() 
        swdw(driver, 0.5, 0, "//table[@class='x6w']/tbody[1]/tr[1]/td[3]/a[1]").click()
        return id_descarga, now


    def generador(driver):
        acceder_oracle(driver)
        # barrier.wait()
        while True:
            # Abrimos un try para permitir se intente varias veces hasta que se tenga éxito en el resultado
            try:
                # Obtenemos fila y actualizamos reportes faltantes
                org, frente, reporte, unicod, row, reportes_faltantes = update_faltantes()
                print(org, frente, reporte, unicod)
                # Accedemos a modulo de construcción y obtenemos la ruta de los reportes
                if len(reportes_faltantes) == 0:
                    reportes_faltantes = get_reportes_faltantes()
                    if len(reportes_faltantes) == 0:
                        driver.quit()
                        organize_firefox()
                        break

                while True:
                    try:
                        print(reporte[-1], org, frente)
                        id_descarga, now = navegador_de_reportes(driver, org, frente, reporte)
                        print(org, frente, reporte, id_descarga, now)
                        break
                    except (TimeoutException, StaleElementReferenceException) as error:
                        print("Error de navegador de reportes", error)
                        continue

                # Añadimos los datos a un dataframe y lo guardamos en xlsx
                lock.acquire()
                reportes_viejos = pd.read_excel(ruta_codigo_reportes, dtype={'Frentes': object})
                reportes_viejos.loc[(reportes_viejos['Unicode'] == unicod) & (reportes_viejos['Reportes'] == reporte[-1]), ['Códigos', 'Fechas']] = [id_descarga, now]
                reportes_viejos.to_excel(ruta_codigo_reportes, index=False)
                lock.release()
            except Exception as e:
                print(e)
                continue
            except NoSuchWindowException:
                driver = reset_driver(driver)

    # Main program
    barrier = threading.Barrier(threads)
    lock = threading.Lock()
    thread_list = []
    reportes_faltantes = get_reportes_faltantes()
    print(reportes_faltantes)

    if threads >= len(reportes_faltantes):
        threads = len(reportes_faltantes)

    if threads != 0:
        for x in range(threads):
            driver = create_driver(headless=headless)
            thread = threading.Thread(target=generador, args=(driver,))
            thread.start()
            thread_list.append(thread)

        for thread in thread_list:
            thread.join()
    print('Éxito')

#------------------------------------------------------------------


def clean_archivos_reportes(días=3):
    folder = get_here(folder = "partial_files")
    archivos = glob.glob(os.path.join(folder, 'R*.xlsx'))
    hoy = dt.now()
    fecha_limite = hoy - datetime.timedelta(days=hoy.weekday() + días)
    for archivo in archivos:
        fecha_creacion = dt.fromtimestamp(os.path.getctime(archivo))

        if fecha_creacion <= fecha_limite:
            try:
                os.remove(archivo)
                print(fecha_limite, "<=", fecha_creacion, "remove", archivo)
            except PermissionError as e:
                print("PermissionError", e)
            except Exception as e:
                print(e)


def descargador_archivos_reportes(threads=4, headless=True):

    def check_downloads(df_in):
        folder = get_here(folder = "partial_files")
        archivos = glob.glob(os.path.join(folder, 'R*.xlsx'))

        reporte_lista = []
        organización_lista = []
        frente_lista = []

        for archivo in archivos:
            directorio, nombre_completo  = os.path.split(archivo)
            nombre, extension = os.path.splitext(nombre_completo)
            reporte_base, organización, frente = nombre.split(" - ")

            if reporte_base == 'RDF':
                reporte = 'Reporte de Finiquito de Obra'
            elif reporte_base == 'RTI':
                reporte = 'JVR-Reportes Transacciones de Inventario por Frente'
            elif reporte_base == 'REC':
                reporte = 'Reporte Estado de Cuenta de Contratos'

            reporte_lista.append(reporte)
            organización_lista.append(organización)
            frente_lista.append(frente)
        
        df_list = pd.DataFrame({
                'Reportes': reporte_lista,
                'Organizaciones': organización_lista,
                'Frentes': frente_lista
            })

        df_list['File'] = df_list['Reportes'] + ' - ' + df_list['Organizaciones'] + ' - ' + df_list['Frentes']

        # Realizar una fusión basada en las columnas coincidentes
        merged_df = df_in.merge(df_list, on=['Reportes', 'Organizaciones', 'Frentes'], how='left', indicator=True)

        # Seleccionar las filas que no tienen coincidencias
        df_out = merged_df[merged_df['_merge'] == 'left_only'].drop('_merge', axis=1)
        df_out.to_excel(folder + "\\lista_descarga_faltantes.xlsx", index=False)
        return df_out

    # Despachador de trabajo
    def take_first_and_save():
        lock.acquire()
        df = pd.read_excel(HERE + "\\flash_memory.xlsx", dtype={'Frentes': object})
        df.sort_values(by="Códigos", ascending=False, inplace=True)
        try:
            row = df.iloc[0]
            df.drop(df.index[0], inplace=True)
        except IndexError:
            row = []
        df.to_excel(HERE + "\\flash_memory.xlsx", index=False)
        lock.release()
        return row, df

    def core_archivos_reportes(driver, barrier2):
        # Definición de botones
        boton_siguiente = "//table[@class='x1p']//table/tbody[1]/tr[1]/td[7]/a[1]"
        boton_siguiente_inactivo = "//table[@class='x1p']//table/tbody[1]/tr[1]/td[7]/span[1]"
        boton_anterior = "//table[@class='x1p']//table/tbody[1]/tr[1]/td[3]/a[1]"
        boton_anterior_inactivo = "//table[@class='x1p']//table/tbody[1]/tr[1]/td[3]/span[1]"

        # Se hace bucle infinito hasta que la lista llegue a 0
        # barrier2.wait()
        while True:
            try:
                row, pendientes = take_first_and_save()
                # Condición de salida para térimno de programa
                if len(row) == 0:
                    barrier.wait()
                    time.sleep(2)
                    break

                # Extraemos el dato del row
                codigo = row['Códigos']
                unicod = row['Unicode']
                reporte = row['Reportes']

                # Aseguramos ir al monitor
                go_monitor(driver)

                while True:
                    # Usamos Beautiful Soup para extraer una tabla en dataframe
                    table = beautiful_table(driver, element="class", name="x1o")
                    table['ID de Solicitud'] = table['ID de Solicitud'].astype(int)
                    lista_codigos = table['ID de Solicitud'].values

                    if codigo in lista_codigos:
                        table = beautiful_table(driver, element="class", name="x1o")
                        table['ID de Solicitud'] = table['ID de Solicitud'].astype(int)

                        try:
                            output_id = table['Output'].loc[table['ID de Solicitud'] == codigo].values[0]
                            Fase = table['Estado'].loc[table['ID de Solicitud'] == codigo].values[0]

                            if Fase == "ERROR:":
                                lock.acquire()
                                print("Error en fase", Fase)
                                ruta_codigo_reportes = get_here(name="codigos_reportes.xlsx", folder="BDD")
                                reportes_generados = pd.read_excel(ruta_codigo_reportes, dtype={'Frentes': object})
                                reportes_generados = reportes_generados.drop(reportes_generados.loc[reportes_generados['Códigos'] == codigo].index, axis=0)
                                reportes_generados.to_excel(ruta_codigo_reportes, index=False)
                                generador_de_reportes(headless=True, threads=1, dias=1)
                                reportes_generados = pd.read_excel(ruta_codigo_reportes, dtype={'Frentes': object})
                                reportes_reporte = reportes_generados.loc[reportes_generados['Reportes'] == reporte]
                                codigo = reportes_reporte.loc[reportes_reporte['Unicode'] == unicod, 'Códigos']
                                lock.release()
                                row, pendientes = take_first_and_save()
                                continue
                            else: 
                                while True:
                                    try:
                                        swdw(driver, 2, 1, output_id).click()
                                        time.sleep(1)
                                        print(reporte, unicod, codigo, output_id, "descargado; Faltan:", len(pendientes))
                                        break
                                    except (TimeoutException, StaleElementReferenceException) as error:
                                        lock.acquire()
                                        df = pd.read_excel(HERE + "\\flash_memory.xlsx", dtype={'Frentes': object})
                                        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                                        df.to_excel(HERE + "\\flash_memory.xlsx", index=False)
                                        lock.release()
                                        print("error en ", codigo, reporte, unicod, Fase, error)
                                        print(df.iloc[-1])
                                        go_monitor(driver)
                                        continue
                            break

                        except IndexError:
                            print('Error al descargar: ', codigo, reporte, unicod)
                            lock.acquire()
                            df = pd.read_excel(HERE + "\\flash_memory.xlsx", dtype={'Frentes': object})
                            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                            df.to_excel(HERE + "\\flash_memory.xlsx", index=False)
                            lock.release()
                            print(df.iloc[-1])


                    elif codigo < min(lista_codigos):
                        while True:
                            try:
                                swdw(driver, 2, 0, boton_siguiente).click()
                                break
                            except TimeoutException:
                                print("no hay siguinete")
                            except StaleElementReferenceException:
                                time.sleep(0.5)
                        continue

                    elif codigo > max(lista_codigos):
                        while True:
                            try:
                                swdw(driver, 1, 0, boton_anterior).click()
                                break
                            except TimeoutException:
                                print("no hay anteriores")
                            except StaleElementReferenceException:
                                time.sleep(0.5)
                        continue
        
            except Exception as error:
                print("What?", error)

            except NoSuchWindowException:
                driver = reset_driver(driver)

        barrier.wait()
        time.sleep(1)
        driver.quit()
        organize_firefox()
        

    # Dataframe de reportes a descargar
    root_codigo_reportes = get_here(folder="BDD", name="codigos_reportes.xlsx")
    root_flash_memory = get_here(name="flash_memory.xlsx")
    root_proob = get_here(name="prueba.xlsx")

    base_df = pd.read_excel(root_codigo_reportes, dtype={'Frentes': object})

    try:
        df = check_downloads(base_df)
        df.sort_values(by="Códigos", ascending=False, inplace=True)
    except:
        df = base_df


    for index, row in df.iterrows():
        print("Reporte a descargar: ", row['Reportes'], row['Unicode'])
    df.to_excel(root_flash_memory, index=False)
    df.to_excel(root_proob)

    # Definiciones thread
    barrier = threading.Barrier(threads)
    lock = threading.Lock()
    thread_list = []

    try:
        if threads != 0:
            for thread in range(threads):
                driver = create_driver(headless=headless)
                thread = threading.Thread(target=core_archivos_reportes, args=(driver, barrier,))
                thread.start()
                thread_list.append(thread)
            for thread in thread_list:
                thread.join()
            driver.quit()
            organize_firefox()
    except Exception as e:
        print(f'Error ocurrido en: {e}')
        logging.error(f"Error: {e}")
        logging.error(f"Stack trace: {traceback.format_exc()}")
    print('Éxito')


def rename_reports(dias=2):
    folder = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\"
    archivos_rdf = glob.glob(os.path.join(folder, 'XXMCAN___Finiquito_de_Obra*.xls'))
    archivos_rti = glob.glob(os.path.join(folder, 'XXMC_Reportes_Transacciones_de_*.xls'))
    archivos_rdp = glob.glob(os.path.join(folder, 'XXMCAN__Reporte_de_Penaliza*.xls'))
    archivos_rec = glob.glob(os.path.join(folder, 'XXMCAN__Reporte_Estado_de_Cuen*.xls'))
    all_files = archivos_rti + archivos_rdf + archivos_rdp + archivos_rec
    
    rout = os.path.join(HERE, "partial_files")
    # Obtenemos la fecha actual
    fecha_actual = dt.now()

    # Recorremos todos los archivos de la carpeta
    for archivo in os.listdir(rout):
        # Obtenemos la ruta del archivo
        ruta_archivo = os.path.join(rout, archivo)

        # Comprobamos si el archivo es un xlsx
        if archivo.endswith(".xlsx"):
            # Obtenemos la fecha de creación del archivo
            fecha_creacion = os.path.getctime(ruta_archivo)

            fecha_actual = dt.now()
            fecha_creacion = dt.fromtimestamp(fecha_creacion)

            # Calculamos la diferencia de tiempo entre la fecha de creación del archivo y la fecha actual
            diferencia_tiempo = fecha_actual - fecha_creacion

            # Si la diferencia de tiempo es mayor de 5 días, eliminamos el archivo
            if diferencia_tiempo.days > dias:
                os.remove(ruta_archivo)

    for file in archivos_rdf:
        try:
            name_df = pd.read_html(file, header=0)[0]
            df = pd.read_html(file, header=0)[1]
            organización = name_df.iloc[0, 1]
            frente = name_df.iloc[1, 1]
            new_name = "RDF - " + organización[:3] + " - " + frente[:2] + ".xlsx"
            print(new_name)
            new_rout = os.path.join(rout, new_name)
            df.to_excel(new_rout, index=False)
        except:
            os.remove(file)

    for file in archivos_rec:
        try:
            base_df = pd.read_excel(file, header=1, dtype={'Frente': object})
            df = base_df[base_df['Contrato'].notna() & (base_df['Contrato'].str.strip() != '') & base_df['Frente'].notna()]
            organización = df['No. Proyecto'].iloc[1]
            frente = df['Frente'].iloc[0]
            new_name = "REC - " + organización[:3] + " - " + frente[:2] + ".xlsx"
            new_rout = os.path.join(rout, new_name)
            df.to_excel(new_rout, index=False)
        except Exception as e:
            print(e)
            os.remove(file)

    for file in archivos_rti:
        try:
            name_df = pd.read_html(file, header=0)[0]
            try:
                organización = name_df['Unidad Operativa'].iloc[0]
                frente = name_df['Segment9'].iloc[0]
                new_name = "RTI - " + organización[:3] + " - " + frente[1:3] + ".xlsx"
                print(new_name)
                new_rout = os.path.join(rout, new_name)
                name_df.to_excel(new_rout, index=False)
            except IndexError as error:
                print(os.path.split(file)[-1], error)
            except AttributeError as error:
                print(error)
        except:
            os.remove(file)

    for file in archivos_rdp:
        try:
            df = pd.read_excel(file, header=1, dtype={'Frente': object})
            print(df)
            organización = df['No. Proyecto'][0]
            frente = df['Frente'][0]
            print(organización, frente)
            new_name = "RDP - " + organización[:3] + " - " + frente[:2] + ".xlsx"
            print(new_name)
            new_rout = os.path.join(folder, new_name)
            df.to_excel(new_rout, index=False)
        except Exception as e:
            root, nombre = os.path.split(file)
            print(e)
            print("no se pudo", nombre)

    for file in all_files:
        try:
            os.remove(file)
        except FileNotFoundError:
            continue


#----------------------------Compilador de archivos--------------------------------------

def get_purge_bdd(end='EC'):
    archivos = os.listdir(HERE + '\\BDD\\CRYSTAL\\')
    nombres_transformados = []
    # Itera sobre la lista de archivos
    for root, dirs, files in os.walk(HERE + "\\BDD\\CRYSTAL\\"):
        for file in files:
            try:
                ruta_archivo = os.path.join(root, file)
                fecha_hoy = time.time()
                fecha_archivo = os.path.getmtime(ruta_archivo)
                diferencia_en_días = (fecha_hoy - fecha_archivo) / (24 * 60 * 60)
                # Verificamos cumpla el patrón de nombre
                if file.lower().endswith(end.lower() + '.xlsx'):
                    nombres_transformados.append(file[:3])
                    if not file.lower().startswith(u) and diferencia_en_días > 1:
                        os.remove(ruta_archivo)
            except:
                continue 
    return(nombres_transformados)


def compile_bdd_xlsx(name='FE'):

    try:
        os.remove(HERE + "\\BDD\\BDD_" + name + ".xlsx")
    except:
        time.sleep(0)

    # Definimos el nombre a buscar
    end_name = name.lower()
    concat = [] # Lista para guardar los df
    names = []

    # Buscamos en todos los archivosde la dirección
    for root, dirs, files in os.walk(HERE + "\\BDD\\CRYSTAL\\"):
        for file in files:
            # Verificamos cumpla el patrón de nombre
            if file.lower().startswith(end_name) and file.lower().endswith('.xlsx'):
                df_new = pd.read_excel(os.path.join(root, file), dtype={'Frentes': object}) # Dataframe con la ruta del archivo
                concat.append(df_new) # Guardamos el dataframe en la lista
    # Concatenamos y guardamos
    try:
        df_concat = pd.concat(concat)
        df_concat['Frentes'] = df_concat['Frentes'].astype(str)
    except ValueError:
        print("Error: No hay objetos para concatenar")
    # df_concat.dropna(subset=[df_concat[1]], inplace=True)
    if name == "EC":
        save_name = "Estatus_en_conjuntos"
    elif name == "FE":
        save_name = "Frentes_existentes"
    df_concat.to_excel(HERE + "\\BDD\\BDD_" + save_name + ".xlsx", index=False)


def GENERADOR_DE_REPORTES(driver, reporte, frentes_dataframe, dias):
    '''
    Genera reportes en modulo de construcción de acuerdo a la lista y reportes

    Args: 
        driver: el navegador de Selenium a usar
        reporte: el nombre del reporte a enrutar
        frentes_dataframe: un dataframe con organizaciones y frentes a descargar, 
                con columnas: Organizaciones y Frentes

    Returns:
        dataframe de códigos'''

    # Revisión de reportes ya descargados vs por descargar
    try:
        reportes_generados = pd.read_excel(HERE + "\\BDD\\codigos_reportes.xlsx", dtype={'Frentes': object})
        # Revisamos cuales tienen mas de un día de existencia y los descartamos
        fecha_actual = dt.now()
        diferencia_en_días = (fecha_actual - reportes_generados['Fechas']).dt.days
        reportes_generados = reportes_generados[diferencia_en_días <= dias]
        reportes_generados.to_excel(HERE + "\\BDD\\codigos_reportes.xlsx")
    except FileNotFoundError:
        print("no se encontró")
        reportes_generados = pd.DataFrame(columns = ['Reportes', 'Unicode', 'Organizaciones', 'Frentes', 'Cod_reportes', 'Códigos', 'Fechas']) # Organización, frente y código de reporte

    # Revisamos cuales reportes coinciden con el reporte
    filtro_1 = reportes_generados['Reportes'] == reporte[-1]
    reportes_coincidentes = reportes_generados.loc[filtro_1]

    # Filtrar frentes_dataframe según los valores de Unicode no existentes en reportes_coincidentes
    filtro_2 = ~frentes_dataframe['Unicode'].isin(reportes_coincidentes['Unicode'])
    reportes_faltantes = frentes_dataframe.loc[filtro_2]

    print("reportes faltantes", len(reportes_faltantes), reportes_faltantes['Unicode'])

    # Accedemos a modulo de construcción y obtenemos la ruta de los reportes
    acceder_oracle(driver)

    # Generación de organizaciones generales para ir una por una en el dataframe
    organizaciones = reportes_faltantes['Organizaciones'].drop_duplicates()
    if len(reportes_faltantes) != 0:
        for org in organizaciones:

            # Generación de los frentes de cada organizcaión para ir uno por uno
            frentes = reportes_faltantes.loc[reportes_faltantes['Organizaciones'] == org, 'Frentes']
            for frente in frentes:
                while True:

                    # Se revisa que se no haya aparecido un error inesperado de oracle y se pueda ir al menu principal
                    try:
                        swdw(driver, 8, 0, "//table[@class='x6w']/tbody[1]/tr[1]/td[3]/a[1]")
                    # En caso de no poder detectar la ruta a main menu, reiniciamos el explorador
                    except (StaleElementReferenceException, TimeoutException):
                        acceder_oracle(driver)
                        continue

                    # Definiciones XPATH de botones
                    CONTINUAR = "//table[@id='CPTrainFooterRG']/tbody[1]/tr[1]/td[8]/button[1]"
                    EJECUTAR = "//table[@id='CPTrainFooterRG']/tbody[1]/tr[1]/td[10]/button[1]"
                    ACEPTAR = "(//table[@class='x6w']/following::table)[9]/tbody[1]/tr[1]/td[2]/button[1]"

                    # Definiciones ID de busquedas
                    PROYECTO = "N330"
                    FRENTE = "N331"
                    TASA_I = "N333"
                    
                    try:
                        # Acceder al reporte y el vaciado de datos
                        new_mainmenu(driver, reporte)
                        swdw(driver, 8, 0, CONTINUAR).click()

                        # Definiciones de datos
                        xpath_programa = "//table[@id='Fndcpprogramnamedisplay__xc_']/tbody[1]/tr[1]/td[3]/span[1]"
                        nombre_programa = swdw(driver, 8, 0, xpath_programa).text

                        # Hacemos loop hasta asegurarnos que se ingresaron correctamente proyecto y frente
                        while True: #PROYECTO
                            try:
                                # el checker permite dar certeza de que estamos en la página correcta
                                swdw(driver, 2, 1, "Fndcpparamlink")
                            except TimeoutException:
                                break
                            try:
                                # Limpeamos lo que pueda estár en el campo proyecto y mandamos el texto
                                swdw(driver, 2, 1, PROYECTO).clear()
                                swdw(driver, 2, 1, PROYECTO).send_keys(org + Keys.TAB)
                                # El comprobador demuestra que se ingreso el dato de forma correcta
                                xpath_comprobador = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[2]/td[2]/span[1]"
                                comprobador = swdw(driver, 8, 0, xpath_comprobador).get_attribute('class')
                                if comprobador == "x2v":
                                    break
                            except TimeoutException:
                                continue

                        while True: #FRENTE
                            try:
                                # el checker permite dar certeza de que estamos en la página correcta
                                swdw(driver, 2, 1, "Fndcpparamlink")
                            except TimeoutException:
                                break
                            try:
                                swdw(driver, 2, 1, FRENTE).clear()
                                swdw(driver, 2, 1, FRENTE).send_keys(str(frente) + Keys.TAB)
                                xpath_comprobador = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[4]/td[2]/span[1]"
                                comprobador = swdw(driver, 8, 0, xpath_comprobador).get_attribute('class')
                                if comprobador == "x2v":
                                    break
                            except TimeoutException:
                                continue

                        # si el programa es finiquito de obra, requiere de iva 16 por lo que es un dato extra para el reporte
                        if nombre_programa == "XXMCAN - Finiquito de Obra":
                            while True:
                                try:
                                # el checker permite dar certeza de que estamos en la página correcta
                                    swdw(driver, 2, 1, "Fndcpparamlink")
                                except TimeoutException:
                                    break
                                # ingresamos el iva requerido
                                try:
                                    swdw(driver, 2, 1, TASA_I).clear()
                                    swdw(driver, 2, 1, TASA_I).send_keys("16" + Keys.TAB)
                                    xpath_comprobador = "(//span[@id='Fndcpparamregion']//table)[2]/tbody[1]/tr[4]/td[2]/span[1]"
                                    comprobador = swdw(driver, 8, 0, xpath_comprobador).get_attribute('class')
                                    if comprobador == "x2v":
                                        break
                                except TimeoutException:
                                    continue

                        # Ejecutamos el reporte y extraemos el código de descarga
                        try:
                            swdw(driver, 2, 0, CONTINUAR).click()
                            swdw(driver, 8, 0, EJECUTAR).click()
                            INFORM = "(//div[@class='x79']//table)[2]/tbody[1]/tr[1]/td[1]/span[1]"
                            id_descarga = int(re.search(r'tud es (.*?)$', swdw(driver, 8, 0, INFORM).text).group(1))
                            now = dt.now()
                            print(org, frente, reporte[-1], id_descarga, now)
                            swdw(driver, 2, 0, ACEPTAR).click()
                            break
                        except (TimeoutException, StaleElementReferenceException):
                            continue
                    except (TimeoutException, StaleElementReferenceException):
                        continue

                # Añadimos los datos a un dataframe y lo guardamos en xlsx
                fila = {'Reportes': reporte[-1], 'Unicode': org + " - " + str(frente), 'Organizaciones': org, 'Frentes': str(frente), 'Cod_reportes': reporte,'Códigos': id_descarga, 'Fechas': now}
                concat_df = pd.DataFrame([fila])
                reportes_generados = pd.concat([reportes_generados, concat_df], ignore_index=True)
                reportes_generados.to_excel(HERE + "\\BDD\\codigos_reportes.xlsx", index=False)
                print(len(reportes_generados))


def create_document_contract(driver, conjunto, documento):

    # Separamos el conjunto en org, frente y conjunto
    org, frente, conjunto = convertir_conjunto(conjunto)
    acceder_oracle(driver)
    go_contract(driver)
    lookup_contract(driver, org, frente, conjunto)


def rename_contracts_purchaser(orgs, purchaser, headless=False):
    driver = create_driver(headless=headless)
    file_bdd = get_here(name = "BDD_Reporte_de_Finiquitos.xlsx", folder = "BDD")
    df_all_sets = pd.read_excel(file_bdd)
    df_sets = df_all_sets[df_all_sets['No. Proyecto'].isin(orgs)]
    df_sets = df_sets[~df_sets['Estado_Conjunto'].isin(['', 'Cerrado'])]
    df_sets = df_sets[~df_sets['Comprador'].isin([purchaser])]
    df_sets = df_sets[~df_sets['Estado'].isin(['Cancelado', 'En Proceso de Definición'])]

    df_sets = df_sets.sort_values(by='Conjunto', ascending=False)

    for i, row in df_sets.iterrows():
        try:
            conjunto = row['Conjunto']
            contrato = row['Contrato']
            get_contract(driver, conjunto, contrato)

            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')    
            actual_purchaser = soup.find(id='Agent').get_text()
            ic(i, len(df_sets))

            if actual_purchaser != purchaser:
                try:
                    swdw(driver, 1, 1, "SetAgent").click()
                    css = '#AgentLOV__xc_0 > a:nth-child(3)'
                    stf(driver, 2, css, purchaser)
                    time.sleep(1)
                    swdw(driver, 2, 1, "Apply").click()

                except Exception as e:
                    print(e)
                    continue

            swdw(driver, 2, 1, "Return").click()
        except Exception as e:
            print(e)
            continue

def multithread_rename_contracts_purchaser(orgs, purchaser, headless=False, threads=4):

    def core(driver, purchaser):
        while True:
            try:
                lock.acquire()
                row, pendientes = take_first_row()
                lock.release()
                # Condición de salida para térimno de programa
                if len(row) == 0:
                    barrier.wait()
                    time.sleep(2)
                    break

                conjunto = row['Conjunto']
                contrato = row['Contrato']
                get_contract(driver, conjunto, contrato)


                html = driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                try:  
                    actual_purchaser = soup.find(id='Agent').get_text()
                except AttributeError as e:
                    print(e)
                    continue

                if actual_purchaser != purchaser:
                    try:
                        swdw(driver, 1, 1, "SetAgent").click()
                        css = '#AgentLOV__xc_0 > a:nth-child(3)'
                        stf(driver, 2, css, purchaser)
                        time.sleep(1)
                        swdw(driver, 2, 1, "Apply").click()

                    except Exception as e:
                        print(e)
                        continue

                swdw(driver, 2, 1, "Return").click()

            except Exception as e:
                print(e)

                lock.acquire()
                flash_memory = get_here(folder='cache', name='flash_memory_mth.xlsx')
                df = pd.read_excel(flash_memory, dtype={'Frentes': object})
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                df.to_excel(flash_memory, index=False)
                lock.release()

                logging.error(f"Error: {e}")
                logging.error(f"Stack trace: {traceback.format_exc()}")
                continue

            except NoSuchWindowException:
                print("\nLa ventana no responde\n")
                driver = reset_driver(driver, headless=headless)
                go_contract(driver)

    file_bdd = get_here(name = "BDD_Reporte_de_Finiquitos.xlsx", folder = "BDD")
    df_all_sets = pd.read_excel(file_bdd)
    df_sets = df_all_sets[df_all_sets['No. Proyecto'].isin(orgs)]
    df_sets = df_sets[~df_sets['Estado_Conjunto'].isin(['', 'Cerrado'])]
    df_sets = df_sets[~df_sets['Comprador'].isin([purchaser])]
    df_sets = df_sets[~df_sets['Estado'].isin(['Cancelado', 'En Proceso de Definición'])]
    df = df_sets.sort_values(by='Conjunto', ascending=True)
    df = df_sets.sort_values(by='Comprador', ascending=True)
    df = df_sets.sort_values(by='Estado', ascending=False)
    lenght_df = len(df)

    if lenght_df < threads:
        threads = lenght_df

    flash_memory = get_here(folder='cache', name='flash_memory_mth.xlsx')
    lock = threading.Lock()
    df.to_excel(flash_memory, index=False)
    thread_list = []

    if threads != 0:
        for thread in range(threads):
            driver = create_driver(headless=headless)
            thread = threading.Thread(target=core, args=(driver, purchaser, ))
            thread.start()
            thread_list.append(thread)
        for thread in thread_list:
            thread.join()


# apagar_pc()
# Checar integridad de datos, con verificación de frentes y llevar un registro de filas.
# Revisar que el monto del contrato corresponda con el estimado y el por estimar.
