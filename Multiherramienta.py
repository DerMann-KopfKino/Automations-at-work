import os
import re
import time
import math
import threading
import configparser
import pyautogui
import numpy as np
import pandas as pd
import tkinter as tk
from icecream import ic
import pygetwindow as gw
from bs4 import BeautifulSoup
import screeninfo
from pathlib import Path
from tkinter import messagebox
from selenium import webdriver
from fuzzywuzzy import fuzz, process
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import FirefoxOptions, ChromeOptions, EdgeOptions
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import (TimeoutException, ElementClickInterceptedException, 
    NoSuchElementException, UnexpectedAlertPresentException, NoSuchFrameException, NoSuchWindowException,
    InvalidArgumentException, StaleElementReferenceException, UnexpectedAlertPresentException, NoAlertPresentException)


#----------------------------D e f i n i c i o n e s-------------------------------------

def apagar_pc():
    os.system("shutdown /s /t 0")

class Mi_ID():
    USER = ''
    PASSWORD = ''

ORACLE_ID = Mi_ID()
ORACLE_ID.USER = "FPRADO"
ORACLE_ID.PASSWORD = "PAMFYjaver12"

ORACLECLOUD = Mi_ID()
ORACLE_ID.USER = "FPRADO"
ORACLECLOUD.PASSWORD = "oprmt141592FRA+-d"

JAVER_ID = Mi_ID()
JAVER_ID.USER = "fprado"
JAVER_ID.PASSWORD = "pamf900509HFA11"


def convertir_numero_a_texto(numero):
    # Definición de listas para unidades, decenas, centenas, etc.
    unidades = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
    decenas = ["", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"]
    teens = ["diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"]
    centenas = ["", "cien", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"]
    
    if numero == 0:
        return "cero"
    
    partes = []
    
    millones = numero // 1000000
    if millones > 0:
        if millones == 1:
            partes.append("un millón")
        else:
            partes.append(convertir_numero_a_texto(millones) + " millones")

    miles = (numero % 1000000) // 1000
    if miles > 0:
        if miles == 1:
            partes.append("mil")
        else:
            partes.append(convertir_numero_a_texto(miles) + " mil")

    cientos = (numero % 1000) // 100
    if cientos > 0:
        if cientos == 1 and numero != 100:
            partes.append("ciento")
        else:
            partes.append(centenas[cientos])
    
    decenas_y_unidades = numero % 100
    if decenas_y_unidades >= 10 and decenas_y_unidades < 20:
        partes.append(teens[decenas_y_unidades - 10])
    else:
        d = decenas_y_unidades // 10
        u = decenas_y_unidades % 10
        if d > 0:
            if d == 2 and u > 0:
                partes.append("veinti"+unidades[u])
            else:
                partes.append(decenas[d])
        if u > 0 and d != 2:
            if d > 0:  # Si hay decenas, añadir "y"
                partes.append("y " + unidades[u])
            else:
                partes.append(unidades[u])
    texto_numero = " ".join(partes).strip()
    texto_numero = texto_numero.capitalize()

    return texto_numero


def search_files(path_in, path_out, list_path, extension):
    df = pd.read_excel(list_path)
    for root, dirs, files in os.walk(path_in):
        for file in files:
            if file.upper().endswith(extension.upper()):
                ruta_path = os.path.join(root, file)
                numbers_match = re.match(r'\d+', file)
                if numbers_match:
                    numbers = int(numbers_match.group())
                    if numbers in df['ID'].values:
                        print("this is true ", numbers)
                        try:
                            shutil.copy(ruta_path, path_out)
                            print(f"'{file}' copiado a la carpeta de salida.")
                        except shutil.Error as e:
                            print(f"Error al copiar '{file}': {e}")


def get_here(name = "", folder = ""):
    # Devuelve la ubicación actual, y se puede agregar un folder y un archivo 
    here_base = os.getcwd()
    if folder != "":
        here_part = os.path.join(here_base, folder)
    else:
        here_part = here_base
    if name != "":
        here = os.path.join(here_part, name)
    else:
        here = here_part
    return here


def mejor_coincidencia(variable, lista):
    mejor_coincidencia = process.extractOne(variable, lista)
    return mejor_coincidencia[0]


def encontrar_coincidencias(dataframe, columna, valor_x, top_n=3):
    # Usa FuzzyWuzzy para entocntrar coincidencias entre un dataframe y un valor
    opciones = dataframe[columna].tolist()
    mejores_coincidencias = process.extractBests(valor_x, opciones, scorer=fuzz.token_sort_ratio, limit=top_n)
    valores_coincidentes = [tupla[0] for tupla in mejores_coincidencias]
    new_dataframe = dataframe[dataframe[columna].isin(valores_coincidentes)]
    return new_dataframe


def convertir_conjunto(CONJUNTO):
    ''' DADO UN CONJUNTO DEVUELVE EL: FRACCIONAMIENTO, FRENTE, ETAPA'''
    E00 = ["E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "E14", "E15", "E16", "E17", "E18", "E19", "E20", "E21", "E22", "E23", "E24", "E25", "E26", "E27", "E28"]
    UJN = ["VSU", "   ", "RLU", "BSU", "S2U", "BDU", "P2U", "   ", "   ", "USV", "ELU", "UED", "UB4", "   ", "URC", "UNL", "JUU", "USI", "   ", "UR7", "UMA", "UFB", "UPR", "UMO", "CJM", "   ", "   ", "   "]
    CJQ = ["CST", "   ", "CRL", "CB2", "CS2", "CB3", "CHM", "   ", "   ", "CS3", "CEL", "CVP", "CB4", "   ", "CRÑ", "CMN", "CBJ", "CSI", "   ", "CR7", "CMA", "CFB", "CPR", "CMS", "CJM", "   ", "PME", "VDV"]

    PRIMER = CONJUNTO[0:3]
    ETAPA = CONJUNTO[7:10]
    FRENTE = CONJUNTO[4:6]
    NUMERO = int(PRIMER[1:3]) - 1

    if ETAPA == "I01" or ETAPA == "U02":
        if int(FRENTE) > 79:
            FRAC = CJQ[NUMERO]
        else:
            FRAC = UJN[NUMERO]
    else: 
        FRAC = CJQ[NUMERO]
    return (FRAC, FRENTE, ETAPA)


def take_first_row():
    # Guarda el dataframe en excel, extrayendo la primera fila y devolviendola
    lock = threading.Lock()
    lock.acquire()
    flash_folder = get_here(folder="cache")
    flash_memory = get_here(folder="cache", name="flash_memory_mth.xlsx")
    df = pd.read_excel(flash_memory)
    try:
        row = df.iloc[0]
        df.drop(df.index[0], inplace=True)
    except IndexError:
        row = []
    try:
        df.to_excel(flash_memory, index=False) 
    except OSError as e:
        os.makedirs(flash_folder)
        df.to_excel(flash_memory, index=False)
    lock.release()
    return row, df

def save_flash_memory(df):
    flash_memory = get_here(folder="cache", name="flash_memory_mth.xlsx")
    df.to_excel(flash_memory, index=False)

def multithreader(app, df, threads=4, headless=True):
    # Convierte un programa en multithread
    lenght_df = len(df)
    if lenght_df < threads:
        threads = lenght_df

    flash_memory = get_here(folder=cache, name=flash_memory_mth.xlxs)
    df.to_excel(flash_memory, index=False)
    thread_list = []

    if threads != 0:
        for thread in range(threads):
            driver = create_driver(headless=headless)
            thread = threading.Thread(target=app, args=(driver, df,))
            thread.start()
            thread_list.append(thread)
        for thread in thread_list:
            thread.join()


def divide_and_place_windows(screen_x, screen_y, screen_width, screen_height):
    # Obtenemos las ventanas de Firefox

    # Lista de palabras a excluir
    excluir_palabras = ['youtube', 'buscar', 'mercado', 'amazon']

    # Filtrar ventanas que no contengan ninguna de las palabras a excluir
    windows = [window for window in gw.getWindowsWithTitle('Mozilla Firefox') 
               if not any(palabra in window.title.lower() for palabra in excluir_palabras)]
    
    num_windows = len(windows)
    if num_windows == 0:
        return
    
    # Calcular columnas y filas
    columns = math.ceil(math.sqrt(num_windows))
    rows = num_windows // columns
    
    normales = columns * rows
    residuo = 0
    extra = False

    # Verificar si hay ventanas extra que no caben en una cuadrícula perfecta
    if num_windows % columns != 0:
        residuo = num_windows - normales
        # Aumentar columnas o filas dependiendo del tamaño disponible
        if screen_width // columns > screen_height // rows:
            extra = "Column"
            columns += 1
        else:
            rows += 1
            extra = "Row"
    
    # Calcular el tamaño estándar de las ventanas
    normal_window_width = screen_width // columns
    normal_window_height = screen_height // rows

    extra_window_width = normal_window_width
    extra_window_height = normal_window_height

    # Ajustar las ventanas extra si existen
    if extra == "Column":
        extra_window_height = screen_height // residuo
    elif extra == "Row":
        extra_window_width = screen_width // residuo

    # Posicionar y redimensionar cada ventana
    for index, window in enumerate(windows):
        row = index // columns
        col = index % columns

        # Calcular posición x e y
        x = screen_x + col * normal_window_width
        y = screen_y + row * normal_window_height

        try:
            if extra == "Column" and col == columns - 1:
                window.resizeTo(normal_window_width, extra_window_height)
                y = screen_y + row * extra_window_height
            elif extra == "Row" and row == rows - 1:
                window.resizeTo(extra_window_width, normal_window_height)
                x = screen_x + col * extra_window_width
            else:
                window.resizeTo(normal_window_width, normal_window_height)

            window.moveTo(x, y)
        except Exception as e:
            print(f"Error resizing or moving window '{window.title}': {e}")


def organize_firefox():
    # Ancho de la barra de tareas en el lado izquierdo (60 píxeles)
    taskbar_width = 55
    
    # Obtener las dimensiones de la pantalla principal
    screen = screeninfo.get_monitors()[1]
    screen_width, screen_height = screen.width, screen.height
    
    # Ajustar el área disponible para las ventanas (descontar la barra de tareas)
    screen_x = taskbar_width
    available_width = screen_width - taskbar_width
    screen_y = 0  # La barra de tareas no afecta la altura
    
    # Ajustar las ventanas según la cantidad, comenzando desde el área completa de la pantalla
    divide_and_place_windows(screen_x, screen_y, available_width, screen_height)




#---------------------------U t i l e r i a   I n t e r f a c e------------------------------------

def window_input(title, label, options):
    result = None

    def seleccionar_letra(letra):
        nonlocal result
        result = letra
        ventana.destroy()

    ventana = tk.Tk()
    ventana.title(title)

    etiqueta = tk.Label(ventana, text=label)
    etiqueta.pack()

    for letra, descripcion in options.items():
        frame = tk.Frame(ventana)
        boton = tk.Button(frame, text=letra, command=lambda l=letra: seleccionar_letra(l))
        descripcion_label = tk.Label(frame, text=descripcion)
        descripcion_label.pack(side=tk.LEFT)
        boton.pack(side=tk.RIGHT)
        frame.pack()

    ventana.mainloop()
    return result

#----------------------------U t i l e r í a   S e l e n i u m-------------------------------------

def create_driver(driver_type='firefox', headless=True, download_folder=None):
    """
    Esta función crea un driver, requisitos:
        driver_type= chrome, firefox o edge
        headless= True o False 
        download_folder = None o Ruta
    """

    driver = None  # Inicializamos el driver como None

    # Opción para usar Chrome
    if driver_type == 'chrome':
        options = ChromeOptions()
        if headless:
            options.add_argument('--headless')
        prefs = {
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': False,
            'plugins.always_open_pdf_externally': True
        }
        if download_folder:
            prefs['download.default_directory'] = download_folder
        options.add_experimental_option('prefs', prefs)
        driver = webdriver.Chrome(options=options)

    # Opción para usar Firefox
    elif driver_type == 'firefox':
        options = FirefoxOptions()
        if headless:
            options.add_argument('-headless')
        options.set_preference("browser.download.folderList", 2)
        if download_folder:
            options.set_preference("browser.download.dir", download_folder)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", True)
        options.set_preference("browser.download.useDownloadDir", True)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.helperApps.alwaysAsk.force", False)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
        options.set_preference("browser.download.manager.showAlertOnComplete", False)
        options.set_preference("browser.download.manager.useWindow", False)
        options.set_preference("dom.block_download_insecure", False)
        options.set_preference("pdfjs.disabled", True)
        options.set_preference("plugin.scan.plid.all", False)
        options.set_preference("dom.popup_maximum", 100)
        options.set_preference("app.update.enabled", False)
        options.set_preference("-disable-updates", True)
        driver = webdriver.Firefox(options=options)

    # Opción para usar Edge
    elif driver_type == 'edge':
        options = EdgeOptions()
        if headless:
            options.add_argument('headless')
        prefs = {
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': False,
            'plugins.always_open_pdf_externally': True
        }
        if download_folder:
            prefs['download.default_directory'] = download_folder
        options.set_capability("prefs", prefs)
        driver = webdriver.Edge(options=options)

    # Aquí va la condición para no headless
    if not headless:
        organize_firefox()  # Función que organizas si no es headless

    return driver  # Retornar el driver al final


def reset_driver(driver, headless = True):
    driver.quit()
    while True:
        try:
            driver = create_driver(headless = headless)
            break
        except Exception as e:
            time.sleep(3)
            print("No se pudo crear Driver apaga el antivirus")
    return(driver)


def easy_select(driver, name_element, to_select, tipo='id'):
    '''
    args: driver, name, select, tipo='id', 'xpath', 'css', class
    '''
    if tipo.lower() == "id":
        select_by = By.ID
    elif tipo.lower() == "xpath":
        select_by = By.XPATH
    elif tipo.lower() == "css":
        select_by = By.CSS_SELECTOR
    elif tipo.lower() == "class":
        select_by = CLASS_NAME
    select_element = driver.find_element(select_by, name_element)
    select = Select(select_element)
    select.select_by_visible_text(to_select)


def swdw(driver, tiempo, LOOKBY, LOOKFOR):
    '''
    Da click al objeto seleccionado: driver, tiempo, que buscas
    
    Args: 
        driver: el driver a usar
        TIME: el tiempo a esperar
        LOOKBY: 0: By.XPATH, 1: By.ID, 2: By.CSS_SELECTOR, 3: By.NAME, 4: By.PARTIAL_LINK_TEXT, 5: By.CLASS
        LOOKFOR: EL CÓDIGO A BUSCAR
    '''
    # Forma de busqueda
    buscar_por = {0: By.XPATH, 1: By.ID, 2: By.CSS_SELECTOR, 3: By.NAME, 4: By.PARTIAL_LINK_TEXT, 5: By.CLASS_NAME}
    max_intentos = 10
    intentos = 0
    tiempo_split = tiempo / (max_intentos + 1)
    while intentos < max_intentos:
        try:
            simple = WebDriverWait(driver, tiempo_split).until(EC.element_to_be_clickable((buscar_por[LOOKBY], LOOKFOR)))
            break
        except (StaleElementReferenceException, ElementClickInterceptedException) as e:
            time.sleep(0.5)
            continue
        except TimeoutException:
            intentos += 1
            continue

    simple = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((buscar_por[LOOKBY], LOOKFOR)))
    return simple


def iswdw(driver, tiempo, LOOKBY, LOOKFOR, WAITBY, WAITFOR):
    while True:
        try:
            swdw(driver, 1, WAITBY, WAITFOR)
            break

        except TimeoutException:
            swdw(driver, 2, LOOKBY, LOOKFOR).click()

        except StaleElementReferenceException:
            time.sleep(2)



def stf(driver, LOOKBY, LOOKFOR, SEND):
    '''DENTRO DEL MODULO DE CONSTRUCCIÓN INTRODUCE UN DATO DENTRO DEL FRAME'''
    while True:
        # Hacemos click en el elemento que querémos abrir en frame
        try:
            driver.switch_to.window(driver.window_handles[0])
            swdw(driver, 3, LOOKBY, LOOKFOR).click()
            # Revisamos que se abra otra ventana
            ventanas = driver.window_handles
            cantidad_ventanas = len(ventanas)
            time.sleep(0.8)
            if cantidad_ventanas == 2:
                break
            else:
                time.sleep(0.5)
                continue
        except (TimeoutException, StaleElementReferenceException):
            continue
        except ElementClickInterceptedException:
            time.sleep(1)
            ventanas = driver.window_handles
            cantidad_ventanas = len(ventanas)
            if cantidad_ventanas == 2:
                break
            else:
                time.sleep(0.5)
                continue
    while True:
        for counter in range(10):
            try:
                driver.switch_to.window(driver.window_handles[1])
                driver.switch_to.frame(0)
                break
            except (NoSuchFrameException, IndexError, NoSuchWindowException):
                time.sleep(0.8)
                continue
        try:
            swdw(driver, 5, 0, "//input[@title='Término de Búsqueda']").clear()
            swdw(driver, 2, 0, "//input[@title='Término de Búsqueda']").send_keys("%" + SEND + "%" + Keys.TAB)
            swdw(driver, 2, 0, "//button[text()='Ir']").click()
            swdw(driver, 1, 0, "//table[@class='x1o']/tbody[1]/tr[2]/td[2]/a[1]/img[1]").click()
            time.sleep(0.8)
        except TimeoutException:
            continue
        except (NoSuchWindowException, IndexError):
            break
        ventanas = driver.window_handles
        cantidad_ventanas = len(ventanas)
        if cantidad_ventanas == 1:
            driver.switch_to.window(driver.window_handles[0])
            driver.switch_to.default_content()
            swdw(driver, 4, 0, '//body').click()
            break


def beautiful_table(driver, element="class", name="x1o"):

    # Almacenadores
    headers = []
    rows = []

    # Obtener el contenido HTML de la página web
    html = driver.page_source
    # Crear un objeto BeautifulSoup a partir del HTML
    soup = BeautifulSoup(html, 'html.parser')    
    # Encontrar la tabla en el HTML
    table = soup.find('table', {element: name})
   
    # Extraer los encabezados de la tabla
    header_row = table.find('tr')
    if header_row:
        list_head = header_row.find_all('th')
        for th in list_head:
            # revisamos textos, ids e indices
            text = th.text.strip()
            idd = th.get('id')
            indexx = list_head.index(th) 
            if text:
                headers.append(text)
            elif idd:
                headers.append(idd)
            else:
                headers.append(indexx)
    
    # Extraer las filas de la tabla
    for row in table.find_all('tr')[1:]:
        cells = []
        for td in row.find_all('td'):
            # revisamos a, sapn e inputs
            span = td.find('span') 
            inputt = td.find('input')
            a = td.find_all('a')

            # Priorizamos la información dependiendo de lo que tengamos en dicha tabla
            if inputt:
                text = ""
                idd = inputt.get('id')
            elif span:
                text = span.text.strip()
                idd = span.get('id')
            elif a:
                # si tenemos mas de un a los concatenamos en un a
                new_a = ""
                for sub_a in a:
                    new_a += sub_a.text
                text = new_a
                idd = a[-1].get('id')
            else:
                text = td.text
                idd = td.get('id')

            # Se asigna el valor
            if text != "":
                cells.append(text)
            elif idd != "":
                cells.append(idd)
            else:
                cells.append(td)
        rows.append(cells)

    # Pasamos a DataFrame
    df = pd.DataFrame(rows, columns=headers)    
    return df


def check_where_we_are(driver, list_xpath):
    get_what = ""
    time.sleep(0.3)
    for x in range(50):
        time.sleep(0.02)
        for xpath in list_xpath:
            try:
                driver.find_element(By.XPATH, xpath)
                get_what = xpath
                break
            except NoSuchElementException:
                continue
        if get_what != "":
            indice = list_xpath.index(get_what)
            break
    if get_what == "":
        indice = "x"
    return indice


#----------------------------U t i l e r i a   M o d u l o-------------------------------


def acceder_oracle(driver):
    '''DADO UN USUARIO Y CONTRASEÑA INGRESA EN ORACLE'''
    user = ORACLE_ID.USER
    password = ORACLE_ID.PASSWORD
    HOME = "//table[@class='x6w']/tbody[1]/tr[1]/td[3]/a[1]"
    driver.get("http://siapp3.javer.com.mx:8010/OA_HTML/AppsLogin")
    time.sleep(1)
    if swdw(driver, 2, 0, "//table[@id='langOptionsTable']/tbody[1]/tr[2]/td[2]/span[1]").text == "Select a Language:":
        swdw(driver, 1, 0, "//img[@title='Latin American Spanish']").click()
    swdw(driver, 10, 3, "usernameField").clear()
    swdw(driver, 1, 3, "usernameField").send_keys(user + Keys.TAB)
    swdw(driver, 1, 3, "passwordField").send_keys(password + Keys.ENTER)
    try:
        swdw(driver, 5, 0, '//*[@id="PageLayoutRN"]/table[1]/tbody/tr[2]/td/table/tbody/tr[1]/td[1]/img')
    except Exception as e:
        print(e)



def from_main_menu(driver, ruta):
    # Desde el main menu damos click a los elementos de la lista ruta 
    # Revizamos estar en mainmenu buscando un elemento en caso de no estarlo damos click a ir al main menu
    try:
        swdw(driver, 1, 0, "//table[@id='respList']/tbody[1]/tr[1]/td[1]/ul[1]/li[1]/a[1]")
    except TimeoutException:
        swdw(driver, 3, 0, "//table[@class='x6w']/tbody[1]/tr[1]/td[3]/a[1]").click()

    # Tratamos de dar click al último elemento, si no, vamos uno por uno
    try:
        xpath = "//a[contains(text(), '" + ruta[-1] + "')]"
        swdw(driver, 3, 0, xpath).click()
    except TimeoutException:
        for element in ruta:
            xpath = "//a[contains(text(), '" + element + "')]"
            swdw(driver, 3, 0, xpath).click()


def new_mainmenu(driver, ruta):
    while True:
        try:
            df = pd.DataFrame(columns=['Text', 'Alt'])
            # Obtener el contenido HTML de la página web
            html = driver.page_source
            # Crear un objeto BeautifulSoup a partir del HTML
            soup = BeautifulSoup(html, 'html.parser')    
            # Encontrar la tabla en el HTML
            table = soup.find('ul', {'id': 'treemenu1'})

            for li in table.findAll('li'):
                a = li.find('a', recursive=False)
                # Encuentra el texto dentro de a
                text = a.get_text()
                # Encuentra el primer elemento img dentro de li
                img = a.find('img', recursive=False)
                try:
                    # Extrae el atributo alt
                    alt = img['alt']
                except KeyError:
                    alt = "Click"
                # Imprime el atributo alt
                df1 = pd.DataFrame({
                    'Text': [text],
                    'Alt': [alt]
                    })
                df = pd.concat([df, df1], ignore_index=True)
            df = df.loc[df['Text'].isin(ruta)]
            df = df[df['Alt'] == 'Collapse']

            goto = [valor for valor in ruta if valor not in df['Text'].values]

            for element in goto:
                xpath = "//a[contains(text(), '" + element + "')]"
                # print(f'click a {element}')
                swdw(driver, 1, 0, xpath).click()
            break

        except AttributeError:
            print("No menú")
            try:
                x_acceso = "//img[@title='Español Latinoamericano']"
                swdw(driver, 0.2, 0, x_acceso)
                acceder_oracle(driver)
            except TimeoutException:
                print("No acceso oracle")
            try:
                xpath = "//a[contains(text(), 'Página Inicial')]"
                swdw(driver, 0.2, 0, xpath).click()
                swdw(driver, 1, 0, "//h2[contains(text(), 'Menú Principal')]")
            except TimeoutException:
                print("No página inicial")

        except Exception as E:
            print("Otro error", E)

# Utencilios go in modulo---------------------------------------------------------------

def go_in_mdc(driver, ruta, check):
    # Se dirige de forma correcta a un lugar en modulo de construcción
    while True:
        try:
            driver.title
            break
        except Exception as e:
            print(e)

    if driver.title == "":
        acceder_oracle(driver)
    
    try:
        swdw(driver, 1, 0, check)

    except TimeoutException:
        new_mainmenu(driver, ruta)


def go_front(driver):
    # Abre los frentes en modulo de construcción
    ruta = ["JAV_MC_CAO_QRO", "Frentes", "Buscar frentes"]
    check = "//h2[text()='Frentes de Construcción']"
    go_in_mdc(driver, ruta, check)


def go_sistema_reportes(driver):
    # Abre los frentes en sistema de reportes
    ruta = ["JAV_MC_CAO_QRO", "Reportes Generales", "Sistema de reportes"]
    check = "(//span[text()='Hacer clic en el icono para ejecutar el reporte'])[1]"
    go_in_mdc(driver, ruta, check)


def go_contract(driver):
    ruta = ["JAV_MC_CAO_QRO", "Contratos", "Buscar contratos"]
    check = "//h2[text()='Listado de Contratos']"
    go_in_mdc(driver, ruta, check)


def go_monitor(driver):
    # Abre los el monitor sistema de reportes
    ruta = ["JAV_MC_CAO_QRO", "Request", "Monitor"]
    check = "//td[text()='Tabla Resumen de Solicitudes']"
    go_in_mdc(driver, ruta, check)


def fill_contract(driver, conjunto):
    org, frente, conjunto_1 = convertir_conjunto(conjunto)
    xpath_org = '//*[@id="OrganizationLOV__xc_0"]/a/img'
    xpath_frente = '//*[@id="FrontSearchLOV__xc_0"]/a/img'
    xpath_conjunto = '//*[@id="FrontBuildSetLOV__xc_0"]/a/img'

    value="CJM-CONSTRUCCION MARQUES DEL RIO"
    org_escrita = swdw(driver, 1, 1, 'OrganizationLOV').get_attribute('value')
    if org_escrita is not None:
        if org not in str(org_escrita):
            stf(driver, 0, xpath_org, org)
    else:
        stf(driver, 0, xpath_org, org)

    frente_escrita = swdw(driver, 1, 1, 'FrontSearchLOV').get_attribute('value')
    if frente_escrita is not None:
        if frente not in str(frente_escrita):
            stf(driver, 0, xpath_frente, frente)
    else:
        stf(driver, 0, xpath_frente, frente)\

    conjunto_escrita = swdw(driver, 1, 1, 'FrontBuildSetLOV').get_attribute('value')
    if conjunto_escrita is not None:
        if conjunto not in str(conjunto_escrita):
            stf(driver, 0, xpath_conjunto, conjunto)
    else:
        stf(driver, 0, xpath_conjunto, conjunto)

    swdw(driver, 0, 1, 'Search').click()
        

def get_contract_table(driver, conjunto):
    while True:
        try:
            # Selecciona el contrato 
            go_contract(driver)
            fill_contract(driver, conjunto)
            easy_select(driver, 'ResultsTable:ResultsDisplayed:0', '100')
            time.sleep(1)
            df_contracts = beautiful_table(driver)
            return df_contracts
        except Exception as e:
            print(e)


def get_contract(driver, conjunto, contract):
    while True:
        try:
            # Selecciona el contrato 
            df_contracts = get_contract_table(driver, conjunto)
            contrato_details = df_contracts[df_contracts['Código'] == contract]['Detalles'].iloc[0]
            swdw(driver, 1, 1, contrato_details).click()
            check_contract = swdw(driver, 2, 1, 'Code').text
            if check_contract == contract:
                print('success ', contract)
            break
        except Exception as e:
            print(e)


def get_contract_contracts(driver, conjunto, contract):
    get_contract(driver, conjunto, contract)
    while True: 
        try:
            swdw(driver, 1, 1, 'LegalContractsLink').click()
            swdw(driver, 2, 1, 'AddDoc')
            break
        except TimeoutException as e:
            print(e)


def charge_document_contract(driver, conjunto, contract, doc, doc_name, doc_desc):
    file_size = round(os.path.getsize(doc) / (1024 * 1024), 3)
    if file_size > 10:
        print(f'Tamaño de {doc} superior al permitido: {file_size}')
        return

    get_contract(driver, conjunto, contract)
    span_cancelado = swdw(driver, 1 , 1, 'Status').text
    if span_cancelado != 'Cancelado':
        iswdw(driver, 2, 1, 'DocumentsLink', 1, 'AddDocument')
        iswdw(driver, 2, 1, 'AddDocument', 1, 'File_oafileUpload')
        swdw(driver, 2, 1, 'File_oafileUpload').send_keys(doc)
        swdw(driver, 2, 1, 'ShortName').send_keys(doc_name)
        swdw(driver, 2, 1, 'Description').send_keys(doc_desc)
        swdw(driver, 2, 1, 'Apply').click()
    time.sleep(2)
    for x in range(18):
        try: 
            swdw(driver, 3, 1, 'DocumentsLink')
            break
        except TimeoutException:
            time.sleep(2)
    swdw(driver, 2, 1, 'XXMCAN_CONTRACT_SEARCH').click()


# Utencilios set --------------------------------------------------------------

     
def go_org(driver, ORG):

    # Definiciones
    xpath_org = "(//span[@id='OrganizationLOV__xc_0']//img)[2]"

    while True:
        # Intentamos dos veces
        try:
            # Enviamos la organización al buscador de frentes y damos click
            stf(driver, 0, xpath_org, ORG)
            swdw(driver, 8, 1, "OrganizationLOV").click()
            swdw(driver, 2, 0, "//button[@id='Search']").click()

            # Ampliamos busqueda a 100
            easy_select(driver, "FrontsTable:ResultsDisplayed:0", "100", tipo='id')
            swdw(driver, 2, 1, 'FrontsTable:ShowDetails:0')
            time.sleep(0.5)
            break

        except (StaleElementReferenceException, ElementClickInterceptedException):
            continue


def busca_reporte_sistema_de_reportes(driver, archivo, what):
    ''' What: colSwitcherExecute, colPdf, colXls
    Busca y da click en un reporte'''
    def get_click_a(driver, what):
        table = beautiful_table(driver)
        row = table.loc[table['Descripción'] == archivo].iloc[0]
        click_a = row.loc[what]
        regx = re.search(r'TableReportsRN:\w{3}(\w):\w', click_a)
        if regx is not None:
            regx = regx.group(1)
        return regx, click_a

    while True:
        # Da click al botón que coincide con el reporte buscado
        try:
            regx, click_a = get_click_a(driver, what)
            if what == 'colSwitcherExecute':
                swdw(driver, 1, 1, click_a).click()
                break
            else:
                while regx != '1':
                    swdw(driver, 2, 1, 'buttonRefreshReport').click()
                    regx, click_a = get_click_a(driver, what)
                    time.sleep(1)
                bot = click_a[:-2] + "img" + click_a[-2:]
                print(bot)
                swdw(driver, 2, 1, bot).click()
                time.sleep(3)
                print(bot)
                break

        except TimeoutException:
            print("Este error es el que nos saca")
            continue
    time.sleep(1)



#----------------------------U t i l e r i a   H o j a   V i a j e r a-------------------

def acceso_hv(driver):

    # Usuario y contraseña
    user = JAVER_ID.USER
    password = JAVER_ID.PASSWORD
        
    # Acceso a página web
    driver.get("https://hojaviajeradigital.javer.com.mx:9260/#/inicio")

    # revisamos que éxsita sesión iniciada, en caso contrario iniciamos sesión
    try:
        icono = "//h3[text()=' fprado']"
        swdw(driver, 2, 0, icono)
    except TimeoutException:
        swdw(driver, 8, 1, "mat-input-0").send_keys(user)
        swdw(driver, 2, 1, "mat-input-1").send_keys(password)
        swdw(driver, 2, 0, "//span[text()='Iniciar sesión']").click()
    inicio = "//h2[text()='INICIO']"
    time.sleep(1)


#---------------------------U t i l e r i a   d o c 2 s i g n----------------------------

def acceso_doc2sign(driver):

    # Usuario y contraseña
    user = "fprado@javer.com.mx"
    password = "dprmt141592FRA+-n"
    continuar = '//*[@id="login"]/div/div[1]/div[4]/button[2]'

    # Acceso a página web
    driver.get("https://www.doc2sign.com/Acceso/Login")

    # Revisamos que éxista sesión iniciada
    try:
        swdw(driver, 1, 5, "icon-home")

    except TimeoutException:
        swdw(driver, 2, 1, "email").send_keys(user)
        time.sleep(0.4)
        swdw(driver, 2, 0, continuar).click()
        time.sleep(1)
        swdw(driver, 2, 1, "password").send_keys(password)
        time.sleep(1)
        swdw(driver, 2, 1, "enviarcontinuar").click()



def go_doc2sign(driver, consulta = True):
    acceso_doc2sign(driver)
    if consulta:
        pestana1 = 'm4'
        pestana2 = 'op16'

    else:
        pestana1 = 'm85'
        pestana2 = 'op9'

    swdw(driver, 2, 1, pestana1).click()
    swdw(driver, 2, 1, pestana2).click()

    ruta_descargas = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\'
    aviso_pdf = os.path.join(ruta_descargas, 'Aviso_de_Privacidad.pdf')
    terminos_pdf = os.path.join(ruta_descargas, 'Terminos_y_condiciones_de_uso.pdf')

    aviso_pdf_exist = os.path.exists(aviso_pdf)
    terminos_pdf_exist = os.path.exists(terminos_pdf)

    if aviso_pdf_exist:
        os.remove(aviso_pdf)

    if terminos_pdf_exist:
        os.remove(terminos_pdf)


# go_doc2sign(create_driver(headless=False))
# driver.quit()