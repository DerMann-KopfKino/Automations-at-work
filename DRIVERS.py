from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService
import MODDECON as MDC
import BDD_A as BDD
import os

def create_driver(driver_type='chrome', headless=False, download_folder=None):
    """
    Esta función crea un driver de selenium con las siguientes opciones:
        driver_type= chrome, firefox o edge
        headless= True o False
        download_folder= None o ruta de carpeta
    Retorna el driver creado.
    """
    if driver_type == 'chrome':
        service = ChromeService(executable_path="chromedriver.exe")
        options = webdriver.ChromeOptions()
    elif driver_type == 'firefox':
        service = FirefoxService(executable_path="geckodriver.exe")
        options = webdriver.FirefoxOptions()
    elif driver_type == 'edge':
        service = EdgeService(executable_path="msedgedriver.exe")
        options = webdriver.EdgeOptions()
    else:
        raise ValueError("El tipo de driver especificado no es válido.")

    if headless:
        options.add_argument('--headless')

    if download_folder:
        if driver_type == 'chrome':
            prefs = {'download.default_directory': download_folder,
                     'download.prompt_for_download': False,
                     'download.directory_upgrade': True,
                     'safebrowsing.enabled': False,
                     'plugins.always_open_pdf_externally': True}
            options.add_experimental_option('prefs', prefs)
        elif driver_type == 'firefox':
            profile = webdriver.FirefoxProfile()
            profile.set_preference("browser.download.folderList", 2)
            profile.set_preference("browser.download.dir", download_folder)
            profile.set_preference("browser.download.useDownloadDir", True)
            profile.set_preference("browser.download.manager.showWhenStarting", False)
            profile.set_preference("browser.helperApps.alwaysAsk.force", False)
            profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
            profile.set_preference("browser.download.manager.showAlertOnComplete", False)
            profile.set_preference("browser.download.manager.useWindow", False)
            profile.set_preference("pdfjs.disabled", True)
            profile.set_preference("plugin.scan.plid.all", False)
            profile.set_preference("dom.popup_maximum", 100)
            profile.set_preference("app.update.enabled", False)
            options.profile = profile
        elif driver_type == 'edge':
            prefs = {'download.default_directory': download_folder,
                     'download.prompt_for_download': False,
                     'download.directory_upgrade': True,
                     'safebrowsing.enabled': False,
                     'plugins.always_open_pdf_externally': True}
            options.add_experimental_option('prefs', prefs)

    if driver_type == 'chrome':
        driver = webdriver.Chrome(service=service, options=options)
    elif driver_type == 'firefox':
        driver = webdriver.Firefox(service=service, options=options)
    elif driver_type == 'edge':
        driver = webdriver.Edge(service=service, options=options)

    return driver

HERE = os.getcwd()
ORGS_ACTIVAS = BDD.PKL_TO_DF(HERE + "\\BDD\\BDD_ORGANIZACIONES.pkl")
driver = create_driver()

MDC.BDD_FRENTES_EXISTENTES(driver, ORGS_ACTIVAS)