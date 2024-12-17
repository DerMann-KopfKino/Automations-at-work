import MODDECON as MDC
from Multiherramienta import *
import BDD_A as BDD
import time
import os

#---GENERALES---

HERE = os.getcwd()
root_bdd = HERE + "\\BDD\\"


LISTA_REPORTES = [["Reporte Estado de Cuenta de Contratos"], ["Reporte de Finiquito de Obra"], ["JVR-Reportes Transacciones de Inventario por Frente"]]



T_ORGS = ["CRÑ", "CMN", "CR7", "CMA", "CFB", "CPR", "CMS", "CJM", "URC", "UNL", "UR7", "UMA", "UFB", "UPR", "UMO", "PME", "VDV"]
BDD.LIST_TO_XLSX(T_ORGS, ['Organizaciones'], os.getcwd() + '\\BDD\\BDD_ORGANIZACIONES.xlsx')
# BDD.DF_TO_PKL(BDD.LIST_TO_DF(T_ORGS, ["Organizaciones"]), os.getcwd() + "\\BDD\\BDD_ORGANIZACIONES.pkl")
M_ORGS = ["CRÑ", "URC", "CR7", "UR7", "CMA", "UMA", "CJM"]


TABLAS_RUTAS = [["JAV_MC_CAO_QRO", "Reportes Generales", "Reporte Estado de Cuenta de Contratos"],
                ["JAV_MC_CAO_QRO", "Reportes Generales", "Reporte de Finiquito de Obra"],
                ["JAV_MC_CAO_QRO", "Reportes Generales", "JVR-Reportes Transacciones de Inventario por Frente"],
                ["JAV_MC_CAO_QRO", "Request", "Monitor"]]


def download_file_reports(threads=6):
    for x in range(3):
        try:
            MDC.descargador_archivos_reportes(threads=threads, headless=headless)
        except:
            time.sleep(1)
        try: 
            MDC.rename_reports(dias=dias)
        except:
            time.sleep(1)


headless = False
threads  = 6
dias = 1

# MDC.update_contract(headless=headless)

# MDC.frentes_existentes(HERE + '//BDD//BDD_ORGANIZACIONES.xlsx', headless=headless , threads=threads)
# MDC.estatus_conjuntos(HERE + '//BDD//BDD_Frentes_existentes.xlsx', headless=headless, threads=threads)
# MDC.clean_archivos_reportes(días=-1)
# MDC.generador_de_reportes(headless=headless, threads=threads, dias=dias)
# download_file_reports(threads=threads)
# MDC.rename_reports(dias=dias)
# BDD.CONCATENADOR_BDD()
# BDD.copy_paste_reports()

# apagar_pc()

# MDC.download_from_contract(driver, df_root, objetos)

# MDC.rename_contracts_purchaser(("CMS", "CRÑ", "CMA", "CPR", "VDV"), "Ortega Villa, Gema Delhi", headless=False)
# MDC.multithread_rename_contracts_purchaser(("CMS", "VDV", "CRÑ","CPR", "CMA", "PME"), "Ortega Villa, Gema Delhi", headless=headless, threads=threads)

desktop_route = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\"
route_bd = os.path.join(desktop_route, "Insumos TIPO 4.xlsx")
route_task = os.path.join(desktop_route, "task.xlsx")
route_out = os.path.join(desktop_route, "new_df.xlsx")


# BDD.buscador_de_insumos(route_bd, route_task, route_out)