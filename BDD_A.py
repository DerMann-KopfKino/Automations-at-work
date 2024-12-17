import os
import re
import time
import glob
import datetime
import requests
import json
import glob
import shutil
import numpy as np
import pandas as pd
import pdfplumber
from PyPDF2 import PdfWriter
from icecream import ic
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from shareplum  import Office365, Site
from shareplum.site import Version
from Multiherramienta import *

#----GLOBALES----

HERE = os.getcwd()

def conjunto(CONJUNTO):
    ''' DADO UN CONJUNTO DEVUELVE EL: FRACCIONAMIENTO, FRENTE, ETAPA'''
    E00 = ["E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "E14", "E15", "E16", "E17", "E18", "E19", "E20", "E21", "E22", "E23", "E24", "E25", "E26", "E27"]
    UJN = ["VSU", "   ", "RLU", "BSU", "S2U", "BDU", "P2U", "   ", "   ", "USV", "ELU", "UED", "UB4", "   ", "URC", "UNL", "JUU", "USI", "   ", "UR7", "UMA", "UFB", "UPR", "UMO", "CJM", "   ", "   "]
    CJQ = ["CST", "   ", "CRL", "CB2", "CS2", "CB3", "CHM", "   ", "   ", "CS3", "CEL", "CVP", "CB4", "   ", "CRÑ", "CMN", "CBJ", "CSI", "   ", "CR7", "CMA", "CFB", "CPR", "CMS", "CJM", "   ", "PME"]
    #SCJ = 
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


def GET_DOWNLOAD_PATH():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


#----CONVERSIONES DE DATAFRAMES, PICKLES Y LISTAS----

def PKL_TO_DF(UBICACIÓN):
	'''DADA LA UBICACIÓN DE UN ARCHIVO PIKLE LO CONVIERTE EN DATAFRAME Y LO DEVUELVE COMO TAL'''
	DATA_FRAME = pd.read_pickle(UBICACIÓN)
	return DATA_FRAME

def LIST_TO_XLSX(LISTA, COLUMNAS, UBICACIÓN):
	df = pd.DataFrame(LISTA, columns = COLUMNAS)
	df.to_excel(UBICACIÓN, index=False)

def DF_TO_PKL(DF, UBICACIÓN):
	''' DADO EL DATAFRAME Y UBICACIÓN CONVIERTE EL DATAFRAME EN PICKLE EN LA UBICACIÓN FUERA DEL PROGRAMA'''
	DF.to_pickle(UBICACIÓN)
	return(str(HERE + UBICACIÓN))


def LIST_TO_DF(LISTAS, COLUMNAS):
	''' DADA UNA LISTA COMO BDDD Y UNA LISTA COMO COLUMNAS LA CONVIERTE EN DATA FRAME'''
	BD = dict(zip(COLUMNAS, LISTAS))
	BDD = pd.DataFrame(LISTAS, columns = COLUMNAS)
	return BDD

def PICKLE_TO_XLS(FROM, WHERE):
    '''CONVIERTE UN ARCHIVO PIKLE EN EXCEL DADAS AMBAS UBICACIONES DE ARCHIVOS'''
    ARCHIVO = pd.read_pickle(FROM)
    ARCHIVO.to_excel(WHERE, sheet_name='CON')


def MOVE_BDD(FILES, WHERE):
	for FILE in FILES:
		os.path.splitext(FILE)


#----OPERACIONES DE BASES DE DATOS----

def almacen_bdd():
	ruta = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas"
	for root, dirs, files in os.walk(ruta):
		for file in files:
			basename, extention = os.path.splitext(file)
			RUTA_BDD = ROOT + file
			if extention == ".xls":
				if basename[:25] == "XXMCAN___Finiquito_de_Obr":
					pd.read_html(RUTA_BDD, header=0)[0]
				elif basename[:25] == "XXMC_Reportes_Transaccion":
					RTI.append(pd.read_html(RUTA_BDD, header=0)[0])

def CONCATENADOR_BDD():
    # DEFINICIONES
    root_reportes = get_here(folder="partial_files")
    root_rti = get_here(folder="BDD", name="BDD_Transacciones_de_Inventario.xlsx")
    root_rdf = get_here(folder="BDD", name="BDD_Reporte_de_Finiquitos.xlsx")
    root_rdfg = get_here(folder="BDD", name="BDD_Finiquito_de_Obra_Generales.xlsx")
    root_rec = get_here(folder="BDD", name="BDD_Reporte_de_Estado_de_Cuenta.xlsx")
    UBICACIÓN = os.getcwd() + "\\BDD\\"
    REC, RFO, RTI = [], [], []

    archivos_rdf = glob.glob(os.path.join(root_reportes, 'RDF*.xlsx'))
    archivos_rti = glob.glob(os.path.join(root_reportes, 'RTI*.xlsx'))
    archivos_rec = glob.glob(os.path.join(root_reportes, 'REC*.xlsx'))
    
    # Se revisan todos los archivos descargados y se agrupan los que coinciden en su texto
    for file in archivos_rdf:
        RFO.append(pd.read_excel(file))

    for file in archivos_rti:
        RTI.append(pd.read_excel(file))

    for file in archivos_rec:
        REC.append(pd.read_excel(file))

    # Intentamos convertir en data frames los exceles compilados
    try:
        DF_REC = pd.concat(REC, ignore_index=True)
    except ValueError:
        print("No hay REC")
    try:     
        DF_RFO = pd.concat(RFO, ignore_index=True)
    except ValueError:
        print("No hay RFO")
    try:        
        DF_RTI = pd.concat(RTI, ignore_index=True)
    except ValueError:
        print("No hay RTI")

    # Quitamos del REC todos los elementos de relleno
    DF_RFO_1 = DF_RFO.dropna(subset=['Estado']).copy()
    DF_RFO_2 = DF_RFO[DF_RFO['Estado'].isna()].copy()

    # Quitamos columnas faltantes
    DF_RFO_2.drop(['Contrato', 'Estado', 'Contratista', 'Contrato Legal'], axis=1, inplace=True)
    DF_RFO_1 = DF_RFO_1.drop(['Anticipo', 'Anticipo Amortizado', 'Anticipo por Amortizar', 
        "Incurrido de MAT", "Por incurrir MAT", "Ahorros pendientes", "Pendiente por contratar", 
        "Incurrido total", "Por incurrir total", "Estimado de cierre", "Total Finiquito"], axis=1)

    # Transferencia de descripción a RFO1
    # DF_RFO_1['Contrato Legal'].fillna('Falta_ubicar', inplace=True)
    DF_RFO_1['Contrato Legal'] = DF_RFO_1['Contrato Legal'].fillna('Falta_ubicar')

    mask = DF_RFO_1['Contrato Legal'].str.contains("_")
    DF_RFO_1.loc[mask, 'Contrato Legal'] = DF_RFO_1.loc[mask, 'Contrato Legal'].str.split("_").str[0]
    mask_contrato = DF_RFO_1['Contrato Legal'].isin(DF_RFO_1['Contrato'])

    # Obtener ruta de estatus de conjuntos
    root_eec = get_here(folder="BDD", name="BDD_Estatus_en_conjuntos.xlsx")
    df_eec = pd.read_excel(root_eec)
    DF_RFO_1 = DF_RFO_1.merge(df_eec[['Conjunto', 'Estado']], on='Conjunto', how='left', suffixes=('', '_Conjunto'))

    # Obtener comprador y descripción
    DF_RFO_1 = DF_RFO_1.merge(DF_REC[['Contrato', 'Descripción.1', 'Comprador']], on='Contrato', how='left')#, suffixes=('', '_Conjunto')

    DF_RFO_1 = DF_RFO_1.rename(columns={'Descripción.1': 'Descripción'})
    DF_RFO_1 = DF_RFO_1[['No. Proyecto', 'Frente', 'Conjunto', 'Contrato', 'Descripción', 'Contratista', 
         'Contrato Legal', 'Estado', 'Estado_Conjunto', 'Comprador', 'Total MO', 'Estimado', 
         'Por Estimar', 'Penalizado en Estimacion', 'Penalizado FG', 
         'Pendiente por penalizar estimación', 'Incurrido MO', 'Por incurrir MO']]

    DF_RTI = DF_RTI.drop_duplicates()
    DF_REC = DF_REC.drop_duplicates()
    DF_RFO_1 = DF_RFO_1.drop_duplicates()
    DF_RFO_2 = DF_RFO_2.drop_duplicates()

    # Convertimos los dataframes en archivos de excel
    DF_RTI.to_excel(root_rti, index=False)
    DF_RFO_1.to_excel(root_rdf, index=False)
    DF_RFO_2.to_excel(root_rdfg, index=False)
    DF_REC.to_excel(root_rec, index=False)

    print("Éxito al ordenar")

def copy_paste_reports():
    file_in = []
    file_name = (
        'BDD_Estatus_en_conjuntos.xlsx', 
        'BDD_Finiquito_de_Obra_Generales.xlsx', 
        'BDD_Reporte_de_Finiquitos.xlsx', 
        'BDD_Transacciones_de_Inventario.xlsx',
        'BDD_Transacciones_de_Inventario.xlsx'
        )
    root_in = 'C:\\ProgramData\\Autodesk\\PEX\\PRGS\\BDD\\'
    root_out = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\REPORTES\\'
    for file in file_name:
        file_path_in = os.path.join(root_in, file)
        file_path_out = os.path.join(root_out, file)
        shutil.copy(file_path_in, file_path_out)

def clean_bdd():
	bdd_ec = pd.read_excel(HERE + "\\BDD\\BDD_Estatus_en_conjuntos.xlsx", dtype={'Frentes': object})
	bdd_ec = bdd_ec.drop_duplicates(subset='Conjunto')
	bdd_ec.to_excel(HERE + "\\BDD\\BDD_Estatus_en_conjuntos.xlsx", index=False)

	bdd_fe = pd.read_excel(HERE + "\\BDD\\BDD_Frentes_existentes.xlsx", dtype={'Frentes': object})
	bdd_fe = bdd_fe.drop_duplicates(subset='Unicode')
	bdd_fe.to_excel(HERE + "\\BDD\\BDD_Frentes_existentes.xlsx", index=False)

def combinador_de_contratos_pdf():


    # Función para encontrar la mejor coincidencia en el DataFrame para un valor dado (value)
    def encontrar_mejor_coincidencia_en_dataframe(value, df):
        mejor_puntaje = 0
        mejor_coincidencia_descripcion = None
        mejor_coincidencia_codigo = None

        for index, row in df.iterrows():
            puntaje = fuzz.ratio(value, row['Descripción'])
            if puntaje > mejor_puntaje:
                mejor_puntaje = puntaje
                mejor_coincidencia_descripcion = row['Descripción']
                mejor_coincidencia_codigo = row['Código']

        return mejor_coincidencia_codigo, mejor_coincidencia_descripcion


# Función para encontrar la mejor coincidencia en la serie para un valor dado (value)
def encontrar_mejor_coincidencia_en_serie(value, serie):
    mejor_puntaje = 0
    mejor_coincidencia = None

    for descripcion in serie:
        puntaje = fuzz.ratio(value, descripcion)
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_coincidencia = descripcion

    return mejor_coincidencia


def vinculador(df_a, df_b):
    df_a.sort_values(by="Código", inplace=True)
    '''['Código', 'Descripción', 'UDM', 'Fecha \nCreación',
       'Plaza \nSolicitante', 'Solicitante', 'Código \nTipo', 'Categoría',
       'Código \nGrupo', 'GRUPO', 'Juego CAT INV 1', 'Juego CAT INV 2',
       'Juego PO ITEM', 'Tipo de Compra', 'Tipo Convenio',
       'Activo / \nInactivo', 'Fecha \nActualización \nInsumo',
       'Subcategoria (manual)']'''
    df_b = df_b[['Código', 'Descripción', 'Activo / \nInactivo']]
    df_b.rename(columns={'Activo / \nInactivo': 'Activo'}, inplace=True)
    df_b = df_b[df_b['Activo'] == 'Active']
    df_b = df_b[df_b['Código'].str.startswith(('4.03', '4.04'))]
    print(df_b)
    # df = df_b.iloc(df_b[''])

    def quitar_texto_despues_de_incluye(descripcion):
        return descripcion.split('INCLUYE')[0].strip()
    df_a['Concepto'] = df_a['Concepto'].apply(quitar_texto_despues_de_incluye)

    for value in df_a['Concepto']:
        cod, desc = encontrar_mejor_coincidencia_en_dataframe(value, df_b)
        print(cod, desc)


def buscador_de_insumos(route_bd, route_task, route_out):

    out_df = pd.read_excel(route_out)
    base_df = pd.read_excel(route_bd)
    task_df = pd.read_excel(route_task)

    task_df = task_df.drop_duplicates(subset='NOMBRE')
    task_df = task_df[task_df['TIPO'] == 'Subpaquete']
    task_df = task_df[~task_df['NOMBRE'].isin(out_df['task'])]

    valores_a_eliminar = ['4.01', '4.02', '4.03', '4.04', '4.05', '4.06']
    filtro = ~base_df['Código'].astype(str).str.startswith(tuple(valores_a_eliminar))
    base_df = base_df[filtro]

    filtro = base_df['UDM'].astype(str).str.startswith('LOT')
    base_df = base_df[filtro]

    new_db = []

    for i, row in task_df.iterrows():

        task = row['NOMBRE']
        code = row['ELEMENTO']
        code = str(code)
        new_code = code[:2]


        filtro = base_df['Código'].astype(str).str.startswith('4.' + new_code)
        best_base_df = base_df[filtro]

        other_base_df = base_df[~filtro]

        best_7 = encontrar_coincidencias(best_base_df, 'Descripción', task, top_n=12)
        other_3 = encontrar_coincidencias(other_base_df, 'Descripción', task, top_n=12)

        concat = pd.concat([best_7, other_3], ignore_index=True)
        concat = concat.set_index('Código')['Descripción'].to_dict()

        choose = window_input("Elección de insumos", task, concat)

        resultado = {
                    'task': task,
                    'code': int(code),
                    'choose': choose
                    }

        new_db.append(resultado)
        new_df = pd.DataFrame(new_db)
        new_df = pd.concat([new_df, out_df], ignore_index=True)
        new_df.to_excel(route_out, index=False)



root_a = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\codigos.xlsx"
root_b = "C:\\Users\\fprado\\REPORTES\\Programas\\PROYECTO_X\\BDD\\BDD_Insumos_IV.xlsx"
path = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\CONTRATOS"


# df_a = pd.read_excel(root_a)
# df_b = pd.read_excel(root_b)

# vinculador (df_a, df_b)