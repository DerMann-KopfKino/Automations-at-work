import os
import math
import tempfile
import mplcyberpunk
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from icecream import ic
from Multiherramienta import *
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image, SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


def listado_contratos_sin_finiquitos():
	# Inputs de rutas y datos
	ruta_2 = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\'
	ruta_reportes = "C:\\Users\\fprado\\REPORTES\\"
	ruta_excel = ruta_2 + "EJJ.xlsm"
	ruta_estado = ruta_reportes + "estado.xlsx"

	fracc_existentes = ["CUMBRE ALTA", "PASEO SAN JUNIPERO", "RANCHO EL SIETE", "VALLE DE SANTIAGO"]

	# Dataframes
	df_viejo_c = pd.read_excel(ruta_excel, sheet_name="C")
	df_viejo_f = pd.read_excel(ruta_excel, sheet_name="F")
	df_estado = pd.read_excel(ruta_estado)

	# Se filtran los que no existan en finiquitos
	filtro = ~df_viejo_c['Contrato'].isin(df_viejo_f['Contrato'])
	sin_finiquito = df_viejo_c.loc[filtro]

	# Se filtran los valores que no correspondan a fraccionamientos propios
	filtro2 = sin_finiquito['Fraccionamiento'].isin(fracc_existentes)
	lista_final = sin_finiquito.loc[filtro2]

	# Se agrega el valor de coincidencia en estado
	# Ensure lista_final is a copy if it's a slice
	lista_final = lista_final.copy()
	existe = lista_final['Contrato'].isin(df_estado['Contrato Legal']).to_frame()
	lista_final.loc[:, "Existe en oracle"] = existe.values

	# Output
	print(f'El listado de Contratos_sin_finiquitos.xlsx: {lista_final.shape[0]}')
	lista_final.to_excel(ruta_reportes + "Contratos_sin_finiquitos.xlsx", index=False)


def listado_de_contratos_inexistentes():
	# Definiciones
	ruta_juridico =  "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\EJJ.xlsm"
	ruta_oracle = "C:\\Users\\fprado\\REPORTES\\estado.xlsx"

	# Dataframes
	df_viejo_c = pd.read_excel(ruta_juridico, sheet_name="C")
	df_viejo_f = pd.read_excel(ruta_juridico, sheet_name="F")
	df_oracle = pd.read_excel(ruta_oracle)

	# Filtros
	mis_fracc = ["CUMBRE ALTA", "CUMBRE ALTA ELITE", "PASEO SAN JUNIPERO", "RANCHO EL SIETE", "MARQUES DEL RIO"]
	filtro_contratos_mios = df_viejo_c['Fraccionamiento'].isin(mis_fracc)
	contratos_mios = df_viejo_c.loc[filtro_contratos_mios]

	filtro_contratos_finiquitados =contratos_mios['Contrato'].isin(df_viejo_f['Contrato'])
	contratos_mios_sin_finiquito = contratos_mios.loc[~filtro_contratos_finiquitados]

	filtro_contratos_existentes = contratos_mios_sin_finiquito['Contrato'].isin(df_oracle['Contrato Legal'])
	contratos_mios_inexistentes = contratos_mios_sin_finiquito.loc[~filtro_contratos_existentes]
	print(f'El listado de Contratos_inexistentes.xlsx: {contratos_mios_inexistentes.shape[0]}')

	contratos_mios_inexistentes.to_excel("C:\\Users\\fprado\\REPORTES\\Contratos_inexistentes.xlsx", index=False)


def listado_contratos_listos_para_finiquitar():

	# Leemos la base de datos de los estados de cuenta de lista_contratos
	ruta_bdd = get_here(folder="BDD", name="BDD_Reporte_de_Finiquitos.xlsx")
	ruta_one_drive = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\"
	ruta_reportes = "C:\\Users\\fprado\\REPORTES\\"
	ruta_excel = os.path.join(ruta_one_drive, "EJJ.xlsm")

	# Leemos los archivos de jurídico
	df_mi_reporte_de_finiquitos = pd.read_excel(ruta_bdd)
	df_por_finiquitar = pd.read_excel(ruta_reportes + "Por_finiquitar.xlsx")

	df_nuevo_contratos = pd.read_excel(ruta_excel, sheet_name="N")
	df_nuevo_modificatorios = pd.read_excel(ruta_excel, sheet_name="MN")
	df_nuevo_finiquitos = pd.read_excel(ruta_excel, sheet_name="FN")

	df_viejo_contratos = pd.read_excel(ruta_excel, sheet_name="C")
	df_viejo_modificatorios = pd.read_excel(ruta_excel, sheet_name="M")
	df_viejo_finiquitos = pd.read_excel(ruta_excel, sheet_name="F")

	df_correccion_contratos = pd.read_excel(ruta_reportes + "correccion.xlsx")

	df_contratistas_oracle = df_mi_reporte_de_finiquitos['Contratista'].drop_duplicates()
	df_contratistas_oracle.sort_values(inplace=True)
	
	# Filtramos mis organizaciones
	filtro_mis_organizaciones = ["URC", "CRÑ", "CR7", "UR7", "CMA", "UMA", "CJM", "PME"]
	filtro_mi_reporte_de_finiquitos = df_mi_reporte_de_finiquitos['No. Proyecto'].isin(filtro_mis_organizaciones)
	df_mi_reporte_de_finiquitos = df_mi_reporte_de_finiquitos.loc[filtro_mi_reporte_de_finiquitos]

	# Crea un diccionario de mapeo a partir de df_correccion_contratos['Contrato'] y df_correccion_contratos['Contrato Legal']
	mapeo_de_correción_de_contratos = df_correccion_contratos.set_index('Contrato')['Contrato Legal'].to_dict()
	df_mi_reporte_de_finiquitos['Contrato Legal'] = df_mi_reporte_de_finiquitos['Contrato'].map(mapeo_de_correción_de_contratos).fillna(df_mi_reporte_de_finiquitos['Contrato Legal'])

	lista_todos_finiquitos = pd.concat([df_viejo_finiquitos['Contrato'], df_nuevo_finiquitos['Contrato']], axis=0)

	# Creamos una columna que indique los que ya están finiquitados
	filtro_lista_de_finiquitos = df_mi_reporte_de_finiquitos['Contrato Legal'].isin(lista_todos_finiquitos)
	df_mi_reporte_de_finiquitos['Finiquito'] = np.where(filtro_lista_de_finiquitos, 'True', 'False')
	df_mi_reporte_de_finiquitos['Contrato Legal'] = df_mi_reporte_de_finiquitos['Contrato Legal'].str.split('_A').str[0]
	df_mi_reporte_de_finiquitos['Contrato Legal'] = np.where(df_mi_reporte_de_finiquitos['Contrato Legal'].str.contains('car|pnz', case=False, na=False), 'CARGO', df_mi_reporte_de_finiquitos['Contrato Legal'])
	df_mi_reporte_de_finiquitos['UltimosDigitos'] = df_mi_reporte_de_finiquitos['Contrato Legal'].str.extract(r'(\d+)$')

	# Creamos un dataframe con todos los lista_contratos y coratistas existentes
	contratos_existentes = pd.concat([df_nuevo_contratos[['Contrato', 'Contratista', 'Conjunto']], df_viejo_contratos[['Contrato', 'Contratista', 'Conjunto']], df_viejo_finiquitos[['Contrato', 'Contratista', 'Conjunto']]], ignore_index=True)
	contratos_existentes.drop_duplicates(subset='Contrato', inplace=True)
	contratos_existentes['UltimosDigitos'] = contratos_existentes['Contrato'].str.extract(r'(\d+)$')
	contratos_existentes.to_excel(ruta_reportes + "lista_contratos existentes.xlsx", index=False)
	
	# Creamos una columna que indique la existencia de dicho contrato
	filtro_contratos_existentes = df_mi_reporte_de_finiquitos['UltimosDigitos'].isin(contratos_existentes['UltimosDigitos'])
	df_mi_reporte_de_finiquitos['Existencia'] = np.where(filtro_contratos_existentes, True, False)

	# Filtramos los lista_contratos existentes a los que importan de la base de datos
	filtro_contratos_existentes = contratos_existentes['UltimosDigitos'].isin(df_mi_reporte_de_finiquitos['UltimosDigitos'])
	contratos_existentes = contratos_existentes.loc[filtro_contratos_existentes]

	# Filtramos los coincidentes
	contratos_existentes['Contratista'] = contratos_existentes['Contratista'].astype(str)
	contratos_existentes['Contratista'] = contratos_existentes['Contratista'].apply(lambda x: mejor_coincidencia(x, df_contratistas_oracle))

	# Combinamos los DataFrames usando la columna 'UltimosDigitos' como clave
	df_estado_contratos = pd.merge(df_mi_reporte_de_finiquitos, contratos_existentes, on='UltimosDigitos', how='left', suffixes=('', '_Jurídico'))

	# Utilizamos np.where() para verificar las condiciones y asignar los valores correspondientes
	df_mi_reporte_de_finiquitos = (df_estado_contratos['Contratista'] != df_estado_contratos['Contratista_Jurídico'])
	df_estado_contratos['Iguales'] = np.where(df_mi_reporte_de_finiquitos, 'Error', 'Correcto')
	df_estado_contratos['Iguales'] = np.where(df_estado_contratos['Contratista'].str.contains('fabiola|abraham', case=False, na=False), 'Correcto', df_estado_contratos['Iguales'])
	filtro_Real_vs_finiquitor_nuevos_y_viejos_contratos = (df_estado_contratos['Contrato Legal'].isin(df_nuevo_contratos['Contrato']))
	df_estado_contratos['Sistema'] = np.where(filtro_Real_vs_finiquitor_nuevos_y_viejos_contratos, 'Nuevo', 'Viejo')

	# Guardamos el excel
	print(f'El listado de estado.xlsx: {df_estado_contratos.shape[0]}')
	ruta_estado = os.path.join(ruta_reportes, "estado.xlsx")
	df_estado_contratos.to_excel(ruta_estado, index=False)

	# Creamos una lista de factibilidad de finiquito
	ruta_estatus_conjuntos = get_here(folder="BDD", name="BDD_Estatus_en_conjuntos.xlsx")
	df_estatus_conjuntos = pd.read_excel(ruta_estatus_conjuntos)
	lista_contratos = df_estado_contratos['Contrato Legal'].drop_duplicates()
	lista_finiquitables = []

	# Creamos resumen de estado de contratos, por cada contrato
	for contrato in lista_contratos:
		fila = []
		data = df_estado_contratos[df_estado_contratos['Contrato Legal'] == contrato]

		if len(data) != 0:

			estimado = data['Estimado']
			por_estimar = data['Por Estimar']
			estado_contrato = data['Estado']

			Total_estimado = estimado.sum()
			Total_pendinete = por_estimar.sum()

			if estado_contrato.all():
				estado_contrato = estado_contrato.drop_duplicates().values[0]
			elif 'En Proceso de Configuración' in estado_contrato or 'En proceso de Definición' in estado_contrato:
				estado_contrato = 'En Proceso'
			elif "Cancelado" in estado_contrato:
				estado_contrato = 'Posible'

		fila.append(contrato)
		fila.append(estado_contrato)
		fila.append(Total_estimado)
		fila.append(Total_pendinete)
		lista_finiquitables.append(fila)

	# Creamos un dataframe nuevo con la suma de información
	nombre_de_columnas = ['Contrato', 'Estado_contrato', 'Total_estimado', 'Total_por_estimar']
	df_finiquitables = pd.DataFrame(lista_finiquitables, columns=nombre_de_columnas)

	# Limpiamos el estado de contratos de datos innecesarios
	columnas_a_borrar_dfec = ['Total MO', 'Estimado', 'Por Estimar', 'Penalizado en Estimacion', 'Penalizado FG', 
						'Pendiente por penalizar estimación', 'Incurrido MO', 'Por incurrir MO', 
						'Contrato_Jurídico', 'Contratista_Jurídico', 'Conjunto_Jurídico', 'Iguales']
	df_estado_contratos = df_estado_contratos.drop(columnas_a_borrar_dfec, axis=1)
	df_estado_contratos = df_estado_contratos.drop_duplicates(subset=['Contrato Legal'])
	df_estado_contratos = df_estado_contratos.sort_values(by=['Contrato'], ascending=[True])

	# Concatenamos todos los reportes para conocer la última fecha de modificación
	df_ultima_modificacion = pd.concat(
		[
			df_nuevo_contratos[['Contrato', 'Fecha creación', 'Fecha terminación']], 
			df_nuevo_modificatorios[['Contrato', 'Fecha creación', 'Fecha terminación']], 
			df_nuevo_finiquitos[['Contrato', 'Fecha creación', 'Fecha terminación']],
			df_viejo_contratos[['Contrato', 'Fecha creación', 'Fecha terminación']],
			df_viejo_modificatorios[['Contrato', 'Fecha creación', 'Fecha terminación']],
			df_viejo_finiquitos[['Contrato', 'Fecha creación', 'Fecha terminación']],
		],
		keys=['Contrato', 'Finiquito', 'Modificatorio', 'Contrato', 'Finiquito', 'Modificatorio'])

	# Ordenamos de forma descendente la última modificación para quedarnos con el dato mas reciente
	df_ultima_modificacion.reset_index(level=0, inplace=True)   
	df_ultima_modificacion['Fecha creación'] = pd.to_datetime(df_ultima_modificacion['Fecha creación'])
	df_ultima_modificacion = df_ultima_modificacion.sort_values(by=['Contrato', 'Fecha creación'], ascending=[True, False])
	df_ultima_modificacion = df_ultima_modificacion.drop_duplicates(subset=['Contrato'], keep='first')

	# Concatenamos reportes de contratos viejo y nuevo sistema
	df_contratos_concat = pd.concat(
		[
			df_nuevo_contratos[['Contrato', 'Descripción de trabajo']],
			df_viejo_contratos[['Contrato', 'Descripción de trabajo']]
		]
		)

	# Concatenamos reportes de finiquitos viejo y nuevo sistema
	df_finiquitos_concat = pd.concat(
	[
		df_nuevo_finiquitos[['Contrato', 'Importe contrato']],
		df_viejo_finiquitos[['Contrato', 'Importe contrato']]
	]
	)
	df_finiquitos_concat = df_finiquitos_concat.rename(columns={'Importe contrato': 'Importe_finiquito'})

	# Agregamos los datos del conjunto
	df_finiquitables = df_finiquitables.merge(
			df_estado_contratos, 
			left_on='Contrato', 
			right_on='Contrato Legal', 
			how='left'
		)

	# Agregamos datos de la última vez que se modifico el contrato
	df_finiquitables = df_finiquitables.rename(columns={'Contrato_x': 'Contrato'})
	df_finiquitables = df_finiquitables.merge(
			df_ultima_modificacion, 
			left_on='Contrato', 
			right_on='Contrato', 
			how='left'
		)

	# Agregamos la descripción del contrato jurídico
	df_finiquitables = df_finiquitables.merge(
		df_contratos_concat, 
		left_on='Contrato', 
		right_on='Contrato', 
		how='left'
	)

	# Agregamos los datos de los finiquitos
	df_finiquitables = df_finiquitables.merge(
		df_finiquitos_concat, 
		left_on='Contrato', 
		right_on='Contrato', 
		how='left'
	)

	# Cambiamos de texto a boleano en Finiquito
	df_finiquitables['Finiquito'] = df_finiquitables['Finiquito'].map({'True': True, 'False': False})

	# Crear la nueva columna 'Vencidos' con las condiciones especificadas
	df_finiquitables['Vencidos'] = np.where(
		(df_finiquitables['Finiquito'] == False) & 
		(df_finiquitables['Estado_contrato'] == 'Publicado') & 
		(df_finiquitables['Fecha terminación'] < pd.Timestamp('today')),
		True,
		False
	)

	# Crear comparativa de estimado vs monto en finiquito
	df_finiquitables['Real_vs_finiquito'] = df_finiquitables['Total_estimado'] - df_finiquitables['Importe_finiquito']
	df_finiquitables['Real_vs_finiquito'] = np.where(
		(df_finiquitables['Real_vs_finiquito'] >= -30) & (df_finiquitables['Real_vs_finiquito'] <= 30), 
		0, 
		df_finiquitables['Real_vs_finiquito']
	)

	# Creación de columna Estatus_contrato
	df_finiquitables['Estatus_contrato'] = df_finiquitables.apply(lambda row: 'Finiquitado' if row['Finiquito'] 
						else ('Vencido' if row['Vencidos'] else 'Abierto'), axis=1)

	# Limpieza de columnas sobrantes
	columnas_a_borrar_dfec = ['Contrato_y', 'Descripción', 'Comprador', 'Contrato Legal', 'Estado', 'Finiquito', 'Vencidos']
	df_finiquitables = df_finiquitables.drop(columnas_a_borrar_dfec, axis=1)

	# Renombra las columnas a un orden mas manejable en
	df_finiquitables.rename(columns={
		'No. Proyecto': 'Organización', 
		'Descripción de trabajo': 'Descripción',
		'level_0': 'Fuente_de_datos',
		'Fecha creación': 'Ultima_modificación', 
		'Fecha terminación': 'Fecha_de_terminación' 
		}, inplace=True)

	# Definir las condiciones para cada valor en la columna 'fraccionamiento'
	conditions = [
	    df_finiquitables['Organización'].isin(['CRÑ', 'URC']),
	    df_finiquitables['Organización'].isin(['CR7', 'UR7']),
	    df_finiquitables['Organización'].isin(['CMA', 'UMA']),
	    df_finiquitables['Organización'] == 'CJM',
	    df_finiquitables['Organización'] == 'PME'
	]

	# Definir los valores correspondientes para cada condición
	choices = [
	    'Paseo San Junípero',
	    'Pedregal del Río',
	    'Cumbre Alta',
	    'Marqués del Río',
	    'Cumbre Alta Élite'
	]

	# Crear la nueva columna 'fraccionamiento' usando np.select
	df_finiquitables['Fraccionamiento'] = np.select(conditions, choices, default='Otro')

	columnas_ordenadas_dfec = [
		'Fraccionamiento', 'Organización', 'Frente', 'Conjunto', 'Contrato', 'Contratista',  
		'Descripción', 'Estado_contrato', 'Total_estimado', 'Total_por_estimar', 
		'Estado_Conjunto', 'Estatus_contrato', 'Importe_finiquito', 'Real_vs_finiquito', 'Fuente_de_datos', 
		'Ultima_modificación', 'Fecha_de_terminación', 'UltimosDigitos', 'Existencia', 'Sistema',
	]
	df_finiquitables = df_finiquitables[columnas_ordenadas_dfec]

	fechas_columnas = ['Ultima_modificación', 'Fecha_de_terminación']
	for col in fechas_columnas:
		df_finiquitables[col] = pd.to_datetime(df_finiquitables[col], errors='coerce').dt.strftime('%d/%m/%Y')

	print(f'El listado de Finiquitables.xlsx: {df_finiquitables.shape[0]}')
	df_finiquitables.to_excel("C:\\Users\\fprado\\REPORTES\\Finiquitables.xlsx", index=False)

# Función auxiliar para añadir gráficos al PDF en una posición específica
def add_plot_to_pdf(c, fig, x, y, w, h):
	# y = y + h
	with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
		fig.savefig(tmpfile.name, format='png', bbox_inches='tight')
		tmpfile.close()
		c.drawImage(tmpfile.name, x, y, width=w, height=h)
		os.remove(tmpfile.name)

def add_dataframe_to_pdfalfa(c, data, style_colors, x, y):
	is_df = isinstance(data, pd.DataFrame)
	is_s = isinstance(data, pd.Series)
	
	# Verificar si data es un DataFrame o una Serie
	if is_df:
		# Convertir DataFrame a lista de listas
		data_list = [data.columns.tolist()] + data.values.tolist()
	elif is_s:
		# Convertir Serie a lista de listas
		data_list = list(zip(data.index, data.values))
	else:
		raise ValueError("El dato debe ser un DataFrame o una Serie de Pandas.")
	
	# Crear la tabla
	table = Table(data_list)

	# Estilo de la tabla
	style = TableStyle(style_colors)
	table.setStyle(style)

	# Colocar la tabla en el canvas
	table.wrapOn(c, 500, 200)  # Ajusta el tamaño según sea necesario
	table.drawOn(c, x, y)

def add_dataframe_to_pdf(c, data, style_colors, x, y_start_from_top):
    # Verificar si data es un DataFrame o una Serie
    is_df = isinstance(data, pd.DataFrame)
    is_s = isinstance(data, pd.Series)

    if is_df:
        # Convertir DataFrame a lista de listas
        data_list = [data.columns.tolist()] + data.values.tolist()
    elif is_s:
        # Convertir Serie a lista de listas
        data_list = list(zip(data.index, data.values))
    else:
        raise ValueError("El dato debe ser un DataFrame o una Serie de Pandas.")

    # Crear la tabla
    table = Table(data_list)

    # Estilo de la tabla
    style = TableStyle(style_colors)
    table.setStyle(style)

    # Calcular altura de la tabla
    table_width, table_height = table.wrap(0, 0)  # Calcula dimensiones de la tabla

    # Ajustar la posición `y` desde la parte superior
    page_height = c._pagesize[1]  # Altura de la página
    y = page_height - y_start_from_top - table_height

    # Dibujar la tabla en el canvas
    table.drawOn(c, x, y)

def plot_my_reports():
	#Theme
	plt.style.use("cyberpunk")

	# Cargar el DataFrame
	df = pd.read_excel("C:\\Users\\fprado\\REPORTES\\Finiquitables.xlsx")

	# Configurar PDF en orientación horizontal
	pdf_path = "C:\\Users\\fprado\\REPORTES\\Reporte_de_Contratos.pdf"
	c = canvas.Canvas(pdf_path, pagesize=landscape(letter))
	width, height = landscape(letter)
	x_line, y_line = width / 36, height / 24

	# Colores predefinidos
	dark = '#212946'
	bright = '#FFFFFF'
	green = '#00ff28'
	yellow = '#ffa300'
	pink = '#ff0534'
	cyan = '#08F7FE'
	magenta = '#ff0534'
	carmesi = '#e51a4c'

	# Estilos
	style_colors_graph = [cyan, yellow, magenta]
	style_table_series = [
		# Outside
		('BOX', (0, 0), (-1, -1), 1.5, cyan),
		# Head
		('BACKGROUND', (0, 0), (-1, 0), cyan),
		('FONTSIZE', (0, 0), (-1, 0), 12),
		('TEXTCOLOR', (0, 0), (-1, 0), dark),
		('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
		('ALIGN', (0, 0), (-1, 0), 'CENTER'),
		('BOTTOMPADDING', (0, 0), (-1, 0), 10),
		# Content rows
		('INNERGRID', (0, 0), (-1, -1), 0.01, cyan),
		('BACKGROUND', (0, 1), (-1, -1), dark),
		('TEXTCOLOR', (0, 1), (-1, -1), bright),
		('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
		# First 4 columns
		('ALIGN', (1, 1), (-2, -1), 'CENTER'),
		# Last column
		('ALIGN', (-1, 1), (-1, -1), 'RIGHT'),
	]

	style_table_I = [
		# Outside
		('BOX', (0, 0), (-1, -1), 1.5, cyan),
		# Head
		('BACKGROUND', (0, 0), (-1, 0), cyan),
		('FONTSIZE', (0, 0), (-1, 0), 10),
		('TEXTCOLOR', (0, 0), (-1, 0), dark),
		('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
		('ALIGN', (0, 0), (-1, 0), 'CENTER'),
		# Content rows
		('INNERGRID', (0, 0), (-1, -1), 0.5, cyan),
		('BACKGROUND', (0, 1), (-1, -1), dark),
		('TEXTCOLOR', (0, 1), (-1, -1), bright),
		('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
		('FONTSIZE', (0, 1), (-1, -1), 8),
		# First 2 columns
		('ALIGN', (1, 1), (2, -1), 'CENTER'),
		# Descripción
		('WORDWRAP', (2, 1), (3, -1))
	]

	# Configurar el color de fondo
	c.setFillColor(dark)
	c.rect(0, 0, *landscape(letter), fill=1)

	# Título de página
	title_font_size = 20
	title_text = "Reporte de avance de finiquitos"
	c.setFont("Helvetica-Bold", title_font_size)
	text_width = c.stringWidth(title_text, "Helvetica-Bold", title_font_size)
	x = (width - text_width) / 2
	y = height - 40
	c.setFillColor(bright)
	c.drawString(x, y, title_text)

	# Gráfica de pastel general
	counts_general = df['Estatus_contrato'].value_counts()
	counts_general_per = df['Estatus_contrato'].value_counts().to_frame(name='Cantidad').reset_index()
	counts_general_per.columns = ['Estatus', 'Cantidad']
	counts_general_per['Porcentaje'] = ((counts_general_per['Cantidad'] / counts_general_per['Cantidad'].sum()) * 100).round(2)
	fig, ax = plt.subplots(figsize=(8, 8))
	counts_general.plot(
	    kind='pie',
	    autopct='%1.1f%%',
	    colors=style_colors_graph,
	    startangle=180,
	    ax=ax,
	    textprops=dict(color=dark)
	)
	ax.set_title('Distribución General de Contratos', fontsize=20)
	ax.set_ylabel('')
	ax.legend(title='Estatus')
	add_plot_to_pdf(c, fig, x=20, y=220, w=300, h=300)
	plt.close(fig)

	# Tabla general de contratos	
	add_dataframe_to_pdf(c, counts_general_per, style_table_series, 45, 400)

	# Gráfica de barras apiladas para fraccionamiento
	estado_por_org = df.pivot_table(index='Fraccionamiento', columns='Estatus_contrato', aggfunc='size', fill_value=0)
	fig, ax = plt.subplots(figsize=(10, 8))
	estado_por_org[['Finiquitado', 'Abierto', 'Vencido']].plot(
		kind='barh', 
		stacked=True, 
		color=style_colors_graph, 
		ax=ax
	)
	ax.set_title('Distribución de Contratos por Organización', fontsize=22)
	ax.set_xlabel('Número de contratos')
	ax.set_ylabel('Fraccionamientos')
	ax.xaxis.set_major_locator(ticker.MultipleLocator(100))
	ax.xaxis.set_minor_locator(ticker.MultipleLocator(50))
	ax.grid(color=carmesi, axis='x', linewidth=0.5)
	ax.legend(title='Estatus')
	add_plot_to_pdf(c, fig, x=360, y=240, w=400, h=280)
	plt.close(fig)

	# Tabla de fraccionamientos
	estado_por_org = estado_por_org.reset_index()
	estado_por_org['Cumplimiento'] = ((
		(estado_por_org['Finiquitado'] + estado_por_org['Abierto']) 
		/ (estado_por_org['Finiquitado'] + estado_por_org['Abierto'] + estado_por_org['Vencido'])
		)*100
		).round(2)
	estado_por_org = estado_por_org[['Fraccionamiento', 'Finiquitado', 'Abierto', 'Vencido', 'Cumplimiento']]
	add_dataframe_to_pdf(c, estado_por_org, style_table_series, 370, 400)

	c.showPage()

	# Dataframe
	columnas = ['Fraccionamiento',  
				'Conjunto', 
				'Contrato', 
				'Contratista',
				# 'Descripción',
				'Estatus_contrato',
				'Fecha_de_terminación',
				'Total_estimado', 
				'Total_por_estimar', 
				'Estado_Conjunto']
	df_top_vencidos = df[columnas]
	df_top_vencidos = df_top_vencidos.rename(columns={
		'Fecha_de_terminación': 'Fecha', 
		'Total_estimado': 'Estimado',
		'Total_por_estimar': 'Pendiente',
		'Estado_Conjunto': 'Estado' 
		})
	df_top_vencidos = df_top_vencidos[df_top_vencidos['Estatus_contrato'] == 'Vencido']
	df_top_vencidos['Fecha'] = pd.to_datetime(df_top_vencidos['Fecha'])
	df_top_vencidos['Fecha'] = df_top_vencidos['Fecha'].dt.strftime('%d/%m/%y')
	df_top_vencidos = df_top_vencidos.sort_values(by=['Fraccionamiento', 'Fecha'])
	top_10_por_fraccionamiento = df_top_vencidos.groupby('Fraccionamiento').head(17)
	dataframes_por_fraccionamiento = {}
	# Agrupar por la columna 'Fraccionamiento'
	for fraccionamiento, grupo in top_10_por_fraccionamiento.groupby('Fraccionamiento'):
	    # Guardar el DataFrame del grupo en el diccionario
	    dataframes_por_fraccionamiento[fraccionamiento] = grupo

	# Ahora, cada fraccionamiento tiene su propio DataFrame
	for fraccionamiento, df in dataframes_por_fraccionamiento.items():
		# Relleno del fondo
		df = df.drop(columns=['Estatus_contrato', 'Fraccionamiento'])
		df = df.sort_values(by='Conjunto')

		# Definir los textos que deseas buscar
		grupo_1 = ['E04', 'C03']
		grupo_2 = ['I01', 'U02', 'P05']

		# Filtrar los valores que contienen los textos en grupo_1
		df_1 = df[df['Conjunto'].str.contains('|'.join(grupo_1))]

		# Filtrar los valores que contienen los textos en grupo_2
		df_2 = df[df['Conjunto'].str.contains('|'.join(grupo_2))]

		c.setFillColor(dark)
		c.rect(0, 0, *landscape(letter), fill=1)

		# Título de página
		title_font_size = 20
		c.setFont("Helvetica-Bold", title_font_size)
		text_width = c.stringWidth(fraccionamiento, "Helvetica-Bold", title_font_size)
		x = (width - text_width) / 2
		y = height - 40
		c.setFillColor(bright)
		c.drawString(x, y, fraccionamiento)

		if not df_1.empty:
			# Título
			title_font_size = 16
			c.setFont("Helvetica-Bold", title_font_size)
			c.setFillColor(bright)
			y = height - 70
			c.drawString(50, y, 'Edificación y Equipamiento')
			add_dataframe_to_pdf(c, df_1, style_table_I, 50, 80)
		if not df_2.empty and not df_1.empty:
			# Título
			title_font_size = 16
			c.setFont("Helvetica-Bold", title_font_size)
			c.setFillColor(bright)
			y = height - 400
			c.drawString(50, y, 'Urbanización, Plataformas e Infraestructura')
			add_dataframe_to_pdf(c, df_2, style_table_I, 50, 410)
		elif not df_2.empty and df_1.empty:
			# Título
			title_font_size = 18
			c.setFont("Helvetica-Bold", title_font_size)
			c.setFillColor(bright)
			y = height - 70
			c.drawString(50, y, 'Urbanización')

			add_dataframe_to_pdf(c, df_1, style_table_I, 50, 80)
		c.showPage()

	# Guardar y cerrar el PDF
	c.save()

def actualizar_BDD_de_mis_fraccionamientos():
	listado_contratos_sin_finiquitos()
	listado_de_contratos_inexistentes()
	listado_contratos_listos_para_finiquitar()

# listado_contratos_listos_para_finiquitar()

actualizar_BDD_de_mis_fraccionamientos()
plot_my_reports()