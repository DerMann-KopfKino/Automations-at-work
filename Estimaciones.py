import os
import time
import tkinter as tk
import pandas as pd
import win32com.client as win32

def manda_estimaciones():
    root_caratulas = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Descargas\\CARATULAS\\'
    root_correos = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Contratistas_correos.xlsx'
    root_javer_correos = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Colaboradores_correos.xlsx'
    
    def capturar_datos(ventana, contratista, correo_entry, df_correos):
        correos = correo_entry.get()
        dic_correos = {"Contratistas": contratista, "Correos": correos}
        row_correos = pd.DataFrame([dic_correos])
        print(row_correos)
        df_correos = pd.concat([df_correos, row_correos])
        df_correos.drop_duplicates(inplace=True)
        print(df_correos)
        df_correos.to_excel(root_correos, index=False)
        # Destroy the window after capturing the data
        ventana.destroy()

    def crear_ventana(contratista, df_correos):
        ventana = tk.Tk()
        ventana.title("Captura de Datos")

        # Etiqueta para mostrar el nombre del contratista
        contratista_label = tk.Label(ventana, text=f"Contratista: {contratista}")
        contratista_label.pack()

        # Campo de entrada para los correos
        correo_label = tk.Label(ventana, text="Correos:")
        correo_label.pack()
        correo_entry = tk.Entry(ventana)
        correo_entry.pack()

        # Botón para capturar los datos
        capturar_button = tk.Button(ventana, text="Capturar", command=lambda: capturar_datos(ventana, contratista, correo_entry, df_correos))
        capturar_button.pack()

        ventana.mainloop()

    def crear_correo(archivos):
        """
        Crea un elemento de correo con los archivos especificados.

        Args:
        archivos: Una lista de archivos a adjuntar.

        Returns:
        Un elemento de correo con los archivos adjuntos.
        """

        # Crear un objeto de correo.
        mail = win32.Dispatch("Outlook.Application").CreateItem(0)

        # Agregar los archivos adjuntos.
        for archivo in archivos:
            mail.Attachments.Add(archivo)

        # Devolver el elemento de correo.
        return mail

    # Obtener la lista de archivos.
    archivos = os.listdir(root_caratulas)
    df_javer = pd.read_excel(root_javer_correos)
    try:
        df_correos = pd.read_excel(root_correos)
    except FileNotFoundError:
        df_correos = pd.DataFrame(columns=["Contratistas", "Correos"])

    # Agrupar los archivos por nombre
    archivos_agrupados = {}
    for archivo in archivos:
        nombre_archivo, extension = os.path.splitext(archivo)
        root_archivo = os.path.join(root_caratulas, archivo)
        contratista = nombre_archivo.split(" - ")[0]
        if contratista not in archivos_agrupados:
            archivos_agrupados[contratista] = []
        archivos_agrupados[contratista].append(root_archivo)


    # Crear un elemento de correo para cada grupo de archivos.
    for contratista, archivos_en_grupo in archivos_agrupados.items():
        print(contratista)
        print(archivos_en_grupo)
        destinatarios = ["ppalacios@javer.com.mx", "mphernandez@javer.com.mx", "jcastillo@javer.com.mx"]
        for archivo in archivos_en_grupo:
            ruta, archivo_nombre = os.path.split(archivo)
            nombre, extension = os.path.splitext(archivo_nombre)
            partes_nombre = nombre.split(" - ")
            conjunto = partes_nombre[1]
            partes_conjunto = conjunto.split("-")
            org = partes_conjunto[0]
            frente = partes_conjunto[1]
            etapa = partes_conjunto[2]
            coordinador = df_javer['']
            control_do =



        mail = crear_correo(archivos_en_grupo)
        if contratista not in df_correos['Contratistas'].values:
            crear_ventana(contratista, df_correos)

        correos_contratista = df_correos.loc[df_correos["Contratistas"] == contratista, "Correos"][0]
        print(correos_contratista)
        # Establecer el destinatario.
        mail.Body = "Este es el cuerpo del mensaje."
        semana = str(time.strftime('%W'))
        mail.Subject = "ESTIMACIONES " + semana
        mail.To = correos_contratista #buscar contratista de una lista
        mail.CC.Add(destinatario)

    # Mostrar el elemento de correo.
    print(mail.Subject)
    print(mail.Body)


manda_estimaciones()
