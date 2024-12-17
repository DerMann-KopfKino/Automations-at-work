import os
import pandas as pd
import win32com.client as win32

def leer_archivo_excel(ruta_archivo):
  """
  Lee un archivo Excel y lo convierte en un DataFrame.

  Args:
    ruta_archivo: Ruta del archivo Excel.

  Returns:
    DataFrame con los datos del archivo Excel.
  """

  df = pd.read_excel(ruta_archivo)
  return df

def descargar_archivos_adjunto(contrato, folio):
  """
  Descarga los archivos adjuntos de los correos de un determinado contrato.

  Args:
    contrato: Nombre del contrato.
    folio: Folio del documento adjunto.

  Returns:
    Nada.
  """

  # Crea la carpeta del contrato.
  print(contrato, folio)
  ruta_carpeta = f"{contrato}"
  if not os.path.exists(ruta_carpeta):
    os.mkdir(ruta_carpeta)

  # Inicia sesión en Outlook.

  outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
  inbox = outlook.GetDefaultFolder(6)
  # Busca los correos del contrato.
  CAO = inbox.Folders('CAO')
  correos = CAO.Folders('Facturas').items

  # Descarga los archivos adjuntos.

  for correo in correos:
    if correo.SenderEmailAddress.lower() == 'llaz_caveso_@hotmail.com' or correo.SenderEmailAddress.lower() == 'lcastaneda@caveso.com.mx':
      adjuntos = correo.Attachments
      for adjunto in adjuntos:
        nombre_archivo, extension = os.path.splitext(adjunto.FileName)
        print(nombre_archivo)
        if extension == ".pdf":
          if str(folio) in str(nombre_archivo):
            # Descarga el archivo adjunto
            print("Se encontró este", folio)
            adjunto.SaveAsFile(f"{ruta_carpeta}/{adjunto.FileName}")
          elif folio.lower() in str(nombre_archivo).lower():
            # Descarga el archivo adjunto
            print("Se encontró este", folio)
            adjunto.SaveAsFile(f"{ruta_carpeta}/{adjunto.FileName}")


def main():
  # Lee el archivo Excel.

  ruta_archivo = "C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\LISTA_FACTURAS.xlsx"
  df = leer_archivo_excel(ruta_archivo)

  # Itera sobre los contratos.

  for _, row in df.iterrows():
    # Descarga los archivos adjuntos.
    descargar_archivos_adjunto(row["Contrtao"], row["Factura"])

if __name__ == "__main__":
  main()