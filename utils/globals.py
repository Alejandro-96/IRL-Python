import os
from PyQt5.QtWidgets import QMessageBox
import sys
from utils.fecha import *
desviacion = ""

fechaActual = Fecha(1, datetime.now().month, datetime.now().year)
# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')

meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def descargarArchivo(url, nombre):
    # Descargar archivo ignorando certificados SSL
    import ssl
    import urllib.request
    import urllib.error
    import urllib.parse
    #Descargar 
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        urllib.request.urlretrieve(url, nombre)
        print(url)
        return True
    except urllib.error.HTTPError as e:
        print('Error: ' + str(e.code))
        QMessageBox.critical(None, 'Error', 'Error: ' + str(e.code))
        return False
    
