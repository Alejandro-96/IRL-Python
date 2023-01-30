from __future__ import barry_as_FLUFL
from utils.convert2csv import convertiracsv
from utils.fecha import *
import os
from tkinter import *
from tkinter import messagebox
import pandas as pd
from utils.convert2csv import convertiracsv
from utils.globals import *

carpetas = [
    'CATALOGO DE CUENTAS',
    'CREDITOS DE BANCOS Y OTRAS OBLIGACIONES FINANCIERAS (NUEVO)',
    'INFORME DEUDORES PATRONALES Y EMPRESAS',
    'INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES',
    'INFORME INDIVIDUAL DE CARTERA DE CREDITO (MODIFICADO)',
    'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)',
    'RELACION DE INVERSIONES',
    'SALDOS DIARIOS DE AHORRO'
]

archivosPrimeraVez = [13, 1, 13, 25, 13, 25, 1, 1]

archivosSegundaVez = [1 , 1, 1 , 2 , 1 , 2, 1, 1]

def crearCarpetas():
    # Ignorar errores de certificado
    #urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    os.mkdir(rutaRobot)
    os.mkdir(rutaRobot + '/' + 'Archivos')
    for i in carpetas:
        os.mkdir(rutaRobot + '/' + 'Archivos' + '/' + i)
    os.mkdir(rutaRobot + '/' + 'PlanosDiligenciados')
    os.mkdir(rutaRobot + '/' + 'ArchivosNuevos')
    found = False
    #url = 'https://www.supersolidaria.gov.co/sites/default/files/public/data/desviacion_estandar_'
    fechaActual = Fecha(1, datetime.now().month, datetime.now().year)
    fechaActual = fechaActual.add_months(-2)
    """print("Buscando la desviación estándar más reciente. . .")
    while not found:
        try:
            if os.path.exists(rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx'):
                found = True
            else:
                resp = requests.get(url + fechaActual.as_Text().lower().replace(' ', '_') + '.xlsx', verify=False, stream=True, timeout=1)
                if resp.status_code == 200:
                    print("Encontrada la desviación estándar " + fechaActual.as_Text().lower() + ". . .")
                    found = True
                    print("Conectando con SuperSolidaria. . .")
                    print("Descargando la desviación estándar " + fechaActual.as_Text().lower() + ". . .")
                    descargarArchivo(url + fechaActual.as_Text().lower().replace(' ', '_') + '.xlsx', rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx')
                
        except Exception as e:
            #Mostrar el error
            print(e)
        fechaActual = fechaActual.add_months(-1)"""

def validarCarpetas():
    print('Validando carpetas. . .')
    # Si no existe la carpeta del robot, crearla
    if not os.path.exists(rutaRobot):
        print('Creando carpeta y archivos del robot. . .')
        crearCarpetas()

        # Abrir el explorador de archivos en la carpeta del robot
        os.system('explorer ' + rutaRobot + '\ArchivosNuevos')
 # MessageBox
        messagebox.showinfo(
            'RobotIRL', 'Parece que no hay archivos para la ejecución, he abierto la carpeta por ti, deja los archivos aqui y vuelve a ejecutar el robot.' )
       
        os.system('python ui.py')
        exit()
    else:
        clasificarArchivos()


def clasificarArchivos():
    # Obtener archivos de la carpeta ArchivosNuevos
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')

    if len(archivos) != 0:
        ajustarNombres()
        convertiracsv()
    
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')

    print('Clasificando los archivos. . .')
    # Clasificar los archivos
    for carpeta in carpetas:
        for archivo in archivos:
            if archivo.startswith(carpeta):
                for mes in meses:
                    if mes in archivo.upper():
                        if os.path.exists(rutaRobot + '/Archivos/' + carpeta + '/' + archivo):
                            os.remove(rutaRobot + '/Archivos/' + carpeta + '/' + archivo)
                            
                        os.rename(rutaRobot + '/ArchivosNuevos/' + archivo,
                                rutaRobot + '/Archivos/' + carpeta + '/' + archivo)
                        break

    # Si quedaron archivos en la carpeta ArchivosNuevos
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')
    if len(archivos) != 0:
        # MessageBox
        messagebox.showinfo(
            'RobotIRL', 'Ups, algunos archivos no pudieron clasificarse, voy a mostrartelos, verifica sus nombres y vuelve a ejecutar el robot.')
        # Abrir el explorador de archivos en la carpeta ArchivosNuevos
        os.system('explorer ' + rutaRobot + '\ArchivosNuevos')
        # Ejecutar ui.py
        os.system('python ui.py')
        exit()


def validarArchivos(primeraVez, fecha: Fecha):
    auxFecha = fecha
    print('Validando archivos para diligenciamiento de ' + str(fecha.as_Text()) + '. . .')
    for carpeta in carpetas:
        archivos = [os.path.splitext(filename)[0] for filename in os.listdir(rutaRobot + '/Archivos/' + carpeta)]
        repeticiones = archivosPrimeraVez[carpetas.index(carpeta)] if primeraVez else archivosSegundaVez[carpetas.index(carpeta)]
        repeticiones -= 1
        auxFecha = fecha.add_months(-repeticiones)
        repeticiones = 1 if repeticiones < 1 else repeticiones
        for j in range(repeticiones):
            nombreArchivo = carpeta + ' ' + auxFecha.as_Text()
            print('Validando {}. . .'.format(nombreArchivo))
            if nombreArchivo not in archivos:
                messagebox.showinfo("RobotIRL", "No pude encontrar  el archivo " + nombreArchivo +
                                    " te abriré la carpeta de archivos para que lo muevas o lo busques y corrijas.")
                os.system('explorer ' + rutaRobot)
                os.system('python ui.py')
                exit()
            else:
                auxFecha = auxFecha.add_months(1)

def ajustarNombres():
    print('Ajustando nombres de archivos nuevos. . .')
    for archivo in os.listdir(rutaRobot + '/ArchivosNuevos'):
        # Renombrar cambiando _ por espacio
        os.rename(rutaRobot + '/ArchivosNuevos/' + archivo,
                    rutaRobot + '/ArchivosNuevos/' + archivo.replace('_', ' '))

        # Cambiar el texto "CSV" por "csv"
        if archivo.endswith('CSV'):
            os.rename(rutaRobot + '/ArchivosNuevos/' + archivo,
                        rutaRobot + '/ArchivosNuevos/' + archivo.replace('CSV', 'csv'))


def buscarPlano(mes, anio):
    if mes == 0:
        print('Buscando plano de {} {}. . .'.format(meses[mes - 1], anio-1))
        if os.path.exists(rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsx'.format(meses[mes - 1], anio-1)):
            print('Encontrado plano de {} {}. . .'.format(meses[mes - 1], anio-1))
            return rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsx'.format(meses[mes - 1], anio-1)
        else:
            QMessageBox.warning(None, "RobotIRL", "No se encontró el plano de {} {}, verifique que este en la carpeta PlanosDiligenciados del robot y que se llame ""PLANOIRL MES AÑO.xlsx"", recuerde si no es la primera vez que se ejecuta el robot, necesita el plano del mes anterior".format(meses[mes - 1], anio-1))
            os.system('python ui.py')
            exit()
    else:
        print('Buscando plano de {} {}. . .'.format(meses[mes - 1], anio))
        if os.path.exists(rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsx'.format(meses[mes - 1], anio)):
            print('Encontrado plano de {} {}. . .'.format(meses[mes - 1], anio))
            return rutaRobot + '/PlanosDiligenciados/PLANOIRL {} {}.xlsx'.format(meses[mes - 1], anio)
        else:
            QMessageBox.warning(None, "RobotIRL", "No se encontró el plano de {} {}, verifique que este en la carpeta PlanosDiligenciados del robot y que se llame ""PLANOIRL MES AÑO.xlsx"", recuerde si no es la primera vez que se ejecuta el robot, necesita el plano del mes anterior".format(meses[mes - 1], anio))
            os.system('python ui.py')
            exit()