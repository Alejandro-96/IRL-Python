from dataclasses import replace
from datetime import datetime
from email import message
import os
import re
from tkinter import messagebox
import pandas as pd

from utils.fecha import Fecha
from utils.globals import *

url = 'https://www.supersolidaria.gov.co/sites/default/files/public/data/desviacion_estandar_'

def obtenerDesviacionEstandar(fecha: datetime):
    archivos = os.listdir(rutaRobot)
    for i in archivos:
        if i.startswith('Desviacion-estandar'):
            archivo = i
            break

    tabla = pd.read_excel(rutaRobot + '/' + archivo, sheet_name='DESVIACION ESTANDAR',
                          skiprows=5, usecols='B:F')

    fechaTexto = Fecha(1, fecha.month - 1, fecha.year).as_Text().replace(' ', ' DE ')

    # Obtener la fila cuyo texto en la columna 1 contenga la fecha ignorando mayusculas
    fila = tabla[tabla['PERIODO CORTE'].str.contains(fechaTexto, case=False)]

    if not fila.empty:
        return fila.iloc[0][-1]
    else:
        messagebox.showerror('Error', 'No se encontro la desviacion estandar de ' + fechaTexto)



def desviacionEstandar(fecha: datetime):
    print('Validando desviación estándar. . .')
    fechaActual = Fecha(1, datetime.now().month, datetime.now().year)
    # Obtener archivos de una carpeta
    archivos = os.listdir(rutaRobot)
    # Determinar si existe un archivo cuyo nombre empieza por Desviacion-estandar
    existeArchivo = False
    archivo = ''
    for i in archivos:
        if i.startswith('Desviacion-estandar'):
            archivo = i
            existeArchivo = True
            break

    # Si no existe el archivo
    if not existeArchivo:
        fechaArchivo = 'ENERO 1570'
    else:
        # Tomar el texto despues de Desviacion-estandar y antes de .xlsx
        fechaArchivo = archivo[len('Desviacion-estandar '):archivo.find('.xlsx')]
    
    fechaArchivoDate = Fecha(1, meses.index(fechaArchivo[:fechaArchivo.find(' ')].upper()) + 1, int(fechaArchivo[fechaArchivo.find(' ') + 1:]))

    # Si la fecha del archivo es menor a la fecha actual, descargar el archivo
    if fechaArchivoDate.as_datetime() < fecha:
        print('Descargando desviación estándar. . .')
        if existeArchivo:
            os.remove(rutaRobot + '/' + archivo)
        found = False
        while not found:
            try:
                if os.path.exists(rutaRobot + '/Desviacion-estandar ' + fechaActual.as_Text().lower() + '.xlsx'):
                    found = True
                else:                
                    desviacion = pd.read_excel(url + fechaActual.as_Text().lower().replace(' ', '_') + '.xlsx')
                    found = True
            except:
                pass
            fechaActual = fechaActual.add_months(-1)
    return obtenerDesviacionEstandar(fecha)