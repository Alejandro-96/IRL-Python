from ctypes import wstring_at
import os
import pandas as pd
from utils.fecha import Fecha
import xlwings as xw
from xlwings import *
from utils.globals import rutaRobot

carpeta = 'CATALOGO DE CUENTAS'

codigos = [
    110500, 111000, 111500, 120400, 121300, 123016, 
    112001, 112003, 112005, 112006, 112007, 112008,
    120305, 120310, 120315, 120320, 120325, 120330
]

columnas = [
    'A', 'B', 'C', 'E', 'F', 'G', 
    'I', 'J', 'K', 'L', 'M', 'N', 
    'O', 'P', 'Q', 'R', 'S', 'T'
]

def activosLiquidos(fecha: Fecha, primeraVez, wb: xw.Book):
    print('Diligenciando Activos Liquidos. . .')
    archivos = os.listdir(rutaRobot + '/Archivos/' + carpeta)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + carpeta, archivo)

    ws = wb.sheets['ActivosLiquidos']
 
    tabla = pd.read_csv(archivo, usecols=['CUENTA', 'Saldo'], skiprows=3, encoding='ANSI', sep=';')
   

    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    ws.range('B5').value = '{}/{}/{}'.format(fecha.add_months(1).add_days(-1).dia, mes, fecha.anio)

    for col in columnas:
        saldo = tabla.loc[tabla['CUENTA'] == codigos[columnas.index(col)]]['Saldo']

        if saldo.empty:
            ws.range(col + '14').value = 0
        else:
            ws.range(col + '14').value = saldo.values[0]

        fecha = fecha.add_months(1)
        fecha = fecha.add_days(-1)
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
