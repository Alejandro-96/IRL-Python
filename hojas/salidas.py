from operator import index
import os
import pandas as pd
from utils.fecha import Fecha
import xlwings as xw
from xlwings import *
from utils.globals import rutaRobot

doc = 'INFORME INDIVIDUAL DE LAS CAPTACIONES (MODIFICADO)'
columnas = ['CodigoContable','NIT','FechaApertura','Plazo','FechaVencimiento','TasaInteresEfectiva','InteresesCausados', 'Saldo']
salidas = ['Salidas de CDAT ', 'Salidas Ahorro Contractual']


def salidaCdatyAC(fecha: Fecha, primeraVez, wb: xw.Book):
    cuentas = [[211005, 211010, 211015,211020], [212505, 212510, 212515, 212520]]
    archivos = os.listdir('{}/Archivos/{}'.format(rutaRobot, doc))
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)

    tabla1 = pd.read_csv(archivo, usecols=columnas, encoding='ANSI', sep=';', skiprows=3)
    tabla1.dtypes
    for salida in salidas:
        print('Diligenciando {}. . .'.format(salida))
        dia = fecha.add_months(1).add_days(-1).dia
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        ws = wb.sheets[salida]
        ws.range('B5').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
        tabla = pd.read_csv(archivo, usecols=columnas, encoding='ANSI', sep=';', skiprows=3)
        tabla = tabla[tabla['CodigoContable'].isin(cuentas[salidas.index(salida)])]
        tabla.drop(columns=['CodigoContable'], inplace=True)
        columnas.remove('CodigoContable')
        tabla = tabla[columnas]
        if not primeraVez:
            ws.range('A9:G' + str(ws.range('A9').end('down').row)).clear()

        ws.range('A9').value = tabla.values
        # Poner en primera posicion de la lista columnas CodigoContable
        columnas.insert(0, 'CodigoContable')

