import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

hoja = 'Creditos Aprobados'
doc = 'CATALOGO DE CUENTAS'

def creditosAprobados(fecha: Fecha, primeraVez:bool, wb: xw.Book):
    print('Diligenciando Creditos Aprobados. . .')
    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)

    ws = wb.sheets[hoja]
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    ws.range('B5').value = '{}/{}/{}'.format(dia, mes, fecha.anio)

    ultimaFila = ws.range('A8').end('down').row

    tabla = pd.read_csv(archivo, skiprows=3, usecols=['CUENTA', 'Saldo'], encoding='ANSI', sep=';')

    # Obtener el saldo de la cuenta 911500
    try:
        saldo = tabla[tabla['CUENTA'] == 911500]['Saldo'].values[0]
    except:
        saldo = 0

    if primeraVez: 
        ws.range('A' + str(ultimaFila)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)
        ws.range('B' + str(ultimaFila)).value = saldo
    else : 
        ws.range('A' + str(ultimaFila + 1 )).value = '{}/{}/{}'.format(dia, mes, fecha.anio)
        ws.range('B' + str(ultimaFila + 1)).value = saldo


    # No se pudo escribir la formula, procedimos a cambiarla por BUSCARV
    # Las 2 formulas que calculan totales de esta hoja estaban mal.


