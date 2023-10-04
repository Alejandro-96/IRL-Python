import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot
doc = 'SALDOS DIARIOS DE AHORRO'


def salidasao(fecha: Fecha, wb: xw.Book):
    print('Diligenciando Salidas de ahorro ordinario. . .')
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    ultimodia = fecha.add_months(1).add_days(-1).dia

    archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
    archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
    archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
    
    ws = wb.sheets['Salidas De Ahorro ordinario']
    tabla = pd.read_csv(archivo, usecols=['Saldo'], encoding='ANSI', sep=';', skiprows=3)

    ws.range('B5').value = '{}/{}/{}'.format(ultimodia,mes,fecha.anio)

    #borrar datos anteriores
    ws.range('A9:B' + str(ultimodia + 8)).value = ''

        

    for i in range(1, ultimodia + 1):
        dia = '0' + str(i) if i < 10 else str(i)
        saldo = tabla.get('Saldo')[i-1]
        ws.range('B' + str(i + 8)).value = saldo  
        ws.range('A' + str(i + 8)).value = '{}/{}/{}'.format(dia,mes,fecha.anio)
    

    #print(ultimodia, ultimodia + 1 )


        
 









