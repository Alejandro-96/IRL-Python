from utils.globals import *
import xlwings as xw
import pandas as pd
from utils.fecha import Fecha


doc = 'INFORME DEUDORES PATRONALES Y EMPRESAS'
hoja = 'Recaudo CXC'
cuentas = [213005, 213010, 213015, 213020]

def cxc(fecha: Fecha, primeraVez: bool , wb: xw.Book):
    ws = wb.sheets[hoja]
    dia = fecha.add_months(1).add_days(-1).dia
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    ws.range('B6').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    archivos = os.listdir('{}/Archivos/{}'.format(rutaRobot, doc))
    print('Diligenciando Recaudo CXC')
    if primeraVez:
        fechaAux = fecha.add_months(-12)
        for i in range(13):
            dia = fechaAux.add_months(1).add_days(-1).dia
            mes = '0' + str(fechaAux.mes) if fechaAux.mes < 10 else str(fechaAux.mes)
            ws.range('A' + str(i + 11)).value = '{}/{}/{}'.format(dia, mes, fechaAux.anio)
            archivo = [archivo for archivo in archivos if fechaAux.as_Text() in archivo][0]
            archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)
   
            tabla = pd.read_csv(archivo, usecols=['SaldoTotal', 'Número Meses de Incumplimiento'], encoding='ANSI', sep=';', skiprows=3)
     
            tabla = tabla[tabla['Número Meses de Incumplimiento'] == 0]
            suma = tabla['SaldoTotal'].sum()

            ws.range('B' + str(i + 11)).value = suma
            fechaAux = fechaAux.add_months(1)
    else:
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        dia = fecha.add_months(1).add_days(-1).dia
        ultimaFila = ws.range('A11').end('down').row
        ws.range('A' + str(ultimaFila + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)
        archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
        archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot, doc), archivo)
 
        tabla = pd.read_csv(archivo, usecols=['SaldoTotal', 'Número Meses de Incumplimiento'], encoding='ANSI', sep=';', skiprows=3)
     
        tabla = tabla[tabla['Número Meses de Incumplimiento'] == 0]
        suma = tabla['SaldoTotal'].sum()

        ws.range('B' + str(ultimaFila + 1)).value = suma