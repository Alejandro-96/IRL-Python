from re import A
import pandas as pd
import xlwings as xw
from xlwings import *
from utils.fecha import *
import os
from utils.globals import rutaRobot

doc = 'INFORME INDIVIDUAL DE APORTES O CONTRIBUCIONES'
hoja = 'salida de Aportes'
def salidasp(fecha: Fecha, primeraVez: bool,wb: xw.Book):
    ws = wb.sheets[hoja]
    mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
    dia = fecha.add_months(1).add_days(-1).dia
    ws.range('B5').value = '{}/{}/{}'.format(dia, mes, fecha.anio)
    archivos = os.listdir('{}/Archivos/{}'.format(rutaRobot,doc))
    print('Diligenciando Salidas de aportes. . .')
    
   
    if primeraVez:
            fechaAux = fecha.add_months(-12)
            for i in range(13):
                dia = fechaAux.add_months(1).add_days(-1).dia
                mes = '0' + str(fechaAux.mes) if fechaAux.mes < 10 else str(fechaAux.mes)
                #ws.range('B' + str(i + 14)).value = '{}/{}/{}'.format(dia, mes, fechaAux.anio)
                archivo = [archivo for archivo in archivos if fechaAux.as_Text() in archivo][0]
                archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot,doc), archivo)
                
            
                tabla = pd.read_csv(archivo, usecols=['Número de identificación','Saldo a fecha'], encoding='ANSI', sep=';', skiprows=3)
                tablaAux = tabla.drop(columns='Saldo a fecha')
            
            
                fechasmespasado = fechaAux.add_months(-1)
                archivo = [archivo for archivo in archivos if fechasmespasado.as_Text() in archivo][0]
                archivo = os.path.join('{}/Archivos/{}'.format(rutaRobot,doc), archivo)
                tablaP = pd.read_csv(archivo, usecols=['Número de identificación','Saldo a fecha'], encoding='ANSI', sep=';', skiprows=3)
                tablaAuxP = tablaP.drop(columns='Saldo a fecha')


                resultado1 = tablaAuxP[~tablaAuxP.apply(tuple,1).isin(tablaAux.apply(tuple,1))]
                resultado2 = resultado1.drop_duplicates(subset='Número de identificación')

                tablaAuxP = tablaP
                #Filtrar en tablaAuxP los NIT que estan en resultado2
                tablaAuxP = tablaAuxP[tablaAuxP['Número de identificación'].isin(resultado2['Número de identificación'])]
            
                saldo = tablaAuxP.sum()['Saldo a fecha']
            

                ws.range('A' + str(i + 14)).value = '{}/{}/{}'.format(dia, mes, fechaAux.anio)  
                ws.range('B' + str(i + 14)).value = resultado2.size
                ws.range('C' + str(i + 14)).value = saldo  

                fechaAux = fechaAux.add_months(1)
            

    else: 
        mes = '0' + str(fecha.mes) if fecha.mes < 10 else str(fecha.mes)
        dia = fecha.add_months(1).add_days(-1).dia
    
        ultimaFilaFecha = ws.range('A' + str(ws.api.UsedRange.Rows.Count+13)).end('up').row
        ultimaFilaSaldo = ws.range('C' + str(ws.api.UsedRange.Rows.Count+13)).end('up').row
        ultimaFilaAso   = ws.range('B' + str(ws.api.UsedRange.Rows.Count+13)).end('up').row
        ws.range('A' + str(ultimaFilaFecha + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)

        archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
        archivo = [archivo for archivo in archivos if fecha.as_Text() in archivo][0]
        archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
        
        
        tablaActual = pd.read_csv(archivo, usecols=['Número de identificación','Saldo a fecha'], encoding='ANSI', sep=';', skiprows=3)
        MesActual= tablaActual.drop(columns='Saldo a fecha')
        MesActual['Número de identificación'] = MesActual['Número de identificación'].astype('str')

        fechasmespasado = fecha.add_months(-1)
      
        archivos = os.listdir(rutaRobot + '/Archivos/' + doc)
        archivo = [archivo for archivo in archivos if fechasmespasado.as_Text() in archivo][0]
        archivo = os.path.join(rutaRobot + '/Archivos/' + doc, archivo)
        tablaAnterior = pd.read_csv(archivo, usecols=['Número de identificación','Saldo a fecha'], encoding='ANSI', sep=';', skiprows=3)
        MesAnterior = tablaAnterior.drop(columns='Saldo a fecha')
        MesAnterior['Número de identificación'] = MesAnterior['Número de identificación'].astype('str')


        resultado1 = MesAnterior[~MesAnterior.apply(tuple,1).isin(MesActual.apply(tuple,1))]
        resultado2 = resultado1.drop_duplicates(subset='Número de identificación')


        tablaAnterior['Número de identificación'] = tablaAnterior['Número de identificación'].astype('str')
        #Filtrar en tablaAuxP los NIT que estan en resultado2
        tablaCombinacion = tablaAnterior[tablaAnterior['Número de identificación'].isin(resultado2['Número de identificación'])]
    
        saldo = tablaCombinacion.sum()['Saldo a fecha']
        print(resultado2.size,saldo)
        

        ws.range('A' + str(ultimaFilaFecha + 1)).value = '{}/{}/{}'.format(dia, mes, fecha.anio)  
        ws.range('B' + str(ultimaFilaAso + 1)).value = resultado2.size
        ws.range('C' + str(ultimaFilaSaldo + 1)).value = saldo
