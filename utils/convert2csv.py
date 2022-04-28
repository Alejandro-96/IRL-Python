from tkinter import messagebox
import pandas as pd
import os
import openpyxl

# Obtener ruta de la carpeta documentos
rutaDocumentos = os.path.join(os.path.expanduser('~'), 'Documents')
rutaRobot = os.path.join(rutaDocumentos, 'RobotIRL')

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
def convertiracsv():
    archivos = os.listdir(rutaRobot + '/ArchivosNuevos')
    for archivo in archivos:
        if archivo.lower().endswith('.xlsx') or archivo.lower().endswith('.xls'):
            df = pd.read_excel(rutaRobot + '/ArchivosNuevos/' + archivo, sheet_name='SIAC', skiprows=3, engine='openpyxl')
            os.remove(rutaRobot + '/ArchivosNuevos/' + archivo)
            nombre = archivo.split('.')[0]
            df.to_csv(rutaRobot + '/ArchivosNuevos/' + nombre + '.csv', index = None, header=True)
        else:
            nombreArchivo = archivo.replace('CSV', 'csv')
            # Renombrar archivo
            os.rename(rutaRobot + '/ArchivosNuevos/' + archivo, rutaRobot + '/ArchivosNuevos/' + nombreArchivo)