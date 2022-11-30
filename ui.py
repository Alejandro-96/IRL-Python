from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtWidgets import *
import datetime
import sys
from utils.globals import *
from utils.validaciones import crearCarpetas
from main import *


class Ui(QtWidgets.QDialog):

    def __init__(self):

        path = resource_path('interfaz.ui')
        super(Ui, self).__init__() # Call the inherited classes __init__ method
        uic.loadUi(path, self) # Load the .ui file

        btnStart = self.findChild(QPushButton, 'btnStart')
        btnStart.clicked.connect(self.iniciar)

        btnFolder = self.findChild(QPushButton, 'btnFolder')
        btnFolder.clicked.connect(self.folder)

        self.inputMes = self.findChild(QComboBox, 'inputMes')
        # Poner lista de meses en el combobox
        self.inputMes.addItem('Seleccione')
        self.inputMes.addItems(meses)
        self.inputAnio = self.findChild(QSpinBox, 'inputAnio')
        self.primerraVez = self.findChild(QCheckBox, 'chkFirstTime')
        self.lblManual = self.findChild(QLabel, 'lblManual')
        self.lblError = self.findChild(QLabel, 'lblError')
        self.lblError.setText('')

        # Quitar el boton ? del formulario
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)

        self.lblManual.mousePressEvent = self.openManual

        self.show()

    def iniciar(self):
        print('Iniciando...')
        mes = str(self.inputMes.currentText())
        anio = self.inputAnio.value()

        if mes == 'Seleccione':
            self.lblError.setText('Seleccione un mes')
            self.lblError.setStyleSheet("color: red; font-weight: bold;")
            self.lblError.setAlignment(QtCore.Qt.AlignCenter)
        elif datetime(anio, meses.index(mes) + 1, 1) > datetime.now():
            self.lblError.setText('Año mayor al actual')
            self.lblError.setStyleSheet("color: red; font-weight: bold;")
            self.lblError.setAlignment(QtCore.Qt.AlignCenter)
        else:
            self.lblError.setText('')
            print('Mes: ' + mes + ' Año: ' + str(anio) + ' Primera vez: ' + str(self.primerraVez.isChecked()))
            self.hide()
            main(mes, anio, self.primerraVez.isChecked())
             # MessageBox
            messagebox.showinfo(
                'RobotIRL', 'Listo!' )
            self.show()

        


    def closeEvent(self, event):
        # This is executed when the window is closed
        print('Cerrando ventana')
        event.accept()

    def openManual(self, event):
        print('Abrir manual')
        os.system('explorer ' + ManualUsuario) 


    def folder(self):
        print('Abrir carpeta')
        if not os.path.exists(rutaRobot):
            crearCarpetas()
            self.print_message('He creado las carpetas necesarias para la ejecución del programa, la abriré ahora para ti, deja tus archivos en la carpeta "ArchivosNuevos" y yo me encargaré del resto.')
            
        os.system('explorer ' + rutaRobot)

    def print_message(self, message):
        QMessageBox.information(self, 'RobotIRL', message)

app = QtWidgets.QApplication(sys.argv) # Create an instance of QtWidgets.QApplication
window = Ui() 
app.exec_() # Execute the app