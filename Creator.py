########################################################################
### CREADOR DE CARPETA, PLANILLA Y UBICACIÓN ESPACIAL DE CLIENTE GTD ###
########################################################################

# Importamos las librerias que necesitaremos (istaladas previamente)
from asyncio.windows_events import NULL
from PyQt5 import QtWidgets, uic
import os, shutil, simplekml, time

# Iniciar la aplicación
app = QtWidgets.QApplication([])

# Cargar archivos .ui
principal = uic.loadUi("principal_window.ui")
creado = uic.loadUi("creado.ui")
error = uic.loadUi("error.ui")


def gui_principal():
    codigo = principal.codigo.text()
    cliente = principal.cliente.text()
    distrito = principal.distrito.text()
    lat = principal.lat.text()
    lon = principal.lon.text()

    if len(codigo) == 0 or len(cliente) == 0 or len(distrito) == 0 or len(lat)==0 or len(lon)==0:
        principal.alerta.setText("Ingrese la información completa")
    
    elif type(lat) == str or type(lon) != str or type(cliente) != str:
        # Definimos el nombre de nuestros archivos
        client_name = f'{codigo} - Factibilidad {cliente} - {distrito}'

        # Creamos la carpeta del cliente
        os.mkdir(client_name)

        # Creamos una copia de la planilla dentro de la carpeta del cliente
        shutil.copy('PLANILLA.xlsm', f'{client_name}/{client_name}.xlsm')

        # Creamos el kml de ubicación del cliente y la guardamos en la carpeta
        lat = float(lat)
        lon = float(lon)
        kml = simplekml.Kml()
        # lon, lat, optional height
        kml.newpoint(name=cliente, coords=[(lon, lat)])
        kml.save(f'{client_name}/{client_name}.kmz')

        # Abrimos los archivos creados
        os.chdir(client_name)
        os.startfile(f'{client_name}.xlsm')
        os.startfile(f'{client_name}.kmz')

        crear_otro()

    else:
        gui_error()


def crear_otro():
    principal.hide()
    creado.show().reset()


def gui_error():
    principal.hide()
    error.show()


def regresar_creado():
    creado.hide()
    principal.alerta.setText("")
    principal.show()


def regresar_error():
    error.hide()
    principal.alerta.setText("")
    principal.show()


def salir():
    app.exit()


# Botones
principal.button_crear.clicked.connect(gui_principal)
principal.salir.clicked.connect(salir)

creado.button_crear.clicked.connect(regresar_creado)
creado.salir.clicked.connect(salir)

error.button_crear.clicked.connect(regresar_error)
error.salir.clicked.connect(salir)

# Ejecutable
principal.show()
app.exec()
