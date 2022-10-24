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
principal = uic.loadUi("FOLDER_FACTIBILIDADES/principal_window.ui")
creado = uic.loadUi("FOLDER_FACTIBILIDADES/creado.ui")
error = uic.loadUi("FOLDER_FACTIBILIDADES/error.ui")


def gui_principal():
    codigo = principal.codigo.text()
    cliente = principal.cliente.text()
    distrito = principal.distrito.text()
    lat = principal.lat.text()
    lon = principal.lon.text()

    # ------------------------------
    # OBTENEMOS LA LONGITUD
    def coordenda(lon):
        lon = float(lon)


    try:
        coordenda(lon)
    except ValueError:
        lon = lon.replace("°", ";")
        lon = lon.replace("'", ";")
        lon = lon.replace('"', "")
        lon = lon.replace("''", "")
        lon = lon.replace("O", "")
        lon = lon.replace("W", "")
        lon = lon.replace("o", "")
        lon = lon.replace("w", "")
        lon = lon.split(";")
        lon = (float(lon[0]) + float(lon[1])/60 + float(lon[2])/3600)*(-1)
        lon = lon

    # OBTEBNEMOS LA LATITUD
    def coordenda_2(lat):
        lat = float(lat)


    try:
        coordenda_2(lat)
    except ValueError:
        lat = lat.replace("°", ";")
        lat = lat.replace("'", ";")
        lat = lat.replace('"', "")
        lat = lat.replace("''", "")
        lat = lat.replace("S", "")
        lat = lat.replace("s", "")
        lat = lat.split(";")
        lat = (float(lat[0]) + float(lat[1])/60 + float(lat[2])/3600)*(-1)
        lat = lat
    # ------------------------------

    if len(codigo) == 0 or len(cliente) == 0 or len(distrito) == 0:
        principal.alerta.setText("Ingrese la información completa")
    
    elif type(lat) == str or type(lon) != str or type(cliente) != str:
        # Definimos el nombre de nuestros archivos
        client_name = f'{codigo} - Factibilidad {cliente} - {distrito}'

        # Creamos la carpeta del cliente
        os.mkdir(client_name)

        # Creamos una copia de la planilla dentro de la carpeta del cliente
        shutil.copy('PLANILLA.xlsm', f'{client_name}/{client_name}.xlsm')

        # Creamos el kml de ubicación del cliente y la guardamos en la carpeta
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
    creado.show()


def gui_error():
    principal.hide()
    error.show()


def regresar_creado():
    creado.hide()
    principal.alerta.setText("")
    principal.show()
    os.chdir("../")


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