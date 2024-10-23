"""
Conectar a la base de datos con try expet, mostrando msgbox al usuario, y si la conexion es exitosa retornar
en una funcion  los resultados de la siguiente consulta:

consulta = SELECT nombre_usuario, rol, area, linea, puesto, estatus_usuario,
            fecha_registro, fecha_maestra FROM personal_esd;
"""

from common import *

db_patch = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\database.sqlite'


def mostrar_mensaje(mensaje, titulo):
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, 1)


def obtener_datos_personal():
    consulta = """SELECT nombre_usuario, rol, area, linea, puesto, estatus_usuario, 
                  fecha_registro, fecha_maestra, fecha_ingreso, fecha_baja FROM personal_esd;"""
    try:
        conexion = sqlite3.connect(db_patch)
        cursor = conexion.cursor()
        cursor.execute(consulta)
        resultados = cursor.fetchall()

        if resultados:
            print("Conexión exitosa a la base de datos.", "Éxito")

        conexion.close()
        return resultados

    except sqlite3.Error as error:
        mostrar_mensaje(f"Error al conectar a la base de datos: {error}", "Error")
        return None