import __init__


db_patch = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\database.sqlite'

"""
Conectar a la base de datos con try expet, mostrando msgbox al usuario, y si la conexion es exitosa retornar 
en una funcion  los resultados de la siguiente consulta:

consulta = SELECT nombre_usuario, rol, area, linea, puesto, estatus_usuario, 
            fecha_registro, fecha_maestra FROM personal_esd;
"""


def mostrar_mensaje(mensaje, titulo):
    # Utiliza la función MessageBox de Windows para mostrar el mensaje al usuario
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, 1)


def obtener_datos_personal():
    consulta = """SELECT nombre_usuario, rol, area, linea, puesto, estatus_usuario, 
                  fecha_registro, fecha_maestra FROM personal_esd;"""
    try:
        # Conectarse a la base de datos SQLite
        conexion = sqlite3.connect(db_patch)
        cursor = conexion.cursor()

        # Ejecutar la consulta
        cursor.execute(consulta)
        resultados = cursor.fetchall()

        # Mostrar mensaje de éxito
        mostrar_mensaje("Conexión exitosa a la base de datos.", "Éxito")

        # Cerrar la conexión
        conexion.close()

        return resultados

    except sqlite3.Error as error:
        # Mostrar mensaje de error
        mostrar_mensaje(f"Error al conectar a la base de datos: {error}", "Error")


# Llamada a la función
datos = obtener_datos_personal()
if datos:
    for fila in datos:
        print(fila)