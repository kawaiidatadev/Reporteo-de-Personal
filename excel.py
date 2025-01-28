from common.__init__ import *

"""
Menajar el excel en este codigo:
1. Copíar el excel
2. pegar el excel el la ruta de descargas del usuario de windows actual guardando el excel con un nombre nuevo
3. nombre nuevo = "Reporte Personal _ {DD-MM-AAAA_HH_MM_SS}" Segun la hora y fecha de mexico guadalajara.
4. abrir internamente el excel (si es necesario) para actualizar la conexiones y todo el libro.
5. esperar a que la actualizaciones terminen, y ahora abrir la ruta donde se descargo el excel en una ventana
maximizada.
"""


# Función para obtener la ruta de descargas del usuario actual
def obtener_ruta_descargas():
    try:
        user_downloads = os.path.join(os.path.expanduser("~"), 'Downloads')
        return user_downloads
    except Exception as e:
        print(f"Error al obtener la ruta de descargas: {e}")
        return None

# Función para formatear la fecha y hora actual según la zona horaria de Guadalajara, México
def obtener_timestamp():
    try:
        timestamp = time.strftime('%d-%m-%Y_%H_%M_%S', time.localtime())
        return timestamp
    except Exception as e:
        print(f"Error al obtener la fecha y hora actual: {e}")
        return None

# Función para copiar el archivo Excel a la ruta de descargas con un nuevo nombre
def copiar_excel(origen, destino):
    try:
        shutil.copy2(origen, destino)
        print(f"Archivo copiado a {destino}")
    except FileNotFoundError:
        print("Archivo de origen no encontrado.")
    except Exception as e:
        print(f"Error al copiar el archivo: {e}")

# Ruta al archivo .bat
ruta_bat = r"\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\excel_error\ejecutable.bat"

def actualizar_conexiones_excel(ruta_excel):
    excel_app = None
    try:
        # Crear la instancia de la aplicación de Excel
        excel_app = win32.Dispatch("Excel.Application")

        # Mantener Excel completamente oculto y desactivar alertas
        excel_app.Visible = False  # Mantener Excel oculto
        excel_app.ScreenUpdating = False  # Evitar que Excel se actualice visualmente
        excel_app.DisplayAlerts = False  # Desactivar alertas

        # Abrir el archivo copiado
        wb = excel_app.Workbooks.Open(ruta_excel)

        # Actualizar cada conexión individualmente 3 veces
        for connection in wb.Connections:
            for i in range(3):  # Repetir 3 veces
                try:
                    print(f"Actualizando conexión '{connection.Name}' (Intento {i + 1})")
                    connection.Refresh()
                    time.sleep(1)  # Pausa para evitar conflictos entre actualizaciones
                except Exception as e:
                    print(f"Error al actualizar la conexión '{connection.Name}' en intento {i + 1}: {e}")

        # Guardar los cambios
        wb.Save()
        print(f"Conexiones actualizadas y archivo guardado en {ruta_excel}")

        # Cerrar el archivo
        wb.Close()

    except Exception as e:
        print(f"Error al actualizar las conexiones de Excel: {e}")
        ejecutar_bat_y_reintentar(ruta_excel)
    finally:
        if excel_app:
            excel_app.Quit()  # Asegurarse de que Excel se cierre correctamente

def ejecutar_bat_y_reintentar(ruta_excel):
    try:
        # Ejecutar el archivo .bat
        print(f"Ejecutando archivo .bat: {ruta_bat}")
        subprocess.run(ruta_bat, check=True, shell=True)
        print("Archivo .bat ejecutado correctamente. Reintentando...")

        # Reintentar la función
        actualizar_conexiones_excel(ruta_excel)
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el archivo .bat: {e}")
    except Exception as e:
        print(f"Error inesperado al ejecutar el archivo .bat: {e}")


# Función para abrir la carpeta de descargas en una ventana maximizada
def abrir_carpeta_descargas(ruta_descargas):
    try:
        os.startfile(ruta_descargas)
        print(f"Carpeta de descargas abierta: {ruta_descargas}")
    except Exception as e:
        print(f"Error al abrir la carpeta de descargas: {e}")

# Función principal para manejar todo el proceso
def manejar_excel(origen_excel):
    try:
        # Obtener la ruta de descargas
        ruta_descargas = obtener_ruta_descargas()
        if not ruta_descargas:
            return

        # Obtener timestamp para el nombre del archivo
        timestamp = obtener_timestamp()
        if not timestamp:
            return

        # Crear el nuevo nombre del archivo
        nuevo_nombre_excel = f"Reporte Personal _ {timestamp}.xlsx"
        nueva_ruta_excel = os.path.join(ruta_descargas, nuevo_nombre_excel)

        # 1. Copiar el archivo de Excel
        copiar_excel(origen_excel, nueva_ruta_excel)

        # 2. Actualizar el archivo Excel y conexiones
        actualizar_conexiones_excel(nueva_ruta_excel)

        # 3. Abrir la carpeta de descargas del usuario
        abrir_carpeta_descargas(ruta_descargas)

    except Exception as e:
        print(f"Error en el proceso general: {e}")


