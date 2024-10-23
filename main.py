from common.__init__ import *
from verificar_db import obtener_datos_personal
from segurity_copy import copiar_base_de_datos
from excel import manejar_excel
from carga import mostrar_gif

def cargar_datos_personal():
    """Obtiene los datos personales desde la base de datos."""
    datos = obtener_datos_personal()
    if datos:
        return datos
    else:
        print("No se obtuvieron datos para escribir.")
        return None


def guardar_datos_en_excel(datos_personal, plantilla_excel):
    """Guarda los datos personales en un archivo Excel."""
    puestos_ocupados = calcular_puestos_ocupados(datos_personal)

    try:
        wb = load_workbook(plantilla_excel)
        ws = wb.active

        # Borrar todo el contenido de la hoja
        ws.delete_rows(1, ws.max_row)
        print(f"Se han borrado {ws.max_row} filas anteriores.")

        # Encabezados
        headers = ["Nombre Usuario", "Rol", "Área", "Línea", "Puesto", "Estatus Usuario", "Fecha Registro",
                   "Fecha Maestra", "Fecha de ingreso", "Fecha Baja", "Puestos Ocupados"]
        ws.append(headers)
        print("Encabezados añadidos.")

        # Escribir los datos en las filas
        for fila in datos_personal:
            linea = fila[3]  # Suponiendo que la línea es el cuarto elemento
            puestos_ocupados_count = puestos_ocupados.get(linea, 0)
            ws.append([str(item) for item in fila] + [puestos_ocupados_count])  # Agregar el conteo

        # Guardar los cambios en el archivo
        wb.save(plantilla_excel)
        print("Datos escritos con éxito en el archivo de Excel.")
        time.sleep(2)

    except Exception as e:
        print(f"Error al escribir en el archivo de Excel: {e}")


def calcular_puestos_ocupados(datos_personal):
    """Calcula el número de puestos ocupados basándose en los datos personales."""
    puestos_ocupados = {}
    for fila in datos_personal:
        linea = fila[3]  # línea es el cuarto elemento
        estatus = fila[5]  # Suponiendo que el estatus es el sexto elemento
        if estatus == "Activo":
            puestos_ocupados[linea] = puestos_ocupados.get(linea, 0) + 1
    return puestos_ocupados


def main():
    if 1 == 0:
        mostrar_gif()
    else:
        print('Todo chido')

    plantilla_excel = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\p1.xlsx'
    excel = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Data\Reportes\Reporte de Personal\Plantilla de produccion.xlsx'

    # Ejecutar carga.py en segundo plano
    carga_process = subprocess.Popen(['python', 'carga.py'])
    time.sleep(2)


    # Llamada a la función para obtener los datos
    datos_personal = cargar_datos_personal()

    if datos_personal:
        # Copia de seguridad de la base de datos
        try:
            copiar_base_de_datos()
        except Exception as e:
            print(f"Error al copiar la base de datos: {e}")

        # Guardar datos en el archivo Excel
        guardar_datos_en_excel(datos_personal, plantilla_excel)
        time.sleep(1)

        # Ejecutar el proceso principal
        manejar_excel(excel)

    # Finalizar el proceso de carga
    carga_process.terminate()  # Termina el proceso de carga.py
    carga_process.wait()  # Espera a que el proceso termine
    print('Función carga termino')
    sys.exit(1)
    sys.exit()

if __name__ == "__main__":
    main()