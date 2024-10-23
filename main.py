from common.__init__ import *
from verificar_db import obtener_datos_personal
from segurity_copy import copiar_base_de_datos

plantilla_excel = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\p1.xlsx'

# Llamada a la función para obtener los datos
datos_personal = obtener_datos_personal()

if datos_personal:
    # Copia de seguridad de la db
    try:
        copiar_base_de_datos()
    except Exception as e:
        print(f"Error al copiar la base de datos: {e}")

    # Calcular puestos ocupados
    puestos_ocupados = {}
    for fila in datos_personal:
        linea = fila[3]  # línea es el cuarto elemento
        estatus = fila[5]  # Suponiendo que el estatus es el sexto elemento
        if estatus == "Activo":
            if linea in puestos_ocupados:
                puestos_ocupados[linea] += 1
            else:
                puestos_ocupados[linea] = 1

    try:
        # Cargar el archivo de Excel
        wb = load_workbook(plantilla_excel)
        ws = wb.active  # Usar la hoja activa o selecciona otra si lo necesitas

        # Borrar todo el contenido de la hoja
        ws.delete_rows(1, ws.max_row)
        print(f"Se han borrado {ws.max_row} filas anteriores.")

        # Encabezados
        headers = ["Nombre Usuario", "Rol", "Área", "Línea", "Puesto", "Estatus Usuario", "Fecha Registro",
                   "Fecha Maestra", "Fecha de ingreso", "Fecha Baja", "Puestos Ocupados"]

        # Escribir los encabezados
        ws.append(headers)
        print("Encabezados añadidos.")

        # Escribir los datos en las filas, incluyendo "Puestos Ocupados"
        for fila in datos_personal:
            linea = fila[3]  # Suponiendo que la línea es el cuarto elemento
            puestos_ocupados_count = puestos_ocupados.get(linea, 0)
            ws.append([str(item) for item in fila] + [puestos_ocupados_count])  # Agregar el conteo

        # Guardar los cambios en el archivo
        wb.save(plantilla_excel)
        print("Datos escritos con éxito en el archivo de Excel.")
    except Exception as e:
        print(f"Error al escribir en el archivo de Excel: {e}")
else:
    print("No se obtuvieron datos para escribir.")