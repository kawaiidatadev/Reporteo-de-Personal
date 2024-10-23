from common.__init__ import *
from verificar_db import obtener_datos_personal
from segurity_copy import copiar_base_de_datos
from excel import manejar_excel

running_gif = True  # Variable para controlar la ejecución del GIF



def set_always_on_top():
    # Hacer la ventana siempre en la parte superior (funciona en Windows)
    hwnd = pygame.display.get_wm_info()['window']
    ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def mostrar_gif():
    global running_gif  # Usar la variable global
    try:
        pygame.init()
        gif_paths = [
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif1.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif2.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif3.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif4.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif5.gif'
        ]

        gif_path = random.choice(gif_paths)
        print(f"GIF seleccionado: {gif_path}")

        pil_image = Image.open(gif_path)

        gif_width, gif_height = pil_image.size
        ventana = pygame.display.set_mode((gif_width, gif_height), pygame.NOFRAME | pygame.SRCALPHA)
        pygame.display.set_caption("Cargando...")

        frames = []
        while True:
            try:
                pil_image_data = pil_image.tobytes()
                frame = pygame.image.fromstring(pil_image_data, pil_image.size, pil_image.mode)
                frames.append(frame)
                pil_image.seek(pil_image.tell() + 1)
            except EOFError:
                break

        running = True
        clock = pygame.time.Clock()
        frame_count = len(frames)
        current_frame = 0

        set_always_on_top()

        while running:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False

            set_always_on_top()

            ventana.fill((0, 0, 0, 0))
            ventana.blit(frames[current_frame], (0, 0))
            pygame.display.update()

            current_frame = (current_frame + 1) % frame_count
            clock.tick(10)  # Ajuste para un GIF más fluido

            if not running_gif:  # Salir del bucle si la variable se establece en False
                running = False

        pygame.quit()
    except Exception as e:
        print(f"Error en la ventana principal: {e}")
        pygame.quit()
        sys.exit()

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
    global running_gif  # Acceso a la variable global

    plantilla_excel = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\p1.xlsx'
    excel = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Data\Reportes\Reporte de Personal\Plantilla de produccion.xlsx'

    # Ejecutar carga.py en segundo plano
    carga_thread = threading.Thread(target=mostrar_gif)
    carga_thread.start()
    time.sleep(2)  # Espera a que la carga inicie correctamente


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
    running_gif = False  # Termina el bucle del GIF
    carga_thread.join()  # Espera a que el hilo de carga termine antes de salir
    print('Función carga terminó')
    sys.exit(1)
    sys.exit()

if __name__ == "__main__":
    main()