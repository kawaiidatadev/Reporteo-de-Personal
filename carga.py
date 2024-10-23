from common.__init__ import *

"""
Este es el codigo que se debe ejecutar primero que nada, en un proceso separado de main, en segundo plano,
utilizando recursos diferentes y ejecutandose en paralelo mientras que se termina de ejecutar main.py.
Este codigo debe de abrir en el centro de la pantalla de cualquier usuario, el gif con movimiento
y ponerlo siempre por encima de todas las ventanas de windows, y finalizar despues de que main.py finalize.
"""



def set_always_on_top():
    # Hacer la ventana siempre en la parte superior (funciona en Windows)
    hwnd = pygame.display.get_wm_info()['window']
    ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def mostrar_gif():
    try:
        pygame.init()
        # Lista de rutas de los GIFs
        gif_paths = [
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif1.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif2.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif3.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif4.gif',
            r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\plantilla_personal\gif5.gif'
        ]

        # Elegir aleatoriamente un gif_path
        gif_path = random.choice(gif_paths)

        # Imprimir el gif_path elegido
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

        # Crear una superficie para dibujar
        buffer_surface = pygame.Surface((gif_width, gif_height), pygame.SRCALPHA)

        # Establecer la ventana como siempre en la parte superior inicialmente
        set_always_on_top()

        while running:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False

            # Asegurarse de que la ventana esté siempre en la parte superior
            set_always_on_top()

            # Limpiar el buffer
            buffer_surface.fill((0, 0, 0, 0))

            # Dibujar el fotograma actual en el buffer
            buffer_surface.blit(frames[current_frame], (0, 0))

            # Actualizar la pantalla con el buffer
            ventana.blit(buffer_surface, (0, 0))
            pygame.display.update()

            current_frame = (current_frame + 1) % frame_count
            clock.tick(10)  # Ajuste para un GIF más fluido

        pygame.quit()
    except Exception as e:
        print(f"Error en la ventana principal: {e}")
        pygame.quit()
        sys.exit()

if __name__ == "__main__":
    print('Función carga en ejecución')
    mostrar_gif()