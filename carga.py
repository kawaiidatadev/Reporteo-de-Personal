from common.__init__ import *

"""
Este es el codigo que se debe ejecutar primero que nada, en un proceso separado de main, en segundo plano,
utilizando recursos diferentes y ejecutandose en paralelo mientras que se termina de ejecutar main.py.
Este codigo debe de abrir en el centro de la pantalla de cualquier usuario, el gif con movimiento
y ponerlo siempre por encima de todas las ventanas de windows, y finalizar despues de que main.py finalize.
"""

def mostrar_gif():
    try:
        ventana = tk.Tk()
        ventana.overrideredirect(True)
        ventana.title("Cargando...")

        # Centrar ventana en la pantalla
        ventana_width = 500
        ventana_height = 300
        screen_width = ventana.winfo_screenwidth()
        screen_height = ventana.winfo_screenheight()
        x = (screen_width // 2) - (ventana_width // 2)
        y = (screen_height // 2) - (ventana_height // 2)
        ventana.geometry(f'{ventana_width}x{ventana_height}+{x}+{y}')

        # Cargar el GIF
        gif_path = r'resources\gif_de_carga.gif'
        img = Image.open(gif_path)
        gif_frames = [ImageTk.PhotoImage(img.copy().convert("RGBA")) for _ in range(img.n_frames)]

        label = tk.Label(ventana)
        label.pack()

        def actualizar_gif(frame_index=0):
            frame = gif_frames[frame_index]
            label.config(image=frame)
            frame_index = (frame_index + 1) % len(gif_frames)
            ventana.after(100, actualizar_gif, frame_index)

        actualizar_gif()
        ventana.attributes('-topmost', True)
        ventana.protocol("WM_DELETE_WINDOW", lambda: None)

        # Mantener la ventana abierta
        ventana.mainloop()

    except Exception as e:
        print(f"Error en la ventana principal: {e}")

if __name__ == "__main__":
    print('Función carga en ejecución')
    mostrar_gif()
    print('Función carga termino')