# importamos las librerias a usar para las interfaces
import customtkinter as ctk
# importamos las librerias para abrir dialogos (slección de archivos, mostrar alertas)
from tkinter import filedialog, messagebox
# ejecuta el proceso largo (main), para mantener la GUI responsiva
import threading
# permite pasar mesajes (progreso, logs) desde los logs de las funciones a la vista del usuario
import queue
# función principal que realiza la automatización
from automatizacion import main # Asumimos que main ahora acepta un stop_event

# Establecemos temas de la ventana
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

""" 
Clase App:
    - Hereda de ctk.CTk, que es la ventana principal de la aplicación
    - Define los elementos de la interfaz, incluyendo el nuevo botón de detención.
    - Gestiona los eventos de la GUI y la comunicación entre hilos.
"""
class App(ctk.CTk):
    """ --- Metodo (__init__) --- """
    def __init__(self):
        
        super().__init__()
        
        # Estado de control de hilo
        self.stop_event = threading.Event()
        self.process_thread = None
        
        # Establece el titulo y propiedades de la ventana
        self.title("Automatización de Aprendices SENA")
        self.geometry("800x700")
        self.resizable(True, True) 
        try:
            # Asegúrate de tener el icono en la ruta correcta para que esto funcione
            self.iconbitmap("Iconos/logoSena.ico")
        except:
            # Mensaje por si no encuentra el icono
            print("Advertencia: No se pudo cargar el icono 'Iconos/logoSena.ico'")

        # Configuración de layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(8, weight=1) # Fila del Textbox

        # Etiqueta para seleccionar el archivo de excel
        self.label = ctk.CTkLabel(self, text="Selecciona el archivo Excel:", font=ctk.CTkFont(size=14, weight="bold"))
        self.label.grid(row=0, column=0, pady=(20, 10), padx=20, sticky="n")

        # Un campo de texto, donde se muestra la ruta del archivo seleccionado
        self.file_entry = ctk.CTkEntry(self, width=500, placeholder_text="Ruta del archivo .xls")
        self.file_entry.grid(row=1, column=0, pady=5, padx=20, sticky="n")

        # Botón para abrir el dialogo de la selección de archivos
        self.browse_button = ctk.CTkButton(self, text="Buscar archivo", command=self.browse_file)
        self.browse_button.grid(row=2, column=0, pady=5, padx=20, sticky="n")

        # Contenedor para los botones de inicio y detención
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=3, column=0, pady=10, padx=20, sticky="n")
        self.button_frame.columnconfigure((0, 1), weight=1)

        # Botón para iniciar el proceso
        self.start_button = ctk.CTkButton(self.button_frame, text="▶️ Iniciar proceso", command=self.start_process, fg_color="#1F7A8C", hover_color="#133D50")
        self.start_button.grid(row=0, column=0, padx=10, pady=0)

        # Nuevo Botón para detener el proceso (inicialmente deshabilitado)
        self.stop_button = ctk.CTkButton(self.button_frame, text="⏹️ Detener proceso", command=self.stop_process, state="disabled", fg_color="#E05B5B", hover_color="#A53D3D")
        self.stop_button.grid(row=0, column=1, padx=10, pady=0)

        # Etiqueta que muestra el progreso
        self.progress_label = ctk.CTkLabel(self, text="Esperando archivo...", text_color="#F9F9F9")
        self.progress_label.grid(row=4, column=0, pady=(10, 0), padx=20, sticky="n")

        # Etiqueta que muestra visualmente el avance
        self.progress_bar = ctk.CTkProgressBar(self, width=600)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=5, column=0, pady=5, padx=20, sticky="n")
        
        # Un area de texto para mostrar los logs
        self.textbox = ctk.CTkTextbox(self, width=600, height=250, wrap="word", activate_scrollbars=True, font=ctk.CTkFont(family="Consolas"))
        self.textbox.grid(row=6, column=0, pady=(10, 20), padx=20, sticky="nsew")
        self.textbox.insert("end", "Consola de Logs:\n")

        # Una cola para recibir mensajes del proceso
        self.progress_queue = queue.Queue()

    """ MÉTODO browse_file """
    def browse_file(self):
        # Utilizamos filedialog para abrir un dialogo donde el usuario seleciona un archivo Excel
        filepath = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        # Actualiza self.file_entry con la ruta seleccionada
        if filepath:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filepath)
            
    """ MÉTODO start_process """
    def start_process(self):
        # Se valida que haya un archivo seleccionado
        ruta = self.file_entry.get()
        # Si no hay ninguna ruta, lanzar una ventana advirtiendo
        if not ruta:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Excel.")
            return
        
        # 1. Configuración de estado y UI
        self.stop_event.clear() # Asegura que la señal de detención esté desactivada
        self.start_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress_bar.set(0)
        self.progress_label.configure(text="Iniciando...")
        self.textbox.insert("end", f"\n--- PROCESO INICIADO ---\nIniciando proceso para {ruta}\n")
        self.textbox.see("end")
        
        # 2. Ejecutar en otro hilo para no congelar la interfaz
        self.process_thread = threading.Thread(
            target=self.run_main, 
            args=(ruta, self.progress_queue, self.stop_event), 
            daemon=True
        )
        self.process_thread.start()
        
        # 3. Iniciar el chequeo de la cola de progreso
        self.after(100, self.check_progress_queue)
        
    """ MÉTODO stop_process """
    def stop_process(self):
        # 1. Activa la señal de detención
        self.stop_event.set()
        
        # 2. Actualiza la UI de inmediato
        self.stop_button.configure(state="disabled")
        self.progress_label.configure(text="Detención solicitada...")
        self.textbox.insert("end", "⚠️ Solicitud de detención enviada. Esperando a que el proceso termine su tarea actual...\n")
        self.textbox.see("end")

    """ MÉTODO run_main """
    # define una funcion que pasa la ruta del archivo, cola de progreso y el evento de detención
    def run_main(self, ruta, progress_queue, stop_event):
        try:
            # Pasamos la cola y el evento de detención a la función principal
            main(ruta, progress_queue=progress_queue, stop_event=stop_event)

            # Verifica si la detención fue solicitada por el usuario
            if stop_event.is_set():
                progress_queue.put(("log", "🛑 Proceso detenido por el usuario.\n"))
            else:
                progress_queue.put(("log", "✅ Proceso completado correctamente.\n"))
                
        except Exception as e:
            progress_queue.put(("log", f"❌ Error fatal: {e}\n"))
        finally:
            # Envia un mensaje "finish" sin importar el resultado
            progress_queue.put(("finish", None))

    """ MÉTODO check_progress_queue """
    def check_progress_queue(self):
        try:
            while True:
                message_type, data = self.progress_queue.get_nowait()
                if message_type == "progress":
                    current, total = data
                    progress_value = current / total
                    self.progress_bar.set(progress_value)
                    self.progress_label.configure(text=f"Procesando: {current} de {total}")
                
                elif message_type == "log":
                    self.textbox.insert("end", data)
                    self.textbox.see("end")
                
                elif message_type == "finish":
                    # El proceso terminó (éxito, error o detención)
                    self.start_button.configure(state="normal")
                    self.browse_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.progress_label.configure(text="Proceso Finalizado.")
                    self.textbox.insert("end", f"--- PROCESO FINALIZADO ---\n")
                    self.textbox.see("end")
                    return # Detener el chequeo
                
        except queue.Empty:
            pass # La cola está vacía, seguir esperando
            
        # Repetir el chequeo después de 100ms
        self.after(100, self.check_progress_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()