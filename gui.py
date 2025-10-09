# importamos las librerias a usar para las interfaces
import customtkinter as ctk
# importamos las librerias para abrir dialogos (slección de archivos, mostrar alertas)
from tkinter import filedialog, messagebox
# ejecuta el proceso largo (main), para mantener la GUI responsiva
import threading
# permite pasar mesajes (progreso, logs) desde los logs de las funciones a la vista del usuario
import queue
# función principal que realiza la automatización
from automatizacion import main

# Establecemos temas de la ventana
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")
                                                                                                                                                                                                                                                                                                                                                                       
""" 
Clase App:
    - Hereda de ctk.CTk, que es la ventana principal de la aplicación
    - Define los elementos de la interfaz (etiquetas, entradas de texto, botones, barra de progreso, etc)
    - Gestiona los eventos como seleccionar un archivo iniciar el proceso y cerrar la ventana
"""
class App(ctk.CTk):
    """ --- Metodo (__init__) --- """
    def __init__(self):
        
        super().__init__()
        
        # Establece el titulo de la ventana
        self.title("Automatización de Aprendices SENA")
        # Establece el tamaño de la ventana
        self.geometry("800x700")
        self.resizable(True, True) # Permitir redimensionar
        self.iconbitmap("Iconos/logoSena.ico")

        # Etiqueta para seleccionar el archivo de excel
        self.label = ctk.CTkLabel(self, text="Selecciona el archivo Excel:")
        self.label.pack(pady=10)

        # Un campo de texto, donde se muestra la ruta del archivo seleccionado
        self.file_entry = ctk.CTkEntry(self, width=400)
        self.file_entry.pack(pady=5)

        # Botón para abrir el dialogo de la selección de archivos
        self.browse_button = ctk.CTkButton(self, text="Buscar archivo", command=self.browse_file)
        self.browse_button.pack(pady=5)

        # Botón para iniciar el proceso
        self.start_button = ctk.CTkButton(self, text="Iniciar proceso", command=self.start_process)
        self.start_button.pack(pady=10)

        # Etiqueta que muestra el progreso
        self.progress_label = ctk.CTkLabel(self, text="")
        self.progress_label.pack()

        # Etiqueta que muestra visualmente el avance
        self.progress_bar = ctk.CTkProgressBar(self, width=600)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=5)
        
        # Un area de texto para mostrar los logs
        self.textbox = ctk.CTkTextbox(self, width=600, height=200)
        self.textbox.pack(pady=10)

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
            
        """ 
        - Resumen:
            Cuando se hace click en "Buscar Archivo". se abre el explorador de archivos. Si
            se selecciona un archivo, su ruta se muestras en el campo de texto.
        """

    """ MÉTODO start_process """
    def start_process(self):
        # Se valida que haya un archivo seleccionado
        ruta = self.file_entry.get()
        # Si no hay ninguna ruta, lanzar una ventana advirtiendo
        if not ruta:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Excel.")
            return
        
        # Se desahabilitan los botones para evitar acciones durante el procesammiento
        self.start_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        # Resetea la barra del progreso y actualiza la etiqueta de progreso
        self.progress_bar.set(0)
        # Muestra un mensaje en el textbox
        self.progress_label.configure(text="Iniciando...")
        self.textbox.insert("end", f"Iniciando proceso para {ruta}\n")
        self.textbox.see("end")
        
        # Ejecutar en otro hilo para no congelar la interfaz
        threading.Thread(target=self.run_main, args=(ruta, self.progress_queue), daemon=True).start()
        # Iniciar el chequeo de la cola de progreso
        self.after(100, self.check_progress_queue)
        
        """
        - Resumen:
            Al hacer click en iniciar proceso, se aseguar que haya un archivo
            se confgura la interfaz para le procesamiento y ejecuta la logica en un
            hilo separado para no congelar la GUI.
        
        """

    """ MÉTODO run_main """
    # define una funcion que pasa la ruta del archivo y cola de progreso
    def run_main(self, ruta, progress_queue):
        # captura la exepciones y envia los mensajes en cola
        try:
            # Pasamos la cola a la función principal
            main(ruta, progress_queue=progress_queue)
            progress_queue.put(("log", "✅ Proceso completado correctamente\n"))
        except Exception as e:
            progress_queue.put(("log", f"❌ Error: {e}\n"))
        finally:
            # envia un mensaje "finish" al final 
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
                    #Actualiza la barra de progreso
                    self.progress_label.configure(text=f"Procesando: {current} de {total}")
                #inserta un mensaje en el textbox (texto de los print() o errores)
                elif message_type == "log":
                    self.textbox.insert("end", data)
                    self.textbox.see("end")
                elif message_type == "finish":
                    self.start_button.configure(state="normal")
                    self.browse_button.configure(state="normal")
                    return # Detener el chequeo
        except queue.Empty:
            pass # La cola está vacía, seguir esperando
        self.after(100, self.check_progress_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()
