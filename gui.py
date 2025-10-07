import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
from automatizacion import main

ctk.set_appearance_mode("dark")  # "light" o "dark"
ctk.set_default_color_theme("blue")  # puedes usar "green", "dark-blue", etc.

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automatización de Aprendices SENA")
        self.geometry("700x400")

        self.label = ctk.CTkLabel(self, text="Selecciona el archivo Excel:")
        self.label.pack(pady=10)

        self.file_entry = ctk.CTkEntry(self, width=400)
        self.file_entry.pack(pady=5)

        self.browse_button = ctk.CTkButton(self, text="Buscar archivo", command=self.browse_file)
        self.browse_button.pack(pady=5)

        self.start_button = ctk.CTkButton(self, text="Iniciar proceso", command=self.start_process)
        self.start_button.pack(pady=10)

        self.progress_label = ctk.CTkLabel(self, text="")
        self.progress_label.pack()

        self.progress_bar = ctk.CTkProgressBar(self, width=600)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=5)

        self.textbox = ctk.CTkTextbox(self, width=600, height=200)
        self.textbox.pack(pady=10)

        self.progress_queue = queue.Queue()

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
        if filepath:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filepath)

    def start_process(self):
        ruta = self.file_entry.get()
        if not ruta:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Excel.")
            return
        
        self.start_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="Iniciando...")
        self.textbox.insert("end", f"Iniciando proceso para {ruta}\n")
        self.textbox.see("end")
        
        # Ejecutar en otro hilo para no congelar la interfaz
        threading.Thread(target=self.run_main, args=(ruta, self.progress_queue), daemon=True).start()
        # Iniciar el chequeo de la cola de progreso
        self.after(100, self.check_progress_queue)

    def run_main(self, ruta, progress_queue):
        try:
            # Pasamos la cola a la función principal
            main(ruta, progress_queue=progress_queue)
            progress_queue.put(("log", "✅ Proceso completado correctamente\n"))
        except Exception as e:
            progress_queue.put(("log", f"❌ Error: {e}\n"))
        finally:
            progress_queue.put(("finish", None))

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
                    self.start_button.configure(state="normal")
                    self.browse_button.configure(state="normal")
                    return # Detener el chequeo
        except queue.Empty:
            pass # La cola está vacía, seguir esperando
        self.after(100, self.check_progress_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()
