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
# Para manejar credenciales
import json
import os
from pathlib import Path

# Establecemos temas de la ventana
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

class CredentialsManager:
    """Clase para manejar las credenciales de forma segura"""
    
    def __init__(self):
        # Carpeta donde se guardarán las credenciales
        self.config_dir = Path.home() / ".sena_automation"
        self.credentials_file = self.config_dir / "credentials.json"
        self._ensure_config_dir()
    
    def _ensure_config_dir(self):
        """Crea la carpeta de configuración si no existe"""
        self.config_dir.mkdir(exist_ok=True)
    
    def save_credentials(self, username, password):
        """Guarda las credenciales en un archivo JSON"""
        data = {
            "username": username,
            "password": password
        }
        with open(self.credentials_file, 'w') as f:
            json.dump(data, f)
        return True
    
    def load_credentials(self):
        """Carga las credenciales desde el archivo JSON"""
        if not self.credentials_file.exists():
            return None, None
        
        try:
            with open(self.credentials_file, 'r') as f:
                data = json.load(f)
                return data.get("username"), data.get("password")
        except:
            return None, None
    
    def credentials_exist(self):
        """Verifica si existen credenciales guardadas"""
        return self.credentials_file.exists()


class CredentialsDialog(ctk.CTkToplevel):
    """Ventana emergente para configurar credenciales"""
    
    def __init__(self, parent, credentials_manager):
        super().__init__(parent)
        
        self.credentials_manager = credentials_manager
        self.result = None
        
        # Configuración de la ventana
        self.title("Configurar Credenciales")
        self.geometry("450x320")
        self.resizable(False, False)
        
        # Centrar la ventana
        self.transient(parent)
        self.grab_set()
        
        # Contenedor principal con padding
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        title_label = ctk.CTkLabel(
            main_container, 
            text="Ingresa tus credenciales SENA",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Campo de usuario
        self.username_label = ctk.CTkLabel(main_container, text="Usuario:")
        self.username_label.pack(pady=(10, 5), anchor="w")
        
        self.username_entry = ctk.CTkEntry(
            main_container, 
            width=380, 
            height=35,
            placeholder_text="Ingresa tu usuario"
        )
        self.username_entry.pack(pady=(0, 10))
        
        # Campo de contraseña
        self.password_label = ctk.CTkLabel(main_container, text="Contraseña:")
        self.password_label.pack(pady=(10, 5), anchor="w")
        
        self.password_entry = ctk.CTkEntry(
            main_container, 
            width=380, 
            height=35,
            placeholder_text="Ingresa tu contraseña", 
            show="•"
        )
        self.password_entry.pack(pady=(0, 20))
        
        # Frame para botones con pack en lugar de grid
        button_container = ctk.CTkFrame(main_container, fg_color="transparent")
        button_container.pack(pady=(10, 0))
        
        # Botón Guardar
        self.save_button = ctk.CTkButton(
            button_container, 
            text="✓ Guardar", 
            command=self.save_credentials,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2b9348",
            hover_color="#1f6f30"
        )
        self.save_button.pack(side="left", padx=10)
        
        # Botón Cancelar
        self.cancel_button = ctk.CTkButton(
            button_container, 
            text="✗ Cancelar", 
            command=self.cancel,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14),
            fg_color="#d62828",
            hover_color="#9d0208"
        )
        self.cancel_button.pack(side="left", padx=10)
        
        # Cargar credenciales existentes si las hay
        username, password = self.credentials_manager.load_credentials()
        if username:
            self.username_entry.insert(0, username)
        if password:
            self.password_entry.insert(0, password)
        
        # Focus en el campo de usuario
        self.username_entry.focus()
    
    def save_credentials(self):
        """Guarda las credenciales ingresadas"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            messagebox.showwarning("Advertencia", "Debes completar ambos campos.")
            return
        
        self.credentials_manager.save_credentials(username, password)
        self.result = True
        messagebox.showinfo("Éxito", "Credenciales guardadas correctamente.")
        self.destroy()
    
    def cancel(self):
        """Cancela la configuración"""
        self.result = False
        self.destroy()


class App(ctk.CTk):
    """ --- Metodo (__init__) --- """
    def __init__(self):
        
        super().__init__()
        
        # Inicializar el gestor de credenciales
        self.credentials_manager = CredentialsManager()
        
        # Establece el titulo de la ventana
        self.title("Automatización de Aprendices SENA")
        # Establece el tamaño de la ventana
        self.geometry("800x750")
        self.resizable(True, True) # Permitir redimensionar
        self.iconbitmap("Iconos/logoSena.ico")

        # Frame superior para credenciales
        credentials_frame = ctk.CTkFrame(self, fg_color="transparent")
        credentials_frame.pack(pady=10, padx=20, fill="x")
        
        # Indicador de credenciales
        self.credentials_status = ctk.CTkLabel(
            credentials_frame,
            text=self._get_credentials_status(),
            font=ctk.CTkFont(size=12)
        )
        self.credentials_status.pack(side="left", padx=10)
        
        # Botón para configurar credenciales
        self.config_credentials_button = ctk.CTkButton(
            credentials_frame,
            text="⚙️ Configurar Credenciales",
            command=self.open_credentials_dialog,
            width=200,
            fg_color="#1f538d"
        )
        self.config_credentials_button.pack(side="right", padx=10)

        # Separador
        separator = ctk.CTkFrame(self, height=2, fg_color="gray")
        separator.pack(pady=10, padx=20, fill="x")

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

    def _get_credentials_status(self):
        """Obtiene el estado de las credenciales"""
        if self.credentials_manager.credentials_exist():
            username, _ = self.credentials_manager.load_credentials()
            return f"✅ Credenciales configuradas (Usuario: {username})"
        return "⚠️ Credenciales no configuradas"
    
    def open_credentials_dialog(self):
        """Abre el diálogo para configurar credenciales"""
        dialog = CredentialsDialog(self, self.credentials_manager)
        self.wait_window(dialog)
        
        # Actualizar el indicador de estado
        self.credentials_status.configure(text=self._get_credentials_status())

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
        # Verificar que existan credenciales
        if not self.credentials_manager.credentials_exist():
            messagebox.showwarning(
                "Advertencia", 
                "Debes configurar tus credenciales antes de iniciar el proceso."
            )
            return
        
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
        self.config_credentials_button.configure(state="disabled")
        # Resetea la barra del progreso y actualiza la etiqueta de progreso
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

    """ MÉTODO run_main """
    def run_main(self, ruta, progress_queue):
        # Obtener credenciales
        username, password = self.credentials_manager.load_credentials()
        
        try:
            # Pasamos la cola y las credenciales a la función principal
            main(ruta, progress_queue=progress_queue, username=username, password=password)
            progress_queue.put(("log", "✅ Proceso completado correctamente\n"))
        except Exception as e:
            progress_queue.put(("log", f"❌ Error fatal: {e}\n"))
        finally:
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
                    self.config_credentials_button.configure(state="normal")
                    return
        except queue.Empty:
            pass
        self.after(100, self.check_progress_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()