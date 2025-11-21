import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import json
import traceback
from pathlib import Path
from automatizacion import main
from debug_exe import log
from perfilesOcupacionales.dialogo_perfil import DialogoPerfilOcupacional
from perfilesOcupacionales.gestorDePerfilesOcupacionales import agregar_perfil_a_json

# --- CONFIGURACIÓN DE ESTILOS GLOBALES ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Definición de paleta de colores personalizada (Opcional, para coherencia)
COLORS = {
    "bg_card": "#2B2B2B",       # Gris oscuro para fondos de tarjetas
    "accent": "#2CC985",        # Verde SENA (aproximado)
    "danger": "#CF6679",        # Rojo suave para errores/detener
    "text_main": "#FFFFFF",
    "text_dim": "#A0A0A0",
    "terminal_bg": "#1E1E1E",   # Fondo muy oscuro para logs
    "terminal_text": "#00FF00"  # Texto verde hacker
}

class CredentialsManager:
    def __init__(self):
        self.config_dir = Path.home() / ".sena_automation"
        self.credentials_file = self.config_dir / "credentials.json"
        self._ensure_config_dir()
    
    def _ensure_config_dir(self):
        self.config_dir.mkdir(exist_ok=True)
    
    def save_credentials(self, username, password):
        data = {"username": username, "password": password}
        with open(self.credentials_file, 'w') as f:
            json.dump(data, f)
        return True
    
    def load_credentials(self):
        if not self.credentials_file.exists():
            return None, None
        try:
            with open(self.credentials_file, 'r') as f:
                data = json.load(f)
                return data.get("username"), data.get("password")
        except:
            return None, None
    
    def credentials_exist(self):
        return self.credentials_file.exists()

class CredentialsDialog(ctk.CTkToplevel):
    """Ventana emergente estilizada"""
    
    def __init__(self, parent, credentials_manager):
        super().__init__(parent)
        self.credentials_manager = credentials_manager
        self.result = None
        
        # Configuración de ventana
        self.title("Gestión de Acceso")
        self.geometry("400x350")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Fondo principal
        self.configure(fg_color=COLORS["bg_card"])

        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(pady=(20, 10))
        ctk.CTkLabel(header_frame, text="🔐 Credenciales SENA", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack()
        ctk.CTkLabel(header_frame, text="Tus datos se guardan localmente", 
                     font=ctk.CTkFont(size=12), text_color=COLORS["text_dim"]).pack()

        # Formulario
        form_frame = ctk.CTkFrame(self, fg_color="transparent")
        form_frame.pack(pady=10, padx=30, fill="x")

        ctk.CTkLabel(form_frame, text="Usuario", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.username_entry = ctk.CTkEntry(form_frame, height=35, placeholder_text="Ej: usuario@sena.edu.co")
        self.username_entry.pack(fill="x", pady=(5, 15))

        ctk.CTkLabel(form_frame, text="Contraseña", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.password_entry = ctk.CTkEntry(form_frame, height=35, show="•", placeholder_text="••••••••")
        self.password_entry.pack(fill="x", pady=(5, 20))

        # Botones
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        ctk.CTkButton(btn_frame, text="Cancelar", command=self.cancel, 
                      fg_color="transparent", border_width=1, border_color=COLORS["danger"], 
                      text_color=COLORS["danger"], width=100).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Guardar Acceso", command=self.save_credentials, 
                      fg_color=COLORS["accent"], width=140).pack(side="left", padx=10)

        # Pre-carga
        username, password = self.credentials_manager.load_credentials()
        if username: self.username_entry.insert(0, username)
        if password: self.password_entry.insert(0, password)
        
    def save_credentials(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        if not username or not password:
            messagebox.showwarning("Atención", "Faltan datos requeridos")
            return
        self.credentials_manager.save_credentials(username, password)
        self.result = True
        self.destroy()

    def cancel(self):
        self.result = False
        self.destroy()


class App(ctk.CTk):
    """Clase principal de la aplicación con la interfaz de usuario."""
    
    def __init__(self):
        
        super().__init__()

        # ===== CONFIGURACIÓN INICIAL (PRIMERO) =====
        self.credentials_manager = CredentialsManager()
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()

        self.process_thread = None
        self.progress_queue = queue.Queue()
        
        # Estado de progreso para reanudar
        self.current_file = None
        self.last_processed_index = 0
        
        # Configuración Ventana
        self.title("Automatización SENA")
        self.geometry("700x700")  # Más ancho para 2 columnas
        self.minsize(1200, 700)
        try: 
            self.iconbitmap("Iconos/logoSena.ico")
        except: 
            pass

        # ===== HEADER SUPERIOR (CREDENCIALES) =====
        credentials_frame = ctk.CTkFrame(self, fg_color="transparent")
        credentials_frame.pack(pady=10, padx=20, fill="x")
        
        self.credentials_status = ctk.CTkLabel(
            credentials_frame,
            text=self._get_credentials_status(),
            font=ctk.CTkFont(size=12)
        )
        self.credentials_status.pack(side="left", padx=10)
        
        self.config_credentials_button = ctk.CTkButton(
            credentials_frame,
            text="⚙️ Configurar Credenciales",
            command=self.open_credentials_dialog,
            width=200,
            fg_color="#1f538d",
            hover_color="#133860"
        )
        self.config_credentials_button.pack(side="right", padx=10)

        # Separador
        separator = ctk.CTkFrame(self, height=2, fg_color="gray")
        separator.pack(pady=10, padx=20, fill="x")

        # ===== CONTENEDOR PRINCIPAL (2 COLUMNAS) =====
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=10)

        # --- COLUMNA IZQUIERDA: CONTROLES ---
        left_panel = ctk.CTkFrame(main_container, fg_color=COLORS["bg_card"], corner_radius=10)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # Tarjeta de archivo
        file_card = ctk.CTkFrame(left_panel, fg_color=COLORS["terminal_bg"], corner_radius=8)
        file_card.pack(pady=15, padx=15, fill="x")

        ctk.CTkLabel(file_card, text=" Archivo Excel", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10,5))

        path_frame = ctk.CTkFrame(file_card, fg_color="transparent")
        path_frame.pack(fill="x", padx=10, pady=(0,10))

        self.file_entry = ctk.CTkEntry(path_frame, height=40, 
                                        placeholder_text="Ningún archivo seleccionado...")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0,10))

        self.browse_button = ctk.CTkButton(path_frame, text="Buscar", width=100,
                                            command=self.browse_file)
        self.browse_button.pack(side="right")

        self.file_status = ctk.CTkLabel(file_card, text="", 
                                        font=ctk.CTkFont(size=11),
                                        text_color=COLORS["text_dim"])
        self.file_status.pack(anchor="w", padx=10, pady=(0,10))

        # Frame de acciones
        actions_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        actions_frame.pack(pady=20, padx=15)

        self.start_button = ctk.CTkButton(
            actions_frame,
            text="▶  Iniciar Proceso",
            command=self.start_process,
            height=45,
            width=200,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=COLORS["accent"],
            hover_color="#25B574"
        )
        self.start_button.pack(pady=5)

        self.stop_button = ctk.CTkButton(
            actions_frame,
            text="⏹  Detener",
            command=self.stop_process,
            height=45,
            width=200,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=COLORS["danger"],
            hover_color="#B5555D",
            state="disabled"
        )
        self.stop_button.pack(pady=5)

        self.pause_button = ctk.CTkButton(
            actions_frame,
            text="⏸  Pausar",
            command=self.toggle_pause,
            height=45,
            width=200,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#FFA500",  # Naranja
            hover_color="#CC8400",
            state="disabled"
        )
        self.pause_button.pack(pady=5)

        # --- COLUMNA DERECHA: MONITOREO ---
        right_panel = ctk.CTkFrame(main_container, fg_color=COLORS["bg_card"], corner_radius=10)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # Container de progreso
        progress_container = ctk.CTkFrame(right_panel, fg_color="transparent")
        progress_container.pack(pady=15, padx=15, fill="x")

        stats_frame = ctk.CTkFrame(progress_container, fg_color="transparent")
        stats_frame.pack(fill="x", pady=(0,10))

        self.progress_label = ctk.CTkLabel(stats_frame, text="Listo para comenzar",
                                            font=ctk.CTkFont(size=13, weight="bold"))
        self.progress_label.pack(side="left")

        self.progress_percentage = ctk.CTkLabel(stats_frame, text="0%",
                                                font=ctk.CTkFont(size=13, weight="bold"),
                                                text_color=COLORS["accent"])
        self.progress_percentage.pack(side="right")

        # Barra de progreso
        self.progress_bar = ctk.CTkProgressBar(progress_container, height=20)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x")

        # Logs
        logs_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        logs_frame.pack(pady=(10,15), padx=15, fill="both", expand=True)

        ctk.CTkLabel(logs_frame, text="📋 Registro de Actividad",
                    font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(0,10))

        self.textbox = ctk.CTkTextbox(logs_frame, 
                                    fg_color=COLORS["terminal_bg"],
                                    text_color="#00FF00",
                                    font=ctk.CTkFont(family="Consolas", size=11),
                                    corner_radius=8)
        self.textbox.pack(fill="both", expand=True)

    def _get_credentials_status(self):
        """Obtiene el estado de las credenciales"""
        if self.credentials_manager.credentials_exist():
            username, _ = self.credentials_manager.load_credentials()
            return f"✓ Credenciales configuradas (Usuario: {username})"
        return " Credenciales no configuradas"
    
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

    def toggle_pause(self):
        """Alterna entre pausar y reanudar el proceso"""
        
        if self.pause_event.is_set():
            # Actualmente en ejecución → PAUSAR
            self.pause_event.clear()
            
            self.pause_button.configure(
                text="▶ Reanudar",
                fg_color="#2CC985"  # Verde
            )
            self.progress_label.configure(text="⏸ Proceso pausado")
            
            self.textbox.insert("end", "\n" + "="*50 + "\n")
            self.textbox.insert("end", "⏸ PROCESO PAUSADO\n")
            self.textbox.insert("end", "="*50 + "\n")
            self.textbox.insert("end", " Presiona 'Reanudar' para continuar\n\n")
            self.textbox.see("end")
            
        else:
            # Actualmente pausado → REANUDAR
            self.pause_event.set()
            
            self.pause_button.configure(
                text="⏸️  Pausar",
                fg_color="#FFA500"  # Naranja
            )
            self.progress_label.configure(text="▶ Reanudando proceso...")
            
            self.textbox.insert("end", "\n" + "="*50 + "\n")
            self.textbox.insert("end", "▶ PROCESO REANUDADO\n")
            self.textbox.insert("end", "="*50 + "\n\n")
            self.textbox.see("end")

    """ MÉTODO start_process """
    def start_process(self):
        if not self.credentials_manager.credentials_exist():
            messagebox.showwarning(
                "Advertencia", 
                "Debes configurar tus credenciales antes de iniciar el proceso."
            )
            return
        
        ruta = self.file_entry.get()
        if not ruta:
            messagebox.showwarning("Advertencia", "Debes seleccionar un archivo Excel.")
            return
        
        # Limpiar cola de mensajes antiguos
        while not self.progress_queue.empty():
            try:
                self.progress_queue.get_nowait()
            except queue.Empty:
                break
        
        # Limpiar UI
        self.textbox.delete("1.0", "end")
        self.progress_bar.set(0)
        self.progress_percentage.configure(text="0%")
        self.progress_bar.configure(progress_color=COLORS["accent"])
        
        # Configuración UI
        self.stop_event.clear()
        self.start_button.configure(state="disabled", text="⏳ Iniciando...")
        self.browse_button.configure(state="disabled")
        self.config_credentials_button.configure(state="disabled")
        self.stop_button.configure(state="normal")

        # Habilitar botón de pausa cuando inicia
        self.pause_button.configure(state="normal")
        
        # Asegurarse de que NO esté pausado al iniciar
        self.pause_event.set()
        
        # Progreso indeterminado
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        
        self.progress_label.configure(text="Preparando automatización...")
        self.progress_percentage.configure(text="...")
        
        # Logs iniciales
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        self.textbox.insert("end", f"{'='*50}\n")
        self.textbox.insert("end", f"▶ PROCESO INICIADO - {timestamp}\n")
        self.textbox.insert("end", f"{'='*50}\n")
        self.textbox.insert("end", f" Archivo: {ruta}\n\n")
        self.textbox.see("end")
        
        # Iniciar thread
        self.process_thread = threading.Thread(
            target=self.run_main, 
            args=(ruta, self.progress_queue, self.stop_event), 
            daemon=True
        )
        self.process_thread.start()
        self.after(100, self.check_progress_queue)

    """ MÉTODO stop_process """
    def stop_process(self):
        """Detiene el proceso de automatización"""
        
        # Activar señal de detención
        self.stop_event.set()
        
        # 2. Actualizar UI inmediatamente
        self.stop_button.configure(state="disabled", text="⏹ Deteniendo...")
        self.progress_label.configure(text="Detención solicitada...")
        
        # Log detallado
        self.textbox.insert("end", "\n" + "="*50 + "\n")
        self.textbox.insert("end", "DETENCIÓN SOLICITADA\n")
        self.textbox.insert("end", "="*50 + "\n")
        self.textbox.insert("end", " Esperando que termine la tarea actual...\n")
        self.textbox.insert("end", "El proceso se detendrá en el próximo punto seguro\n\n")
        self.textbox.see("end")
        
        # Deshabilitar inicio mientras se detiene
        self.start_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        self.config_credentials_button.configure(state="disabled")
        
        # Iniciar verificación de detención
        self.check_stop_completion()


    def check_stop_completion(self):
        """Verifica si el proceso se detuvo completamente"""
        
        if self.process_thread and self.process_thread.is_alive():
            # El thread sigue vivo, verificar de nuevo en 500ms
            self.after(500, self.check_stop_completion)
        else:
            # El thread terminó
            self.textbox.insert("end", "✓  Proceso detenido correctamente\n\n")
            self.textbox.see("end")
            
            # Resetear UI
            self._reset_ui_after_dialog()
    
    def _reset_ui_after_dialog(self):
        """Resetea la UI después de cerrar el diálogo (SIN reiniciar proceso)"""
        self.start_button.configure(state="normal", text="▶ Iniciar Proceso")
        self.browse_button.configure(state="normal")
        self.config_credentials_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.pause_button.configure(state="disabled") 
        self.progress_label.configure(text="Listo para comenzar")
        self.progress_bar.set(0)
        self.progress_percentage.configure(text="0%")

    def _reset_ui_for_restart(self):
        """Resetea la UI preparándola para reiniciar el proceso"""
        # No cambiar los botones aquí - start_process() lo hará
        self.progress_bar.set(0)
        self.progress_label.configure(text="Preparando reinicio...")
        self.textbox.insert("end", "\n")


    def show_dialog_profile(self, nombre_programa):
        """Muestra el diálogo para solicitar un perfil ocupacional"""
        try:
            # Deshabilitar botones de control
            self.pause_button.configure(state="disabled")
            self.stop_button.configure(state="disabled")
            
            # Crear y mostrar el diálogo (BLOQUEA hasta que el usuario responda)
            dialogo = DialogoPerfilOcupacional(self, nombre_programa)
            self.wait_window(dialogo)
            
            if dialogo.resultado:
                perfil_ingresado = dialogo.resultado
                exito = agregar_perfil_a_json(nombre_programa, perfil_ingresado)
                
                if exito:
                    self.textbox.insert("end", f"Perfil agregado: {nombre_programa} -> {perfil_ingresado}\n")
                    self.textbox.see("end")
                    
                    respuesta = messagebox.askyesno(
                        "Perfil Agregado",
                        f"El perfil '{perfil_ingresado}' ha sido agregado correctamente.\n\n"
                        f"¿Deseas reiniciar el proceso automáticamente?"
                    )
                    
                    if respuesta:
                        self.textbox.insert("end", f"Reiniciando proceso automáticamente...\n\n")
                        self.textbox.see("end")
                        
                        # Limpiar UI
                        self._reset_ui_for_restart()
                        
                        # Reiniciar después de un breve delay
                        self.after(500, self.start_process)
                    else:
                        self.textbox.insert("end", f"ℹ Reinicia manualmente cuando estés listo.\n\n")
                        self.textbox.see("end")
                        self._reset_ui_after_dialog()
                else:
                    self.textbox.insert("end", f" Error al guardar el perfil\n")
                    messagebox.showerror("Error", "No se pudo guardar el perfil en el archivo JSON")
                    self._reset_ui_after_dialog()
            else:
                # Usuario canceló
                self.textbox.insert("end", f"⏭Configuración de perfil cancelada\n\n")
                self.textbox.see("end")
                self._reset_ui_after_dialog()
                
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            self.textbox.insert("end", f"❌ Error en diálogo: {e}\n{error_detail}\n")
            self.textbox.see("end")
            self._reset_ui_after_dialog()

    """ MÉTODO run_main """
    def run_main(self, ruta, progress_queue, stop_event): 
        # Obtener credenciales
        username, password = self.credentials_manager.load_credentials()
        
        try:
            # Pasamos todos los argumentos, incluyendo credenciales y stop_event
            main(ruta, progress_queue=progress_queue, username=username, password=password, stop_event=stop_event, pause_event=self.pause_event)
            
            # Revisar el estado de detención para reportar el resultado final
            if stop_event.is_set():
                progress_queue.put(("log", "🛑 Proceso detenido por el usuario.\n"))
                progress_queue.put(("finish", None))
            else:
                progress_queue.put(("log", "✓  Proceso completado correctamente\n"))
                
        except Exception as e:
            error_msg = traceback.format_exc()
            progress_queue.put(("log", f"✗ Error inesperado:\n{error_msg}\n"))
            progress_queue.put(("finish", None))
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
                    percentage = int(progress_value * 100)
                    
                    if self.progress_bar.cget("mode") == "indeterminate":
                        self.progress_bar.stop()
                        self.progress_bar.configure(mode="determinate")
                    
                    self.progress_bar.set(progress_value)
                    self.progress_percentage.configure(text=f"{percentage}%")
                    self.progress_label.configure(text=f"Procesando: {current} de {total}")
                    
                    if percentage < 30:
                        self.progress_bar.configure(progress_color="#FF6B6B")
                    elif percentage < 70:
                        self.progress_bar.configure(progress_color="#FFD93D")
                    else:
                        self.progress_bar.configure(progress_color=COLORS["accent"])
                            
                elif message_type == "log":
                    self.textbox.insert("end", data)
                    self.textbox.see("end")

                elif message_type == "solicitar_perfil":
                    nombre_programa = data
                    
                    # Actualizar UI para mostrar que está esperando
                    self.progress_label.configure(text=f"⏸ Esperando configuración de perfil...")
                    if self.progress_bar.cget("mode") == "indeterminate":
                        self.progress_bar.stop()
                        self.progress_bar.configure(mode="determinate")
                    
                    # Deshabilitar botones mientras se muestra el diálogo
                    self.stop_button.configure(state="disabled")
                    self.pause_button.configure(state="disabled")
                    
                    # Mostrar diálogo (ejecutado en el thread principal - correcto)
                    self.show_dialog_profile(nombre_programa)
                    
                    # No continuar procesando mensajes - show_dialog_profile maneja el resto
                    return

                elif message_type == "finish":
                    if self.progress_bar.cget("mode") == "indeterminate":
                        self.progress_bar.stop()
                        self.progress_bar.configure(mode="determinate")
                    
                    self.pause_button.configure(state="disabled")
                    self.start_button.configure(state="normal", text="▶ Iniciar Proceso")
                    self.browse_button.configure(state="normal")
                    self.config_credentials_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.progress_label.configure(text="Proceso Finalizado")
                    
                    self.textbox.insert("end", f"\n{'='*50}\n")
                    self.textbox.insert("end", "✓ PROCESO FINALIZADO\n")
                    self.textbox.insert("end", f"{'='*50}\n")
                    self.textbox.see("end")
                    return
                        
        except queue.Empty:
            pass
        
        # Continuar verificando la cola
        self.after(100, self.check_progress_queue)

if __name__ == "__main__":
    app = App()
    app.mainloop()