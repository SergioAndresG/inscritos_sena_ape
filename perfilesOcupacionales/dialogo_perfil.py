import customtkinter as ctk
from tkinter import messagebox

class DialogoPerfilOcupacional(ctk.CTkToplevel):
    """Diálogo para solicitar un perfil ocupacional no encontrado"""
    
    def __init__(self, parent, nombre_programa):
        super().__init__(parent)
        
        self.nombre_programa = nombre_programa
        self.resultado = None  # Almacenará el perfil ingresado
        
        # Configuración de la ventana
        self.title("Perfil Ocupacional No Encontrado")
        self.geometry("550x350")
        self.resizable(False, False)
        
        # Hacer modal
        self.transient(parent)
        self.grab_set()
        
        # Contenedor principal
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título con advertencia
        warning_label = ctk.CTkLabel(
            main_container,
            text="⚠️ Perfil Ocupacional No Encontrado",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#FFA500"
        )
        warning_label.pack(pady=(0, 10))
        
        # Mensaje explicativo
        info_label = ctk.CTkLabel(
            main_container,
            text=f"No se encontró un perfil ocupacional para:\n\n'{nombre_programa}'\n\n¿Deseas agregarlo al sistema?",
            font=ctk.CTkFont(size=13),
            wraplength=500,
            justify="center"
        )
        info_label.pack(pady=(0, 20))
        
        # Campo para ingresar el perfil
        perfil_label = ctk.CTkLabel(
            main_container,
            text="Ingresa el perfil ocupacional del APE:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        perfil_label.pack(pady=(10, 5), anchor="w")
        
        self.perfil_entry = ctk.CTkEntry(
            main_container,
            width=480,
            height=40,
            placeholder_text="Ejemplo: Auxiliar administrativo",
            font=ctk.CTkFont(size=13)
        )
        self.perfil_entry.pack(pady=(0, 20))
        
        # Frame para botones
        button_container = ctk.CTkFrame(main_container, fg_color="transparent")
        button_container.pack(pady=(10, 0))
        
        # Botón Guardar
        save_button = ctk.CTkButton(
            button_container,
            text="✓ Guardar y Continuar",
            command=self.guardar_perfil,
            width=200,
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2b9348",
            hover_color="#1f6f30"
        )
        save_button.pack(side="left", padx=10)
        # Focus en el campo de entrada
        self.perfil_entry.focus()
        
        # Vincular Enter para guardar
        self.perfil_entry.bind("<Return>", lambda e: self.guardar_perfil())
    
    def guardar_perfil(self):
        """Guarda el perfil ingresado"""
        perfil = self.perfil_entry.get().strip()
        
        if not perfil:
            messagebox.showwarning("Advertencia", "Debes ingresar un perfil ocupacional.")
            return
        
        self.resultado = perfil
        self.destroy()
    
    def saltar(self):
        """Salta este estudiante sin agregar perfil"""
        self.resultado = None
        self.destroy()