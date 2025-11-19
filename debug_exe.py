"""Sistema de logging específico para debugging del .exe"""
import sys
import os
from datetime import datetime

class ExeLogger:
    def __init__(self):
        # Obtener la carpeta donde está el .exe
        if getattr(sys, 'frozen', False):
            # Ejecutándose como .exe
            app_path = os.path.dirname(sys.executable)
        else:
            # Ejecutándose como script
            app_path = os.path.dirname(os.path.abspath(__file__))
        
        # Crear archivo de log en la misma carpeta del ejecutable
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(app_path, f"SENA_Debug_{timestamp}.txt")
        
        try:
            # Abrir archivo
            self.file = open(self.log_file, 'w', encoding='utf-8')
            self.file.write(f"=== SENA Automation Debug Log ===\n")
            self.file.write(f"Timestamp: {timestamp}\n")
            self.file.write(f"Ejecutando desde: {sys.executable}\n")
            self.file.write(f"App path: {app_path}\n")
            self.file.write(f"Log file: {self.log_file}\n")
            self.file.write(f"Frozen: {getattr(sys, 'frozen', False)}\n\n")
            self.file.flush()
        except Exception as e:
            # Si falla, intentar crear en temp
            import tempfile
            temp_dir = tempfile.gettempdir()
            self.log_file = os.path.join(temp_dir, f"SENA_Debug_{timestamp}.txt")
            self.file = open(self.log_file, 'w', encoding='utf-8')
            self.file.write(f"=== SENA Automation Debug Log (TEMP) ===\n")
            self.file.write(f"Error original: {e}\n")
            self.file.write(f"Usando ruta temporal: {self.log_file}\n\n")
            self.file.flush()
    
    def log(self, message):
        """Escribe un mensaje en el log"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            line = f"[{timestamp}] {message}\n"
            self.file.write(line)
            self.file.flush()
            print(message)  # También imprimir en consola si existe
        except Exception as e:
            print(f"Error escribiendo log: {e}")
    
    def close(self):
        """Cierra el archivo de log"""
        try:
            self.file.write("\n=== Fin del log ===\n")
            self.file.write(f"Ubicación del archivo: {self.log_file}\n")
            self.file.close()
            print(f"\n📝 Log guardado en: {self.log_file}")
        except:
            pass

# Instancia global
_logger = None

def init_logger():
    """Inicializa el logger"""
    global _logger
    if _logger is None:
        _logger = ExeLogger()
    return _logger

def log(message):
    """Función helper para loggear"""
    global _logger
    if _logger is None:
        _logger = init_logger()
    _logger.log(message)

def close_logger():
    """Cierra el logger"""
    global _logger
    if _logger:
        _logger.close()

def get_log_path():
    """Retorna la ruta del archivo de log"""
    global _logger
    if _logger:
        return _logger.log_file
    return None