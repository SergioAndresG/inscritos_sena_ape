import os # maneja el directorio
import logging # para configurar los loggs
import glob # para buscar archivos en la carpeta
from datetime import datetime, timedelta # para manejar fechas

def loggs(): 
    # directorio donde se guardaran los logging 
    log_dir = "Logs" 
    os.makedirs(log_dir, exist_ok=True) 
    
    # Configurar logging, se crea con la hora y fecha del mismo dia 
    log_filename = os.path.join(log_dir, f"automatizacion_aprendices_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log") 
    logging.basicConfig( filename=log_filename, level=logging.ERROR, # En los loggs saldran solo los errores 
                        format='%(asctime)s - %(levelname)s - %(message)s' ) 
    # Se eliminan los logs con mas de 7 en el archivo 
    threshold = datetime.now() - timedelta(days=7) 
    for log_file in glob.glob(os.path.join(log_dir, "*.log")): 
        file_time = datetime.fromtimestamp(os.path.getmtime(log_file)) 
        if file_time < threshold: 
            os.remove(log_file)