import logging
import pandas as pd
import os

def convertir_xls_a_xlsx(ruta_xls):
    """
    Convierte un archivo .xls a .xlsx y elimina el original
    Retorna la nueva ruta .xlsx
    """
    try:
        if not ruta_xls.endswith('.xls'):
            return ruta_xls  # Ya es xlsx o no necesita conversión
        
        print(f"Convirtiendo {os.path.basename(ruta_xls)} a formato .xlsx...")
        logging.info(f"Iniciando conversión de .xls a .xlsx: {ruta_xls}")
        
        # Leer archivo .xls completo con xlrd
        df_completo = pd.read_excel(ruta_xls, header=None, engine='xlrd')
        
        # Crear ruta .xlsx
        ruta_xlsx = ruta_xls.replace('.xls', '.xlsx')
        
        # Guardar como .xlsx
        df_completo.to_excel(ruta_xlsx, index=False, header=False, engine='openpyxl')
        
        print(f"✓ Archivo convertido: {os.path.basename(ruta_xlsx)}")
        logging.info(f"Conversión exitosa: {ruta_xlsx}")
        
        # Eliminar archivo .xls original
        try:
            os.remove(ruta_xls)
            print(f"Archivo .xls original eliminado")
            logging.info(f"Archivo .xls eliminado: {ruta_xls}")
        except Exception as e:
            print(f"No se pudo eliminar el .xls original: {e}")
            logging.warning(f"No se pudo eliminar .xls: {e}")
        
        return ruta_xlsx
        
    except Exception as e:
        logging.error(f"Error en conversión XLS → XLSX: {e}")
        print(f"✗ Error en conversión: {e}")
        raise