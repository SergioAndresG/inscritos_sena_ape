import json
import logging
import pandas as pd

# Definimos la función que va decodificar los datos de el archivo json
def cargar_mapeo_perfiles(ruta_json='perfiles_ocupacionales.json'):
    """
    Carga el mapeo de programas a perfiles ocupacionales desde el json
    """
    try:
        # abrimos la ruta del archivo
        with open(ruta_json, 'r', encoding='utf-8') as f: # decodificamos en formato utf-8
            return json.load(f) # retornamos decodificado
    except FileNotFoundError:
        # Si no existe la ruta devolvemos un error
        logging.warning(f"No se encontro el archivo {ruta_json}")
        return {}
    except json.JSONDecodeError:
        # Si hay un error mas profundo retornamos el error exacto
        logging.error(f"Error al lerr el json {ruta_json}")
        return {}
    
def extraer_nombre_ficha(nombre_ficha):
    """
    Extrae el nombre de la "ficha de caracterizacion"
    Ejemplo => "3150717 - GESTION DEL TALENTO HUMANO" -> "GESTION DEL TALENTO HUMANO"
    """
    if not nombre_ficha or pd.isna(nombre_ficha):
        return None
    
    #Dividir por el guion y tomar segunda parte
    partes = str(nombre_ficha).split("-", 1)
    if len(partes) > 1:
        return partes[1].strip().upper()
    return str(nombre_ficha).strip().upper()

