import json
import logging

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