import json
import logging
import pandas as pd

# Definimos la función que va decodificar los datos de el archivo json
def cargar_mapeo_perfiles(ruta_json='perfilesOcupacionales/perfiles_ocupacionales.json'):
    """
    Carga el mapeo de programas a perfiles ocupacionales desde el json
    """
    try:
        import sys
        import os
        
        # Cuando se ejecuta como .exe, el directorio base es diferente
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            ruta_json = os.path.join(base_path, ruta_json)
        
        with open(ruta_json, 'r', encoding='utf-8') as f:
            contenido = json.load(f)
            return contenido
            
    except FileNotFoundError:
        logging.warning(f"No se encontro el archivo {ruta_json}")
        return {}
    except json.JSONDecodeError as e:
        logging.error(f"Error al leer el json {ruta_json}")
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

def buscar_perfil_ocupacional(nombre_del_programa, mapeo):
    """
    Busca el perfil ocupacional correspondiente al programa 
    """
    from debug_exe import log
    
    if not nombre_del_programa:
        return None
    
    # Coincidencia exacta
    if nombre_del_programa in mapeo:
        perfil = mapeo[nombre_del_programa]
        log(f" Coincidencia exacta encontrada: {perfil}")
        return perfil
    
    return None

def obtener_nombre_ficha(read_sheet):
    """
    Obtiene el nombre de la ficha desde el Excel.
    Maneja celdas combinadas buscando en múltiples columnas posibles.
    """
    fila_ficha = 1  # Fila 2 en Excel = índice 1 en Python
    
    # Intentar leer desde varias columnas posibles (B, C, D, E)
    columnas_posibles = [1, 2, 3, 4]  # Índices de columnas B, C, D, E
    
    for col_idx in columnas_posibles:
        try:
            valor = read_sheet.cell_value(fila_ficha, col_idx)
            if valor and str(valor).strip():  # Si tiene contenido
                # Verificar que contenga el número de ficha (patrón: números seguidos de guion)
                if '-' in str(valor) or str(valor).isdigit():
                    logging.info(f"Ficha encontrada en fila {fila_ficha+1}, columna {col_idx+1}: {valor}")
                    return str(valor).strip()
        except IndexError:
            continue
    
    # Si no encontró nada, intentar leer merged_cells
    try:
        merged_ranges = read_sheet.merged_cells
        for crange in merged_ranges:
            rlo, rhi, clo, chi = crange  # row_low, row_high, col_low, col_high
            if rlo <= fila_ficha < rhi:  # Si la fila está en el rango combinado
                # Leer el valor de la primera celda del rango
                valor = read_sheet.cell_value(rlo, clo)
                if valor and str(valor).strip():
                    logging.info(f"Ficha encontrada en celda combinada: {valor}")
                    return str(valor).strip()
    except AttributeError:
        # merged_cells no disponible en todas las versiones de xlrd
        pass
    
    logging.warning("No se pudo encontrar el nombre de la ficha en el archivo")
    return None

def agregar_perfil_a_json(nombre_programa, perfil_ocupacional, ruta_json='perfilesOcupacionales/perfiles_ocupacionales.json'):
    """
    Agrega un nuevo mapeo de programa -> perfil al archivo JSON
    """
    try:
        # Cargar el JSON actual
        mapeo_actual = cargar_mapeo_perfiles(ruta_json)
        
        # Agregar el nuevo mapeo
        mapeo_actual[nombre_programa] = perfil_ocupacional
        
        # Guardar de vuelta en el JSON
        with open(ruta_json, 'w', encoding='utf-8') as f:
            json.dump(mapeo_actual, f, ensure_ascii=False, indent=4)
        
        logging.info(f"Perfil agregado al JSON: {nombre_programa} -> {perfil_ocupacional}")
        return True
        
    except Exception as e:
        logging.error(f"Error al agregar perfil al JSON: {e}")
        return False