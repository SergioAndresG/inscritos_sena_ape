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

def buscar_perfil_ocupacional(nombre_del_prgrama, mapeo):
    """
    Busca  el perfil ocupacional correspodiente al programa 
    los busca por conincidencia exacta
    """
    if not nombre_del_prgrama:
        return None
    
    # Coincidencia exacta
    if nombre_del_prgrama in mapeo:
        return mapeo[nombre_del_prgrama]
    
    # Si no encuentra conicidencia no retorna nada
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