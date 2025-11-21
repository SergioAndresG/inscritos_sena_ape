import json
import os
import sys
from pathlib import Path
import shutil
import logging

def obtener_ruta_archivo_json():
    """
    Obtiene la ruta correcta del archivo JSON tanto en desarrollo como en .exe
    Si el archivo no existe en la carpeta de usuario, lo copia desde los recursos
    """
    # Directorio de configuración del usuario
    if sys.platform == "win32":
        config_dir = Path.home() / ".sena_automation"
    else:
        config_dir = Path.home() / ".config" / "sena_automation"
    
    # Crear directorio si no existe
    config_dir.mkdir(parents=True, exist_ok=True)
    
    # Ruta del archivo en la carpeta del usuario
    user_json_path = config_dir / "perfiles_ocupacionales.json"
    
    # Si no existe, copiar desde recursos empaquetados
    if not user_json_path.exists():
        # Obtener ruta de recursos empaquetados
        if getattr(sys, 'frozen', False):
            # Ejecutando como .exe
            base_path = sys._MEIPASS
        else:
            # Ejecutando como script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        source_json = os.path.join(base_path, 'perfilesOcupacionales', 'perfiles_ocupacionales.json')
        
        # Si existe el archivo fuente, copiarlo
        if os.path.exists(source_json):
            shutil.copy2(source_json, user_json_path)
            logging.info(f"Archivo JSON copiado a: {user_json_path}")
            print(f"✓ Archivo de perfiles inicializado en: {user_json_path}")
        else:
            # Crear archivo vacío
            with open(user_json_path, 'w', encoding='utf-8') as f:
                json.dump({}, f, ensure_ascii=False, indent=4)
            logging.warning(f"Creado archivo JSON vacío en: {user_json_path}")
            print(f"⚠️ Creado archivo de perfiles vacío en: {user_json_path}")
    
    return user_json_path


def cargar_mapeo_perfiles():
    """
    Carga el mapeo de perfiles ocupacionales desde el archivo JSON
    Usa la ruta correcta según si es .exe o desarrollo
    """
    try:
        json_path = obtener_ruta_archivo_json()
        
        if not os.path.exists(json_path):
            logging.error(f"No se encontró el archivo: {json_path}")
            return None
        
        with open(json_path, 'r', encoding='utf-8') as f:
            mapeo = json.load(f)
        
        logging.info(f"Mapeo cargado correctamente desde: {json_path}")
        return mapeo
        
    except FileNotFoundError:
        logging.error(f"Archivo no encontrado: {json_path}")
        return None
    except json.JSONDecodeError as e:
        logging.error(f"Error al decodificar JSON: {e}")
        return None
    except Exception as e:
        logging.error(f"Error inesperado al cargar mapeo: {e}")
        return None


def agregar_perfil_a_json(nombre_programa, perfil_ocupacional):
    """
    Agrega un nuevo perfil ocupacional al archivo JSON
    """
    try:
        json_path = obtener_ruta_archivo_json()
        
        # Cargar mapeo existente
        mapeo = cargar_mapeo_perfiles()
        if mapeo is None:
            mapeo = {}
        
        # Agregar nuevo perfil
        mapeo[nombre_programa] = perfil_ocupacional
        
        # Guardar de vuelta al archivo
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(mapeo, f, ensure_ascii=False, indent=4)
        
        logging.info(f"Perfil agregado: {nombre_programa} -> {perfil_ocupacional}")
        print(f"✓ Perfil guardado en: {json_path}")
        return True
        
    except PermissionError:
        logging.error(f"Sin permisos para escribir en: {json_path}")
        print(f"❌ Sin permisos de escritura en: {json_path}")
        return False
    except Exception as e:
        logging.error(f"Error al guardar perfil: {e}")
        print(f"❌ Error al guardar: {e}")
        import traceback
        traceback.print_exc()
        return False


def buscar_perfil_ocupacional(nombre_programa, mapeo_perfiles):
    """
    Busca el perfil ocupacional correspondiente al programa
    Retorna None si no lo encuentra
    """
    if not mapeo_perfiles or not nombre_programa:
        return None
    
    # Buscar en el mapeo
    perfil = mapeo_perfiles.get(nombre_programa)
    
    # Validar que el perfil no esté vacío
    if perfil and str(perfil).strip():
        return str(perfil).strip()
    
    return None


def extraer_nombre_ficha(ficha_completa):
    """
    Extrae el nombre del programa de la ficha de caracterización
    Ejemplo: "3309028 - ELECTRICIDAD BÁSICA" -> "ELECTRICIDAD BÁSICA"
    """
    try:
        if not ficha_completa:
            return None
        
        # Separar por guion
        partes = str(ficha_completa).split('-', 1)
        
        if len(partes) >= 2:
            nombre_programa = partes[1].strip()
            logging.info(f"Nombre extraído: {nombre_programa}")
            return nombre_programa
        
        # Si no hay guion, retornar todo
        return str(ficha_completa).strip()
        
    except Exception as e:
        logging.error(f"Error extrayendo nombre: {e}")
        return None


def obtener_nombre_ficha(read_sheet):
    """
    Obtiene el nombre de la ficha desde el Excel
    Compatible con xlrd (para .xls)
    """
    try:
        fila_ficha = 1  # Fila 2 en Excel
        
        # Intentar leer desde varias columnas posibles
        columnas_posibles = [1, 2, 3, 4]  # B, C, D, E
        
        for col_idx in columnas_posibles:
            try:
                valor = read_sheet.cell_value(fila_ficha, col_idx)
                if valor and str(valor).strip():
                    if '-' in str(valor) or str(valor).isdigit():
                        logging.info(f"Ficha encontrada: {valor}")
                        return str(valor).strip()
            except (IndexError, AttributeError):
                continue
        
        # Intentar con celdas combinadas
        try:
            if hasattr(read_sheet, 'merged_cells'):
                merged_ranges = read_sheet.merged_cells
                for crange in merged_ranges:
                    rlo, rhi, clo, chi = crange
                    if rlo <= fila_ficha < rhi:
                        valor = read_sheet.cell_value(rlo, clo)
                        if valor and str(valor).strip():
                            logging.info(f"Ficha en celda combinada: {valor}")
                            return str(valor).strip()
        except AttributeError:
            pass
        
        logging.warning("No se encontró ficha de caracterización")
        return None
        
    except Exception as e:
        logging.error(f"Error obteniendo nombre de ficha: {e}")
        return None