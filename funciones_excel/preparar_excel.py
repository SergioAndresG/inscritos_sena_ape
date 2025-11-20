import os
import logging
import pandas as pd
import xlrd
from xlutils.copy import copy
from perfilesOcupacionales.perfilExcepcion import PerfilOcupacionalNoEncontrado
from perfilesOcupacionales.gestorDePerfilesOcupacionales import (
    extraer_nombre_ficha, 
    cargar_mapeo_perfiles, 
    buscar_perfil_ocupacional, 
    obtener_nombre_ficha
)

# Mapeo de tipos de documento
TIPOS_DOCUMENTO = {
    "CC": "1",
    "TI": "2", 
    "CE": "3",
    "Otro Documento de Identidad": "5",
    "PEP": "8",
    "PPT": "9"
}

def preparar_excel(ruta_excel):
    """
    Prepara el archivo Excel .xls para procesamiento
    ADVERTENCIA: xlutils puede corromper archivos .xls complejos
    Se recomienda usar formato .xlsx
    """
    try:
        # ===== VALIDAR ARCHIVO =====
        programa_sin_perfil = None

        if not os.path.exists(ruta_excel):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_excel}")
        
        logging.info(f"Abriendo archivo: {ruta_excel}")
        
        # ===== CREAR BACKUP =====
        import shutil
        backup_path = ruta_excel.replace('.xls', '_backup.xls')
        shutil.copy2(ruta_excel, backup_path)
        logging.info(f"Backup creado: {backup_path}")
        
        # ===== CONFIGURAR PANDAS =====
        pd.set_option('display.float_format', lambda x: '%.0f' % x)
        
        # ===== LEER CON PANDAS =====
        df = pd.read_excel(
            ruta_excel, 
            header=4,
            dtype={'Número de Documento': str, 'Celular': str}, 
            engine='xlrd'
        )
        
        # ===== CARGAR CON XLRD =====
        rb = xlrd.open_workbook(ruta_excel, formatting_info=True)
        read_sheet = rb.sheet_by_index(0)
        
        # ===== MAPEAR COLUMNAS =====
        expected_columns = [
            'Tipo de Documento', 'Número de Documento', 'Nombre', 
            'Apellidos', 'Celular', 'Correo Electrónico', 
            'Estado', 'Perfil Ocupacional'
        ]
        
        header_row = 4
        column_indices = {}
        
        for col in range(read_sheet.ncols):
            cell_value = read_sheet.cell_value(header_row, col)
            for expected_column in expected_columns:
                if cell_value and expected_column in cell_value:
                    column_indices[expected_column] = col
                    break
        
        logging.info(f"Columnas encontradas: {list(column_indices.keys())}")
        
        # ===== PROCESAR PERFIL OCUPACIONAL =====
        col_perfil_ocupacional = 'Perfil Ocupacional'
        perfil_encontrado = None
        necesita_guardar = False
        
        # Cargar mapeo
        mapeo_perfiles = cargar_mapeo_perfiles()
        
        if mapeo_perfiles:
            logging.info(f"Mapeo cargado con {len(mapeo_perfiles)} programas")
            print(f"✓ Mapeo de perfiles cargado")
            
            try:
                ficha_caracterizacion = obtener_nombre_ficha(read_sheet)
                if ficha_caracterizacion:
                    nombre_programa = extraer_nombre_ficha(ficha_caracterizacion)
                    logging.info(f"Programa detectado: {nombre_programa}")
                    print(f"✓ Programa: {nombre_programa}")
                    
                    if nombre_programa:
                        perfil_encontrado = buscar_perfil_ocupacional(nombre_programa, mapeo_perfiles)
                        if perfil_encontrado:
                            logging.info(f"Perfil encontrado: {perfil_encontrado}")
                            print(f"✓ Perfil encontrado: {perfil_encontrado}")
                            
                            # Agregar al DataFrame
                            if col_perfil_ocupacional not in df.columns:
                                df[col_perfil_ocupacional] = ''
                            df[col_perfil_ocupacional] = perfil_encontrado
                            
                            necesita_guardar = True
                        else:
                            logging.error(f"Perfil no encontrado: {nombre_programa}")
                            print(f"✗ Perfil no encontrado para: {nombre_programa}")
                            raise PerfilOcupacionalNoEncontrado(nombre_programa)
                            
            except PerfilOcupacionalNoEncontrado:
                raise
            except Exception as e:
                logging.error(f"Error procesando perfil: {e}")
                print(f"✗ Error: {e}")
        
        # ===== CREAR COPIA EDITABLE =====
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        
        # ===== VERIFICAR/CREAR COLUMNA =====
        if col_perfil_ocupacional not in column_indices:
            logging.info(f"Creando columna '{col_perfil_ocupacional}'...")
            print(f"ℹ Creando columna '{col_perfil_ocupacional}'")
            
            next_col_index = max(column_indices.values()) + 1 if column_indices else read_sheet.ncols
            column_indices[col_perfil_ocupacional] = next_col_index
            
            # Escribir encabezado
            sheet.write(header_row, next_col_index, col_perfil_ocupacional)
            
            # Escribir celdas vacías
            doc_col = column_indices.get('Número de Documento')
            if doc_col is not None:
                for row_idx in range(header_row + 1, read_sheet.nrows):
                    try:
                        doc_value = read_sheet.cell_value(row_idx, doc_col)
                        if doc_value:
                            sheet.write(row_idx, next_col_index, '')
                    except:
                        pass
            
            necesita_guardar = True
            
        elif col_perfil_ocupacional not in df.columns:
            df[col_perfil_ocupacional] = ''
        
        # ===== ESCRIBIR PERFIL EN EXCEL =====
        if perfil_encontrado and col_perfil_ocupacional in column_indices:
            col_perfil = column_indices[col_perfil_ocupacional]
            doc_col = column_indices.get('Número de Documento')
            
            filas_escritas = 0
            if doc_col is not None:
                for row_idx in range(header_row + 1, read_sheet.nrows):
                    try:
                        doc_value = read_sheet.cell_value(row_idx, doc_col)
                        if doc_value and str(doc_value).strip():
                            sheet.write(row_idx, col_perfil, perfil_encontrado)
                            filas_escritas += 1
                    except Exception as e:
                        logging.warning(f"Error escribiendo fila {row_idx}: {e}")
            
            print(f"✓ Perfil asignado a {filas_escritas} filas")
            logging.info(f"Perfil '{perfil_encontrado}' asignado a {filas_escritas} filas")
        
        # ===== GUARDAR UNA SOLA VEZ =====
        if necesita_guardar:
            try:
                # Guardar en archivo temporal primero
                temp_file = ruta_excel.replace('.xls', '_temp.xls')
                wb.save(temp_file)
                logging.info(f"Guardado temporal: {temp_file}")
                
                # Verificar que el temporal se creó correctamente
                if os.path.exists(temp_file):
                    temp_size = os.path.getsize(temp_file)
                    original_size = os.path.getsize(ruta_excel)
                    
                    # Verificar que el tamaño es razonable
                    if temp_size > 0 and temp_size >= original_size * 0.8:
                        # Intentar abrir el temporal para verificar integridad
                        try:
                            test_rb = xlrd.open_workbook(temp_file)
                            test_rb.release_resources()
                            
                            # Si llegamos aquí, el archivo es válido
                            import time
                            time.sleep(0.5)  # Esperar que se liberen recursos
                            
                            # Reemplazar original
                            shutil.move(temp_file, ruta_excel)
                            print(f"✓ Archivo guardado exitosamente")
                            logging.info("Archivo guardado y verificado")
                            
                        except Exception as e:
                            logging.error(f"Archivo temporal corrupto: {e}")
                            print(f"⚠ Error: Archivo temporal corrupto")
                            # Restaurar backup
                            shutil.copy2(backup_path, ruta_excel)
                            raise Exception("Archivo corrupto, backup restaurado")
                    else:
                        logging.error(f"Tamaño inválido: {temp_size} vs {original_size}")
                        raise Exception("Archivo temporal tiene tamaño inválido")
                else:
                    raise Exception("No se pudo crear archivo temporal")
                    
            except Exception as e:
                logging.error(f"ERROR AL GUARDAR: {e}")
                print(f"✗ Error guardando: {e}")
                
                # Restaurar backup
                if os.path.exists(backup_path):
                    shutil.copy2(backup_path, ruta_excel)
                    print(f"✓ Backup restaurado")
                raise

        # ===== LIMPIAR DATOS =====
        df = df.dropna(subset=['Número de Documento']).copy()
        
        missing_columns = [col for col in expected_columns if col not in column_indices]
        if missing_columns:
            logging.warning(f"Columnas faltantes: {missing_columns}")

        logging.info(f"Total de registros válidos: {len(df)}")
        print(f"Total de registros: {len(df)}")
        
        # ===== RECARGAR ARCHIVO ACTUALIZADO =====
        if necesita_guardar:
            rb = xlrd.open_workbook(ruta_excel, formatting_info=True)
            read_sheet = rb.sheet_by_index(0)
            wb = copy(rb)
            sheet = wb.get_sheet(0)
        
        # ===== RETORNAR =====
        return df, wb, sheet, read_sheet, column_indices, header_row, programa_sin_perfil
        
    except Exception as e:
        logging.error(f"Error configurando Excel: {e}")
        raise