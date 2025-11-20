import os
import logging
import pandas as pd
from openpyxl import load_workbook
from perfilesOcupacionales.perfilExcepcion import PerfilOcupacionalNoEncontrado
from perfilesOcupacionales.gestorDePerfilesOcupacionales import (
    extraer_nombre_ficha, 
    cargar_mapeo_perfiles, 
    buscar_perfil_ocupacional, 
    obtener_nombre_ficha
)
from funciones_excel.conversion_excel import convertir_xls_a_xlsx

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
    Prepara el archivo Excel para procesamiento
    Convierte .xls a .xlsx automáticamente
    """
    programa_sin_perfil = None
    
    try:
        # ===== VALIDAR Y CONVERTIR =====
        if not os.path.exists(ruta_excel):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_excel}")
        
        logging.info(f"Abriendo archivo: {ruta_excel}")
        
        # Convertir si es .xls
        if ruta_excel.endswith('.xls'):
            ruta_excel = convertir_xls_a_xlsx(ruta_excel)
        
        # ===== CONFIGURAR PANDAS =====
        pd.set_option('display.float_format', lambda x: '%.0f' % x)
        
        # ===== LEER CON PANDAS =====
        df = pd.read_excel(
            ruta_excel, 
            header=4,
            dtype={'Número de Documento': str, 'Celular': str}, 
            engine='openpyxl'
        )
        
        # ===== CARGAR CON OPENPYXL =====
        wb = load_workbook(ruta_excel)
        sheet = wb.active
        
        # ===== MAPEAR COLUMNAS =====
        expected_columns = [
            'Tipo de Documento', 'Número de Documento', 'Nombre', 
            'Apellidos', 'Celular', 'Correo Electrónico', 
            'Estado', 'Perfil Ocupacional'
        ]
        
        header_row = 4  # Fila 5 en Excel (índice 4 en pandas, fila 5 en openpyxl)
        column_indices = {}
        
        # Obtener encabezados (fila 5 en openpyxl es índice 5)
        for col_idx, cell in enumerate(list(sheet.rows)[header_row], start=0):
            cell_value = cell.value
            for expected_column in expected_columns:
                if cell_value and expected_column in str(cell_value):
                    column_indices[expected_column] = col_idx
                    break
        
        logging.info(f"Columnas encontradas: {list(column_indices.keys())}")
        
        # ===== PROCESAR PERFIL OCUPACIONAL =====
        col_perfil_ocupacional = 'Perfil Ocupacional'
        perfil_encontrado = None
        necesita_guardar = False
        
        mapeo_perfiles = cargar_mapeo_perfiles()
        
        if mapeo_perfiles:
            logging.info(f"Mapeo cargado con {len(mapeo_perfiles)} programas")
            print(f"✓ Mapeo de perfiles cargado")
            
            try:
                # Obtener nombre del programa (adaptado para openpyxl)
                ficha_caracterizacion = None
                # Buscar en las primeras filas la ficha de caracterización
                for row in list(sheet.rows)[:5]:
                    for cell in row:
                        if cell.value and 'Ficha de Caracterización' in str(cell.value):
                            # El valor está en la celda siguiente
                            ficha_caracterizacion = row[cell.column - 1 + 1].value
                            break
                    if ficha_caracterizacion:
                        break
                
                if ficha_caracterizacion:
                    nombre_programa = extraer_nombre_ficha(str(ficha_caracterizacion))
                    logging.info(f"Programa detectado: {nombre_programa}")
                    print(f"✓ Programa: {nombre_programa}")
                    
                    if nombre_programa:
                        perfil_encontrado = buscar_perfil_ocupacional(nombre_programa, mapeo_perfiles)
                        if perfil_encontrado:
                            logging.info(f"Perfil encontrado: {perfil_encontrado}")
                            print(f"✓ Perfil encontrado: {perfil_encontrado}")
                            
                            if col_perfil_ocupacional not in df.columns:
                                df[col_perfil_ocupacional] = ''
                            df[col_perfil_ocupacional] = perfil_encontrado
                            
                            necesita_guardar = True
                        else:
                            programa_sin_perfil = nombre_programa
                            logging.error(f"Perfil no encontrado: {nombre_programa}")
                            print(f"✗ Perfil no encontrado para: {nombre_programa}")
                            raise PerfilOcupacionalNoEncontrado(nombre_programa)
                            
            except PerfilOcupacionalNoEncontrado:
                raise
            except Exception as e:
                logging.error(f"Error procesando perfil: {e}")
                print(f"✗ Error: {e}")
                import traceback
                traceback.print_exc()
        
        # ===== VERIFICAR/CREAR COLUMNA =====
        if col_perfil_ocupacional not in column_indices:
            logging.info(f"Creando columna '{col_perfil_ocupacional}'...")
            print(f"ℹ Creando columna '{col_perfil_ocupacional}'")
            
            next_col_index = max(column_indices.values()) + 1 if column_indices else len(column_indices)
            column_indices[col_perfil_ocupacional] = next_col_index
            
            # Escribir encabezado (fila 5 en Excel = índice 5 en openpyxl)
            sheet.cell(row=header_row + 1, column=next_col_index + 1, value=col_perfil_ocupacional)
            
            # Escribir celdas vacías en filas de datos
            doc_col_idx = column_indices.get('Número de Documento')
            if doc_col_idx is not None:
                for row_idx in range(header_row + 2, sheet.max_row + 1):
                    doc_cell = sheet.cell(row=row_idx, column=doc_col_idx + 1)
                    if doc_cell.value:
                        sheet.cell(row=row_idx, column=next_col_index + 1, value='')
            
            necesita_guardar = True
            
        elif col_perfil_ocupacional not in df.columns:
            df[col_perfil_ocupacional] = ''
        
        # ===== ESCRIBIR PERFIL EN EXCEL =====
        if perfil_encontrado and col_perfil_ocupacional in column_indices:
            col_perfil_idx = column_indices[col_perfil_ocupacional]
            doc_col_idx = column_indices.get('Número de Documento')
            
            filas_escritas = 0
            if doc_col_idx is not None:
                # Iterar desde fila 6 (header_row + 2) en adelante
                for row_idx in range(header_row + 2, sheet.max_row + 1):
                    doc_cell = sheet.cell(row=row_idx, column=doc_col_idx + 1)
                    if doc_cell.value and str(doc_cell.value).strip():
                        sheet.cell(row=row_idx, column=col_perfil_idx + 1, value=perfil_encontrado)
                        filas_escritas += 1
            
            print(f"✓ Perfil asignado a {filas_escritas} filas")
            logging.info(f"Perfil '{perfil_encontrado}' asignado a {filas_escritas} filas")
        
        # ===== GUARDAR =====
        if necesita_guardar:
            try:
                wb.save(ruta_excel)
                print(f"✓ Archivo guardado: {os.path.basename(ruta_excel)}")
                logging.info(f"Archivo guardado: {ruta_excel}")
                
                # Verificar
                test_df = pd.read_excel(ruta_excel, header=4)
                if col_perfil_ocupacional in test_df.columns:
                    print(f"✓ Columna '{col_perfil_ocupacional}' verificada")
                
            except Exception as e:
                logging.error(f"ERROR AL GUARDAR: {e}")
                print(f"✗ Error guardando: {e}")
                raise
        
        # ===== LIMPIAR DATOS =====
        df = df.dropna(subset=['Número de Documento']).copy()
        
        missing_columns = [col for col in expected_columns if col not in column_indices]
        if missing_columns:
            logging.warning(f"Columnas faltantes: {missing_columns}")
        
        logging.info(f"Total de registros válidos: {len(df)}")
        print(f"Total de registros: {len(df)}")
        
        # ===== RETORNAR =====
        return df, wb, sheet, sheet, column_indices, header_row, programa_sin_perfil, ruta_excel
        #                                                                            ^ NUEVA RUTA
        
    except Exception as e:
        logging.error(f"Error configurando Excel: {e}")
        raise