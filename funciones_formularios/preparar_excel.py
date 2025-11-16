import os
import logging
import pandas as pd
import xlrd
from xlutils.copy import copy
from perfilesOcupacionales.gestorDePerfilesOcupacionales import extraer_nombre_ficha, cargar_mapeo_perfiles, buscar_perfil_ocupacional, obtener_nombre_ficha

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
    # --- Cargar el archivo Excel con pandas y preparar para colorear celdas ---
    try:
        if not os.path.exists(ruta_excel):
            raise FileNotFoundError(f"El archivo no se encuentra en la ruta especificada: {ruta_excel}")

        # Configurar pandas para leer correctamente los números de documento
        pd.set_option('display.float_format', lambda x: '%.0f' % x)
        
        # Ahora leemos el archivo
        df = pd.read_excel(ruta_excel, header=4, dtype={
            'Número de Documento': str,
            'Celular': str
        }, engine='xlrd')
        
        # Verificar que se hayan cargado correctamente las columnas
        expected_columns = ['Tipo de Documento', 'Número de Documento', 'Nombre', 
                        'Apellidos', 'Celular', 'Correo Electrónico', 'Estado', 'Perfil Ocupacional']
        
        # Cargar el libro de trabajo con xlrd para leer (necesario para formato)
        rb = xlrd.open_workbook(ruta_excel, formatting_info=True)
        # Hacer una copia editable
        wb = copy(rb)
        sheet = wb.get_sheet(0)  # Obtener la primera hoja
        
        # Obtener la hoja de lectura
        read_sheet = rb.sheet_by_index(0)
        
        # Encontrar los índices de las columnas en Excel
        header_row = 4  # Fila 5 (índice 4) contiene los encabezados
        column_indices = {}
        
        # Buscar las columnas por su nombre
        for col in range(read_sheet.ncols):
            cell_value = read_sheet.cell_value(header_row, col)
            for expected_column in expected_columns:
                if cell_value and expected_column in cell_value:
                    column_indices[expected_column] = col
                    break
        
        # Se crea la Columna de perfil ocupacional SI NO EXISTE EN EL ARCHIVO
        col_perfil_ocupacional = 'Perfil Ocupacional'
        if col_perfil_ocupacional not in column_indices:
            logging.info(f"La columna '{col_perfil_ocupacional}' no existe en el archivo. Creándola...")
            
            # Encontrar la siguiente columna disponible
            next_col_index = max(column_indices.values()) + 1 if column_indices else read_sheet.ncols
            
            # Escribir el encabezado en el Excel
            sheet.write(header_row, next_col_index, col_perfil_ocupacional)
            
            # Actualizar el diccionario de índices
            column_indices[col_perfil_ocupacional] = next_col_index
            
            # Escribir celdas vacías en todas las filas de datos
            # Contar las filas con datos (excluyendo las primeras 5 filas de encabezados)
            for row_idx in range(header_row + 1, read_sheet.nrows):
                # Verificar si la fila tiene datos (revisando si tiene número de documento)
                if 'Número de Documento' in column_indices:
                    doc_col = column_indices['Número de Documento']
                    try:
                        doc_value = read_sheet.cell_value(row_idx, doc_col)
                        if doc_value:  # Si hay un documento, es una fila válida
                            sheet.write(row_idx, next_col_index, '')
                    except:
                        pass
            
            # Guardar el archivo con la nueva columna
            temp_file = ruta_excel.replace('.xls', '_temp.xls')
            wb.save(temp_file)
            
            # Reemplazar el archivo original
            import shutil
            shutil.move(temp_file, ruta_excel)
            
            # Volver a cargar el archivo actualizado
            rb = xlrd.open_workbook(ruta_excel, formatting_info=True)
            wb = copy(rb)
            sheet = wb.get_sheet(0)
            read_sheet = rb.sheet_by_index(0)
            
            # También crear la columna en el DataFrame
            df[col_perfil_ocupacional] = ''
            
            logging.info(f"Columna '{col_perfil_ocupacional}' creada exitosamente en el archivo Excel")
        else:
            # Si la columna ya existe en el Excel, asegurarse de que también esté en el DataFrame
            if col_perfil_ocupacional not in df.columns:
                df[col_perfil_ocupacional] = ''

        # Bloque de codigo para ingresar el perfil ocupacional
        mapeo_perfiles = cargar_mapeo_perfiles()
                
        if mapeo_perfiles:
            logging.info("Mapeo de perfiles cargado. Rellenando columna")

            # Obtener el nombre del programa desde la celda de "Ficha de Caracterización"
            # Esta info esta en la fila 1, despues de "Ficha de Caracterizacion"
            try:
                ficha_caracterización = obtener_nombre_ficha(read_sheet)
                if ficha_caracterización:
                    nombre_programa = extraer_nombre_ficha(ficha_caracterización)
                
                    if nombre_programa:
                        perfil = buscar_perfil_ocupacional(nombre_programa, mapeo_perfiles)

                        if perfil:
                            logging.info(f"Progrma: {nombre_programa} -> Perfil: {perfil}")

                            # LLenar la columna de perfil ocupacional en el DataFrame
                            df['Perfil Ocupacional'] = perfil

                            # Tambien escribir en el Excel 
                            col_perfil = column_indices['Perfil Ocupacional']
                            for row_idx in range(header_row + 1, read_sheet.nrows):
                                # Solo escirbir en filas que tengan documeto
                                if 'Número de Documento' in column_indices:
                                    doc_col = column_indices['Número de Documento']
                                    try:
                                        doc_value = read_sheet.cell_value(row_idx, doc_col)
                                        if doc_value:
                                            sheet.write(row_idx, col_perfil, perfil)
                                    except:
                                        pass
                            # Guardar cambios en el Excel
                            wb.save(ruta_excel)
                            
                            logging.info(f"Perfil ocupacional '{perfil}' asignado a todos los aprendices")
                        else:
                            logging.warning(f"No se encontró perfil para el programa: {nombre_programa}")
                            logging.warning("Se debe asignar manualmente el perfil ocupacional")
                    else:
                        logging.warning("No se pudo extraer el nombre del programa")
                    
            except Exception as e:
                logging.error(f"Error al procesar perfil ocupacional: {e}")
            wb.save(ruta_excel)                

        # Comprobar si existen las columnas esperadas
        missing_columns = [col for col in expected_columns if col not in column_indices]
        
        if missing_columns:
            logging.warning(f"Advertencia: No se encontraron algunas columnas esperadas: {missing_columns}")
            logging.warning(f"Columnas disponibles en Excel: {list(column_indices.keys())}")
        
        # Eliminar filas con valores NaN en la columna de documento
        df = df.dropna(subset=['Número de Documento']).copy()
                
        # Registrar información del archivo
        logging.info(f"Archivo Excel cargado correctamente: {ruta_excel}")
        logging.info(f"Total de registros: {len(df)}")
        logging.info(f"Columnas en el archivo Excel: {list(column_indices.keys())}")
        
        # Devolver todos los objetos necesarios para el script principal
        return df, wb, sheet, read_sheet, column_indices, header_row
        
    except Exception as e:
        logging.error(f"Error configurando el excel: {e}")
        raise e
    