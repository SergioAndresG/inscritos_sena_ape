import os
import logging
import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy

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
        
        # Leer el archivo Excel, saltando las primeras 4 filas y definiendo los encabezados
        # La fila 5 (índice 4) contiene los encabezados reales
        df_raw = pd.read_excel(ruta_excel, header=None, engine='xlrd')
        
        # Extraer la información de la ficha desde las primeras filas
        ficha_info = {
            'Ficha': df_raw.iloc[1, 1] if df_raw.shape[0] > 1 and df_raw.shape[1] > 1 else "N/A",
            'Estado': df_raw.iloc[2, 1] if df_raw.shape[0] > 2 and df_raw.shape[1] > 1 else "N/A",
            'Fecha': df_raw.iloc[3, 1] if df_raw.shape[0] > 3 and df_raw.shape[1] > 1 else "N/A"
        }
        
        logging.info(f"Información de ficha: {ficha_info}")
        print(f"Información de ficha: {ficha_info}")
        
        # Ahora leemos el archivo nuevamente pero estableciendo la fila 5 como encabezado
        df = pd.read_excel(ruta_excel, header=4, dtype={
            'Número de Documento': str,
            'Celular': str
        }, engine='xlrd')
        
        # Verificar que se hayan cargado correctamente las columnas
        expected_columns = ['Tipo de Documento', 'Número de Documento', 'Nombre', 
                        'Apellidos', 'Celular', 'Correo Electrónico', 'Estado', 'Perfil Ocupacional']
        
        # Comprobar si existen las columnas esperadas (verificando parcialmente)
        missing_columns = [col for col in expected_columns if not any(existing_col.startswith(col) for existing_col in df.columns)]
        
        if missing_columns:
            logging.warning(f"Advertencia: No se encontraron algunas columnas esperadas: {missing_columns}")
            logging.warning(f"Columnas disponibles: {df.columns.tolist()}")
            print(f"Advertencia: No se encontraron algunas columnas esperadas: {missing_columns}")
            print(f"Columnas disponibles: {df.columns.tolist()}")
        
        # Eliminar filas con valores NaN en la columna de documento
        df = df.dropna(subset=['Número de Documento']).copy()

        # Registrar información del archivo
        logging.info(f"Archivo Excel cargado correctamente: {ruta_excel}")
        logging.info(f"Total de registros: {len(df)}")
        
        # ----- PREPARAR EL ARCHIVO PARA COLOREAR CELDAS -----
        # Cargar el libro de trabajo con xlrd para leer (necesario para formato)
        rb = xlrd.open_workbook(ruta_excel, formatting_info=True)
        # Hacer una copia editable
        wb = copy(rb)
        sheet = wb.get_sheet(0)  # Obtener la primera hoja
        
        # Definir estilos de colores para .xls
        style_procesando = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue')
        style_procesado = xlwt.easyxf('pattern: pattern solid, fore_colour light_green')
        style_ya_existe = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow')
        style_error = xlwt.easyxf('pattern: pattern solid, fore_colour red')
        
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
        
        print(f"Índices de columnas encontrados: {column_indices}")
        logging.info(f"Índices de columnas encontrados: {column_indices}")
        
        # Registrar información del archivo
        logging.info(f"Archivo Excel cargado correctamente: {ruta_excel}")
        logging.info(f"Total de registros: {len(df)}")
        
        # Mostrar las primeras filas para verificar la estructura
        print("Estructura del archivo Excel:")
        print(df.head())
        
        # Devolver todos los objetos necesarios para el script principal
        return df, wb, sheet, read_sheet, column_indices, header_row
        
    except Exception as e:
        logging.error(f"Error configurando el excel: {e}")
        raise e # Re-lanzar la excepción original
    