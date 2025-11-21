import logging
from perfilesOcupacionales.gestorDePerfilesOcupacionales import extraer_nombre_ficha

def extraer_info_antes_conversion(ruta_xls):
    """
    Extrae información crítica del archivo .xls ANTES de convertirlo
    Retorna: (nombre_ficha, nombre_programa)
    """
    try:
        import xlrd
        
        logging.info(f"Leyendo información del archivo .xls original: {ruta_xls}")
        print(f"📖 Extrayendo información del .xls original...")
        
        # Leer con xlrd (soporta .xls con formato)
        rb = xlrd.open_workbook(ruta_xls, formatting_info=False)
        sheet = rb.sheet_by_index(0)
        
        # Buscar "Ficha de Caracterización" en las primeras filas
        nombre_ficha = None
        for row_idx in range(min(5, sheet.nrows)):  # Buscar en primeras 5 filas
            for col_idx in range(min(10, sheet.ncols)):  # Buscar en primeras 10 columnas
                try:
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    
                    # Si encontramos "Ficha de Caracterización"
                    if cell_value and 'Ficha de Caracterización' in str(cell_value):
                        # Buscar el valor en las celdas siguientes de la misma fila
                        for next_col in range(col_idx + 1, min(col_idx + 5, sheet.ncols)):
                            next_value = sheet.cell_value(row_idx, next_col)
                            if next_value and str(next_value).strip():
                                nombre_ficha = str(next_value).strip()
                                logging.info(f"Ficha encontrada: {nombre_ficha}")
                                print(f"✓ Ficha encontrada: {nombre_ficha}")
                                break
                        
                        if nombre_ficha:
                            break
                except:
                    continue
            
            if nombre_ficha:
                break
        
        if not nombre_ficha:
            logging.warning("No se encontró 'Ficha de Caracterización' en el archivo")
            print("⚠️ No se encontró 'Ficha de Caracterización'")
            return None, None
        
        # Extraer nombre del programa
        nombre_programa = extraer_nombre_ficha(nombre_ficha)
        
        if nombre_programa:
            logging.info(f"Programa extraído: {nombre_programa}")
            print(f"✓ Programa: {nombre_programa}")
        else:
            logging.warning(f"No se pudo extraer programa de: {nombre_ficha}")
            print(f"⚠️ No se pudo extraer programa")
        
        rb.release_resources()
        return nombre_ficha, nombre_programa
        
    except Exception as e:
        logging.error(f"Error extrayendo info del .xls: {e}")
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None
