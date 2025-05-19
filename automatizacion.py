import os
import time
import logging
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import traceback
from openpyxl import load_workbook
import xlrd
import xlwt
from xlutils.copy import copy

from functions.login import login
from functions.campo_estrato import llenar_estrato
from functions.campo_sueldo import llenar_formulario_sueldo
from functions.campo_telefono_correo import llenar_formulario_telefono_correo
from functions.experincia_laboral_campos import experiencia_laboral
from functions.form_campo_estado_civil import llenar_formulario_estado_civil
from functions.form_campo_perfil_ocupacional import llenar_input_perfil_ocupacional
from functions.form_campos_nacimiento import llenar_formulario_ubicaciones_nacimiento
from functions.form_campos_ubicacion_identificacion import llenar_formulario_ubicaciones
from functions.pre_inscripcion import llenar_datos_antes_de_inscripcion
from functions.verificacion import verificar_estudiante_con_CC_primero, verificar_estudiante
from functions.form_datos_residencia import llenar_formulario_ubicacion_residencia
from URLS.urls import URL_FORMULARIO, URL_LOGIN, URL_VERIFICACION

# Configurar logging
log_filename = f"automatizacion_aprendices_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Cargar variables de entorno
load_dotenv()

RUTA_EXCEL = 'C:/Users/SENA/Downloads/Reporte de Aprendices Ficha 3204166.xls'



# Mapeo de tipos de documento
TIPOS_DOCUMENTO = {
    "CC": "1",
    "TI": "2", 
    "CE": "3",
    "Otro Documento de Identidad": "5",
    "PEP": "8",
    "PPT": "9"
}


chrome_options = Options()

# Descomentar la siguiente línea para modo headless (sin interfaz gráfica), para visualizar el funcionamiento del aplicativo
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)  # Espera explícita de 10 segundos


# --- Cargar el archivo Excel con pandas y preparar para colorear celdas ---
try:
    # Configurar pandas para leer correctamente los números de documento
    pd.set_option('display.float_format', lambda x: '%.0f' % x)
    
    # Ruta del archivo Excel
    print(f"RUTA_EXCEL --> {RUTA_EXCEL}")
    
    # Leer el archivo Excel, saltando las primeras 4 filas y definiendo los encabezados
    # La fila 5 (índice 4) contiene los encabezados reales
    df_raw = pd.read_excel(RUTA_EXCEL, header=None, engine='xlrd')
    
    # Extraer la información de la ficha desde las primeras filas
    ficha_info = {
        'Ficha': df_raw.iloc[1, 1] if df_raw.shape[0] > 1 and df_raw.shape[1] > 1 else "N/A",
        'Estado': df_raw.iloc[2, 1] if df_raw.shape[0] > 2 and df_raw.shape[1] > 1 else "N/A",
        'Fecha': df_raw.iloc[3, 1] if df_raw.shape[0] > 3 and df_raw.shape[1] > 1 else "N/A"
    }
    
    logging.info(f"Información de ficha: {ficha_info}")
    print(f"Información de ficha: {ficha_info}")
    
    # Ahora leemos el archivo nuevamente pero estableciendo la fila 5 como encabezado
    df = pd.read_excel(RUTA_EXCEL, header=4, dtype={
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
    logging.info(f"Archivo Excel cargado correctamente: {RUTA_EXCEL}")
    logging.info(f"Total de registros: {len(df)}")
    
    # ----- PREPARAR EL ARCHIVO PARA COLOREAR CELDAS -----
    # Cargar el libro de trabajo con xlrd para leer (necesario para formato)
    rb = xlrd.open_workbook(RUTA_EXCEL, formatting_info=True)
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
    logging.info(f"Archivo Excel cargado correctamente: {RUTA_EXCEL}")
    logging.info(f"Total de registros: {len(df)}")
    
    # Mostrar las primeras filas para verificar la estructura
    print("Estructura del archivo Excel:")
    print(df.head())
    
except FileNotFoundError:
    error_msg = f"Error: No se encontró el archivo Excel en la ruta: {RUTA_EXCEL}"
    logging.error(error_msg)
    print(error_msg)
    exit()
except Exception as e:
    error_msg = f"Error al abrir el archivo Excel: {str(e)}"
    logging.error(error_msg)
    print(error_msg)
    exit()




def main():

    try:
        # Realizar login
        if not login(driver):
            logging.error("No se pudo completar el login. Abortando proceso.")
            return
            
        # Establecer los nombres de columnas según tu archivo Excel
        COLUMNA_TIPO_DOC = 'Tipo de Documento'
        COLUMNA_NUM_DOC = 'Número de Documento'
        COLUMNA_NOMBRES = 'Nombre'
        COLUMNA_APELLIDOS = 'Apellidos'
        COLUMNA_CELULAR = 'Celular'
        COLUMNA_CORREO = 'Correo Electrónico'
        COLUMNA_ESTADO = 'Estado'
        COLUMNA_PERFIL = 'Perfil Ocupacional'
        
        # Procesar cada registro en el DataFrame de pandas
        total_registros = len(df)
        for i, fila in df.iterrows():
            # El índice real en Excel es el índice en pandas + 6 (header_row + 2)
            excel_row = i + header_row + 1
            
            # Inicializar variables para evitar errores en bloques except
            nombres = "Sin nombre"
            apellidos = "Sin apellido"
            tipo_doc = ""
            num_doc = ""
            celular = ""
            correo = ""
            estado = ""
            perfil_ocupacional = ""
            
            try:
                # Colorear la fila actual como "procesando"
                for col_name, col_idx in column_indices.items():
                    try:
                        sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_procesando)
                    except Exception as e:
                        print(f"Error al colorear celda {col_name}: {str(e)}")
                
                # Guardar los cambios para que sean visibles inmediatamente
                try:
                    wb.save(RUTA_EXCEL)
                    print(f"Excel actualizado: marcando fila {excel_row + 1} como 'procesando'")
                except Exception as e:
                    print(f"Error al guardar Excel: {str(e)}")
                
                # Extraer datos del estudiante
                tipo_doc = str(fila[COLUMNA_TIPO_DOC])
                num_doc = str(fila[COLUMNA_NUM_DOC])
                nombres = str(fila[COLUMNA_NOMBRES])
                apellidos = str(fila[COLUMNA_APELLIDOS])
                celular = str(fila[COLUMNA_CELULAR])
                correo = str(fila[COLUMNA_CORREO])
                estado = str(fila[COLUMNA_ESTADO])
                perfil_ocupacional = str(fila[COLUMNA_PERFIL])
                
                logging.info(f"Procesando estudiante {i+1}/{total_registros}: {nombres} {apellidos}")
                print(f"\n===== Procesando estudiante {i+1}/{total_registros}: {nombres} {apellidos} =====\n")
                
                # Verificar si el estudiante ya existe
                existe = verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos, driver, wait)
                
                # Si existe es None, hubo error en la verificación
                if existe is None:
                    logging.warning(f"Saltando estudiante debido a error en verificación: {nombres} {apellidos}")
                    print(f"⚠️ Saltando estudiante debido a error en verificación: {nombres} {apellidos}")
                    
                    # Colorear fila como error
                    for col_name, col_idx in column_indices.items():
                        try:
                            sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                        except Exception as e:
                            print(f"Error al colorear celda {col_name}: {str(e)}")
                    wb.save(RUTA_EXCEL)
                    continue
                    
                # Si el estudiante ya existe, pasar al siguiente
                if existe:
                    logging.info(f"El estudiante {nombres} {apellidos} ya existe en el sistema. Pasando al siguiente.")
                    print(f"✅ El estudiante {nombres} {apellidos} ya existe en el sistema. Pasando al siguiente.")
                    
                    # Colorear fila como ya existente
                    for col_name, col_idx in column_indices.items():
                        try:
                            sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_ya_existe)
                        except Exception as e:
                            print(f"Error al colorear celda {col_name}: {str(e)}")
                    wb.save(RUTA_EXCEL)
                    continue
                
                # Si llegamos aquí, el estudiante no existe
                logging.info(f"El estudiante {nombres} {apellidos} no existe. Procediendo con el registro.")
                print(f"📝 El estudiante {nombres} {apellidos} no existe. Procediendo con el registro.")
                
                # Verificar si ya estamos en el formulario de pre-inscripción
                if URL_VERIFICACION in driver.current_url:
                    print("Detectado formulario de pre-inscripción")
                    # Llenar los datos iniciales
                    if llenar_datos_antes_de_inscripcion(nombres, apellidos, driver):
                        print("Pre-inscripción completada. Esperando formulario completo...")
                        # Esperar a que cargue el formulario completo
                        time.sleep(3)
                        # Verificar si estamos en la página del formulario completo
                        if driver.current_url == URL_FORMULARIO or "formulario" in driver.current_url.lower():
                            print("Formulario completo detectado. Procediendo a llenar...")
                            llenar_formulario_ubicaciones(driver)
                            llenar_formulario_ubicaciones_nacimiento(driver)
                            llenar_formulario_ubicacion_residencia(driver)
                            llenar_formulario_estado_civil(driver)
                            llenar_formulario_sueldo(driver)
                            llenar_estrato(driver)
                            llenar_formulario_telefono_correo(celular, correo, driver)
                            llenar_input_perfil_ocupacional(estado,driver)
                            print("✅ Formulario con datos basicos llenado correcctamente")
                            #boton de guardar inforamacion
                            botonGuardar = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, 'submitNewOft'))
                            )
                            botonGuardar.click()
                            print("✅ Se hizo click en el boton de Guardar Correctamente")
                            print("Esperando respuesta")
                            logging.info("Se hizo click en el boton de Guardar")
                            time.sleep(10)
                            experiencia_laboral(driver,  perfil_ocupacional)
                            time.sleep(10)
                            
                            # Colorear fila como procesado exitosamente
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_procesado)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                            wb.save(RUTA_EXCEL)
                            print(f"Excel actualizado: marcando fila {excel_row + 1} como 'procesado'")

                        else:
                            print(f"⚠️ No se detectó redirección al formulario completo. URL actual: {driver.current_url}")
                            # Colorear fila como error
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                            wb.save(RUTA_EXCEL)
                    else:
                        print("❌ No se pudo completar la pre-inscripción")
                        # Colorear fila como error
                        for col_name, col_idx in column_indices.items():
                            try:
                                sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                            except Exception as e:
                                print(f"Error al colorear celda {col_name}: {str(e)}")
                        wb.save(RUTA_EXCEL)
                else:
                    print(f"⚠️ No se redirigió al formulario de pre-inscripción. URL actual: {driver.current_url}")
                    # Intentar redirigir manualmente
                    driver.get(URL_VERIFICACION)
                    print("Reintentando verificación...")
                    existe = verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait) 
                    if not existe:
                        print("Reintentando llenar datos...")
                        if llenar_datos_antes_de_inscripcion(nombres, apellidos,driver,wait):
                            # Si la reinscripción funciona, colorear como procesado
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_procesado)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                        else:
                            # Si falla, colorear como error
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                        wb.save(RUTA_EXCEL)
                    
            except Exception as e:
                logging.error(f"Error procesando estudiante {nombres} {apellidos}: {str(e)}")
                print(f"❌ Error procesando estudiante {nombres} {apellidos}: {str(e)}")
                
                # Colorear fila como error
                for col_name, col_idx in column_indices.items():
                    try:
                        sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                    except Exception as write_error:
                        print(f"Error al colorear celda {col_name}: {str(write_error)}")
                try:
                    wb.save(RUTA_EXCEL)
                    print(f"Excel actualizado: marcando fila {excel_row + 1} como 'error'")
                except Exception as save_error:
                    print(f"Error al guardar Excel: {str(save_error)}")
                
                # Volver a la página de verificación para el siguiente estudiante
                try:
                    driver.get(URL_VERIFICACION)
                    time.sleep(2)
                except Exception as nav_error:
                    print(f"Error al navegar: {str(nav_error)}")
                
        logging.info("✅ Proceso completado exitosamente")
        print("\n===== ✅ Proceso completado exitosamente =====\n")
            
    except Exception as e:
        logging.error(f"Error general en el proceso: {str(e)}")
        print(f"❌ Error general: {str(e)}")
        
    finally:
        # Cerrar el navegador al finalizar
        driver.quit()
        logging.info("Navegador cerrado")
        print("🔒 Navegador cerrado")

if __name__ == "__main__":
        main()


