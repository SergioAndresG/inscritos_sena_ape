import time
import os
import sys
import os
import logging
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import xlwt
from funciones_formularios.preparar_excel import preparar_excel
from funciones_formularios.login import login
from funciones_formularios.campo_estrato import llenar_estrato
from funciones_formularios.campo_sueldo import llenar_formulario_sueldo
from funciones_formularios.campo_telefono_correo import llenar_formulario_telefono_correo
from funciones_formularios.experincia_laboral_campos import experiencia_laboral
from funciones_formularios.form_campo_estado_civil import llenar_formulario_estado_civil
from funciones_formularios.form_campo_perfil_ocupacional import llenar_input_perfil_ocupacional
from funciones_formularios.form_campos_nacimiento import llenar_formulario_ubicaciones_nacimiento
from funciones_formularios.form_campos_ubicacion_identificacion import llenar_formulario_ubicaciones
from funciones_formularios.pre_inscripcion import llenar_datos_antes_de_inscripcion
from funciones_formularios.verificacion import verificar_estudiante_con_CC_primero, verificar_estudiante
from funciones_formularios.form_datos_residencia import llenar_formulario_ubicacion_residencia
from funciones_formularios.meses_busqueda import verificar_meses_busqueda
from funciones_loggs.loggs_funciones import loggs
from URLS.urls import URL_FORMULARIO, URL_VERIFICACION


#Funcion que preapara los logs en un archivo y maneja su durabilidad en la aplicación
loggs()

# Cargar variables de entorno
load_dotenv()

# Mapeo de tipos de documento
TIPOS_DOCUMENTO = {
    "CC": "1",
    "TI": "2", 
    "CE": "3",
    "Otro Documento de Identidad": "5",
    "PEP": "8",
    "PPT": "9"
}

# --- Variables Globales ---
# Se definirán dentro de la función main para que sean accesibles en todo el script
RUTA_EXCEL = ""

chrome_options = Options()

# Descomentar la siguiente línea para modo headless (sin interfaz gráfica), para visualizar el funcionamiento del aplicativo
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")


service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 10)  # Espera explícita de 10 segundos
wait_rapido = WebDriverWait(driver, 3) # Espera mas corta

# Definir estilos de colores para .xls
style_procesando = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue')
style_procesado = xlwt.easyxf('pattern: pattern solid, fore_colour light_green')
style_ya_existe = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow')
style_error = xlwt.easyxf('pattern: pattern solid, fore_colour red')

class QueueStream:
    """Un objeto tipo archivo que escribe en una cola de progreso."""
    def __init__(self, queue):
        self.queue = queue

    def write(self, text):
        # Envía el texto a la cola para que la GUI lo muestre.
        if self.queue:
            self.queue.put(("log", text))

    def flush(self):
        # Necesario para la interfaz de archivo, pero no hace nada aquí.
        pass


def main(ruta_excel_param, progress_queue=None, username=None, password=None, stop_event=threading.Event):
    # Hacemos globales las variables que se usarán en todo el script
    global RUTA_EXCEL, df, wb, sheet, read_sheet, column_indices, header_row, programa_sin_perfil
    RUTA_EXCEL = ruta_excel_param
    original_stdout = sys.stdout  # Guardar la salida estándar original
    if progress_queue:
        sys.stdout = QueueStream(progress_queue) # Redirigir print() a la GUI

    try:
        # Llamamos a la nueva función para preparar el Excel
        df, wb, sheet, read_sheet, column_indices, header_row, programa_sin_perfil = preparar_excel(RUTA_EXCEL)
        logging.info(f"Archivo Excel '{RUTA_EXCEL}' cargado y listo para procesar.")
        print(f" Archivo Excel '{os.path.basename(RUTA_EXCEL)}' cargado y listo.")
    except (FileNotFoundError, Exception) as e:
        logging.error(f"Error fatal al preparar el archivo Excel: {e}")
        print(f" Error fatal al preparar el archivo Excel: {e}")
        return # Salir de la función main si no se puede cargar el Excel

    try:
        # Realizar login
        if not login(driver, username, password):
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
        
        # Contadores para estadísticas
        contador_procesados_exitosamente = 0
        contador_ya_existentes = 0
        contador_errores = 0
        contador_saltados = 0
        
        # Procesar cada registro en el DataFrame de pandas
        total_registros = len(df)
        for i, fila in df.iterrows():
            if stop_event.is_set():
                progress_queue.put(("log", f"Proceso detenido por el usuario en Tarea {i-1}.\n"))
                return # Salimos de la función limpiamente
            # El índice real en Excel es el índice en pandas + 6 (header_row + 2)
            excel_row = i + header_row + 1
            
            # Enviar progreso a la GUI si la cola está disponible
            if progress_queue:
                progress_queue.put(("progress", (i + 1, total_registros)))

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
                        # leer del dataframe si la columna existe, sino del readsheet
                        if col_name in df.columns:
                            valor = fila[col_name]
                        else:
                            valor = read_sheet.cell_value(excel_row, col_idx)
                        sheet.write(excel_row, col_idx, valor, style_procesando)
                    except:
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
                print( f"\n===== Procesando estudiante {i+1}/{total_registros}: {nombres} {apellidos} =====\n")
                
                # Verificar si el estudiante ya existe
                existe = verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido)
                
                # Si existe es None, hubo error en la verificación
                if existe is None:
                    logging.warning(f"Saltando estudiante debido a error en verificación: {nombres} {apellidos}")
                    print(f"⚠️ Saltando estudiante debido a error en verificación: {nombres} {apellidos}")
                    
                    # Colorear fila como error
                    for col_name, col_idx in column_indices.items():
                        try:
                            # leer del dataframe si la columna existe, sino del readsheet
                            if col_name in df.columns:
                                valor = fila[col_name]
                            else:
                                valor = read_sheet.cell_value(excel_row, col_idx)
                            sheet.write(excel_row, col_idx, valor, style_error)
                        except:
                            print(f"Error al colorear celda {col_name}: {str(e)}")  
                    wb.save(RUTA_EXCEL)
                    contador_saltados += 1
                    continue
                    
                # Si el estudiante ya existe, pasar al siguiente
                if existe:
                    logging.info(f"El estudiante {nombres} {apellidos} ya existe en el sistema. Pasando al siguiente.")
                    print(f"✅ El estudiante {nombres} {apellidos} ya existe en el sistema. Pasando al siguiente.")
                    
                    # Colorear fila como ya existente
                    for col_name, col_idx in column_indices.items():
                        try:
                            # leer del dataframe si la columna existe, sino del readsheet
                            if col_name in df.columns:
                                valor = fila[col_name]
                            else:
                                valor = read_sheet.cell_value(excel_row, col_idx)
                            sheet.write(excel_row, col_idx, valor, style_ya_existe)
                        except:
                            print(f"Error al colorear celda {col_name}: {str(e)}")  
                    wb.save(RUTA_EXCEL)
                    contador_ya_existentes += 1
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
                        
                        # Esperar a que cargue el formulario este visible y no este el load
                        wait.until(EC.invisibility_of_element_located((By.ID, "content-load")))
                        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "well")))

                        # verificamos que no tenga meses de busqueda dentro del fomulario
                        ya_registrado = verificar_meses_busqueda(driver)
                        
                        # si los tiene pasamos con el siguiente usuario
                        if ya_registrado:
                            logging.info(f"El estudiante {nombres} {apellidos} ya está registrado (mesesBusqueda > 1). Pasando al siguiente.")
                            print(f"✅ El estudiante {nombres} {apellidos} ya está registrado. Pasando al siguiente.")
                            
                            # Colorear fila como ya existente
                            for col_name, col_idx in column_indices.items():
                                try:
                                    # leer del dataframe si la columna existe, sino del readsheet
                                    if col_name in df.columns:
                                        valor = fila[col_name]
                                    else:
                                        valor = read_sheet.cell_value(excel_row, col_idx)
                                    sheet.write(excel_row, col_idx, valor, style_ya_existe)
                                except:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                                    
                            # Guardamos en el archivo la modificacion del estudiante
                            wb.save(RUTA_EXCEL)
                            # Aumentamos el contador
                            contador_ya_existentes += 1
                            continue  # Pasar al siguiente estudiante
                        
                        
                        # Verificar si estamos en la página del formulario completo
                        if driver.current_url == URL_FORMULARIO or "formulario" in driver.current_url.lower():
                            print("Formulario completo detectado. Procediendo a llenar...")
                            llenar_formulario_ubicaciones(driver)
                            llenar_formulario_ubicaciones_nacimiento(driver, wait)
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
                            resultado_esperiencia_laboral = experiencia_laboral(driver,  perfil_ocupacional)
                            if resultado_esperiencia_laboral == False:
                                # Colorear fila como error
                                for col_name, col_idx in column_indices.items():
                                    try:
                                        # leer del dataframe si la columna existe, sino del readsheet
                                        if col_name in df.columns:
                                            valor = fila[col_name]
                                        else:
                                            valor = read_sheet.cell_value(excel_row, col_idx)
                                        sheet.write(excel_row, col_idx, valor, style_error)
                                    except:
                                        print(f"Error al colorear celda {col_name}: {str(e)}")  
                                wb.save(RUTA_EXCEL)
                                print(f"Excel actualizado: marcando fila {excel_row + 1} como 'error")
                                contador_errores += 1

                                driver.get(URL_VERIFICACION)
                                time.sleep(2)
                                continue
                            
                            # Colorear fila como procesado exitosamente
                            for col_name, col_idx in column_indices.items():
                                try:
                                    # leer del dataframe si la columna existe, sino del readsheet
                                    if col_name in df.columns:
                                        valor = fila[col_name]
                                    else:
                                        valor = read_sheet.cell_value(excel_row, col_idx)
                                    sheet.write(excel_row, col_idx, valor, style_procesado)
                                except:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")  
                            wb.save(RUTA_EXCEL)
                            print(f"Excel actualizado: marcando fila {excel_row + 1} como 'procesado'")
                            contador_procesados_exitosamente += 1

                        else:
                            print(f"⚠️ No se detectó redirección al formulario completo. URL actual: {driver.current_url}")
                            # Colorear fila como error
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                            wb.save(RUTA_EXCEL)
                            contador_errores += 1
                    else:
                        print(" No se pudo completar la pre-inscripción")
                        # Colorear fila como error
                        for col_name, col_idx in column_indices.items():
                            try:
                                # leer del dataframe si la columna existe, sino del readsheet
                                if col_name in df.columns:
                                    valor = fila[col_name]
                                else:
                                    valor = read_sheet.cell_value(excel_row, col_idx)
                                sheet.write(excel_row, col_idx, valor, style_error)
                            except:
                                print(f"Error al colorear celda {col_name}: {str(e)}")
                        wb.save(RUTA_EXCEL)
                else:
                    print(f"⚠️ No se redirigió al formulario de pre-inscripción. URL actual: {driver.current_url}")
                    # Intentar redirigir manualmente
                    driver.get(URL_VERIFICACION)
                    print("Reintentando verificación...")
                    existe = verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido) 
                    if not existe:
                        print("Reintentando llenar datos...")
                        if llenar_datos_antes_de_inscripcion(nombres, apellidos,driver,wait):
                            # Si la reinscripción funciona, colorear como procesado
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_procesado)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                            contador_procesados_exitosamente += 1
                        else:
                            # Si falla, colorear como error
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except Exception as e:
                                    print(f"Error al colorear celda {col_name}: {str(e)}")
                            contador_errores += 1
                        wb.save(RUTA_EXCEL)
                    
            except Exception as e:
                logging.error(f"Error procesando estudiante {nombres} {apellidos}: {str(e)}")
                print(f" Error procesando estudiante {nombres} {apellidos}: {str(e)}")
                
                # Colorear fila como error
                for col_name, col_idx in column_indices.items():
                    try:
                        # leer del dataframe si la columna existe, sino del readsheet
                        if col_name in df.columns:
                            valor = fila[col_name]
                        else:
                            valor = read_sheet.cell_value(excel_row, col_idx)
                        sheet.write(excel_row, col_idx, valor, style_error)
                    except:
                        print(f"Error al colorear celda {col_name}: {str(e)}")
                try:
                    wb.save(RUTA_EXCEL)
                    print(f"Excel actualizado: marcando fila {excel_row + 1} como 'error'")
                except Exception as save_error:
                    print(f"Error al guardar Excel: {str(save_error)}")
                    
                    contador_errores += 1
                
                # Volver a la página de verificación para el siguiente estudiante
                try:
                    driver.get(URL_VERIFICACION)
                    time.sleep(2)
                except Exception as nav_error:
                    print(f"Error al navegar: {str(nav_error)}")
            
        try:
            # Encontrar la primera fila vacía después de los datos
            fila_resumen = total_registros + header_row + 3  # +3 para dejar espacio
            
            # Crear estilos para el resumen
            style_titulo_resumen = xlwt.XFStyle()
            font_titulo = xlwt.Font()
            font_titulo.bold = True
            font_titulo.colour_index = xlwt.Style.colour_map['white']
            style_titulo_resumen.font = font_titulo
            pattern_titulo = xlwt.Pattern()
            pattern_titulo.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern_titulo.pattern_fore_colour = xlwt.Style.colour_map['blue']
            style_titulo_resumen.pattern = pattern_titulo
            
            style_resumen = xlwt.XFStyle()
            font_resumen = xlwt.Font()
            font_resumen.bold = True
            style_resumen.font = font_resumen
            
            # Escribir el resumen
            sheet.write(fila_resumen, 0, "RESUMEN DE PROCESAMIENTO", style_titulo_resumen)
            sheet.write(fila_resumen + 1, 0, f"Aprendices procesados exitosamente:", style_resumen)
            sheet.write(fila_resumen + 1, 1, contador_procesados_exitosamente, style_resumen)
            sheet.write(fila_resumen + 5, 0, "Total procesados:", style_resumen)
            sheet.write(fila_resumen + 5, 1, total_registros, style_resumen)
            
            # Guardar los cambios finales
            wb.save(RUTA_EXCEL)
            
        except Exception as e:
            print(f"Error al escribir resumen en Excel: {str(e)}")
            logging.error(f"Error al escribir resumen en Excel: {str(e)}")
                
                
        logging.info("✅ Proceso completado exitosamente")
        print("\n===== ✅ Proceso completado exitosamente =====\n")
            
    except Exception as e:
        logging.error(f"Error general en el proceso: {str(e)}")
        print(f"Error general: {str(e)}")
        
    finally:
        sys.stdout = original_stdout # Restaurar la salida estándar
        # Cerrar el navegador al finalizar
        driver.quit()
        logging.info("Navegador cerrado")
        print("🔒 Navegador cerrado")
        
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ruta_archivo_excel = sys.argv[1]
        main(ruta_archivo_excel)
