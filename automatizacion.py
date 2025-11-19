# === LIBRERÍAS ESTÁNDAR DE PYTHON ===
import logging
import os
import sys
import threading
import time

# === LIBRERÍAS DE TERCEROS ===
import xlwt
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# === MÓDULOS LOCALES - Funciones de formularios ===
from funciones_formularios.campo_estrato import llenar_estrato
from funciones_formularios.campo_sueldo import llenar_formulario_sueldo
from funciones_formularios.campo_telefono_correo import llenar_formulario_telefono_correo
from funciones_formularios.experincia_laboral_campos import experiencia_laboral
from funciones_formularios.form_campo_estado_civil import llenar_formulario_estado_civil
from funciones_formularios.form_campo_perfil_ocupacional import llenar_input_perfil_ocupacional
from funciones_formularios.form_campos_nacimiento import llenar_formulario_ubicaciones_nacimiento
from funciones_formularios.form_campos_ubicacion_identificacion import llenar_formulario_ubicaciones
from funciones_formularios.form_datos_residencia import llenar_formulario_ubicacion_residencia
from funciones_formularios.login import login
from funciones_formularios.meses_busqueda import verificar_meses_busqueda
from funciones_formularios.preparar_excel import preparar_excel
from funciones_formularios.pre_inscripcion import llenar_datos_antes_de_inscripcion
from funciones_formularios.verificacion import (
    verificar_estudiante,
    verificar_estudiante_con_CC_primero
)

# === MÓDULOS LOCALES - Otros ===
from funciones_loggs.loggs_funciones import loggs
from perfilesOcupacionales.perfilExcepcion import PerfilOcupacionalNoEncontrado
from URLS.urls import URL_FORMULARIO, URL_VERIFICACION
from debug_exe import close_logger, get_log_path



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

def debug_log(mensaje, progress_queue=None):
    """Helper para loggear tanto en consola como en GUI"""
    print(mensaje)
    logging.info(mensaje)
    if progress_queue:
        progress_queue.put(("log", f"🔍 {mensaje}\n"))


def main(ruta_excel_param, progress_queue=None, username=None, password=None, stop_event=threading.Event(), pause_event=None):
    from debug_exe import init_logger, log, close_logger
    init_logger()
    log("="*50)
    log("INICIO DE main()")
    log(f"Ruta Excel: {ruta_excel_param}")
    log(f"progress_queue: {progress_queue is not None}")
    # ===== VALIDAR EVENTOS =====
    if stop_event is None:
        stop_event = threading.Event()
    if pause_event is None:
        pause_event = threading.Event()
        pause_event.set()  # Por defecto no pausado
    
    # Hacemos globales las variables que se usarán en todo el script
    global RUTA_EXCEL, df, wb, sheet, read_sheet, column_indices, header_row, programa_sin_perfil
    RUTA_EXCEL = ruta_excel_param
    original_stdout = sys.stdout
    
    if progress_queue:
        sys.stdout = QueueStream(progress_queue)

    try:
        # ===== PREPARAR EXCEL =====
        debug_log("Llamando a preparar_excel...", progress_queue)
        df, wb, sheet, read_sheet, column_indices, header_row, programa_sin_perfil = preparar_excel(RUTA_EXCEL)
        debug_log("preparar_excel completado", progress_queue)
        
    except PerfilOcupacionalNoEncontrado as e:
        nombre_programa = e.nombre_programa
        logging.warning(f"Perfil no encontrado para: {nombre_programa}")
        
        if progress_queue:
            progress_queue.put(("solicitar_perfil", nombre_programa))
            progress_queue.put(("log", f"⚠️ Proceso detenido: falta perfil para '{nombre_programa}'\n"))
            progress_queue.put(("log", f"📝 Ingresa el perfil ocupacional y reinicia el proceso\n"))
            progress_queue.put(("finish", None))
        else:
            return
    
    except (FileNotFoundError, Exception) as e:
        logging.error(f"Error fatal al preparar el archivo Excel: {e}")
        print(f"❌ Error fatal al preparar el archivo Excel: {e}")
        return

    try:
        # ===== VERIFICAR DETENCIÓN ANTES DE LOGIN =====
        if stop_event.is_set():
            print("🛑 Proceso detenido antes de iniciar login")
            return
        
        # Realizar login
        if not login(driver, username, password):
            logging.error("No se pudo completar el login. Abortando proceso.")
            return

        # Establecer columnas
        COLUMNA_TIPO_DOC = 'Tipo de Documento'
        COLUMNA_NUM_DOC = 'Número de Documento'
        COLUMNA_NOMBRES = 'Nombre'
        COLUMNA_APELLIDOS = 'Apellidos'
        COLUMNA_CELULAR = 'Celular'
        COLUMNA_CORREO = 'Correo Electrónico'
        COLUMNA_ESTADO = 'Estado'
        COLUMNA_PERFIL = 'Perfil Ocupacional'
        
        # Contadores
        contador_procesados_exitosamente = 0
        contador_ya_existentes = 0
        contador_errores = 0
        contador_saltados = 0
        
        total_registros = len(df)
        
        # =====  BUCLE PRINCIPAL CON CONTROL DE PAUSA/DETENCIÓN =====
        for i, fila in df.iterrows():
            excel_row = i + header_row + 1
            
            # ===== VERIFICAR DETENCIÓN =====
            # Verificar si se debe detener
            if stop_event.is_set():
                progress_queue.put(("log", f"Proceso detenido por el usuario.\n"))
                return
            
            if pause_event:
                pause_event.wait() 

            # ===== VERIFICAR PAUSA =====
            if not pause_event.is_set():
                print(f"\n{'='*50}")
                print(f"⏸️ PROCESO PAUSADO")
                print(f"{'='*50}")
                print(f"📍 Pausado en registro: {i + 1}/{total_registros}")
                print(f"💡 Esperando reanudación...")
                
                if progress_queue:
                    progress_queue.put(("log", f"\n⏸️ Pausado en registro {i + 1}/{total_registros}\n"))
                
                # Esperar hasta que se reanude o se detenga
                while not pause_event.is_set():
                    if stop_event.is_set():
                        print("🛑 Detenido durante pausa")
                        if progress_queue:
                            progress_queue.put(("log", "🛑 Proceso detenido durante pausa\n"))
                        return
                    time.sleep(0.5)  # Verificar cada 500ms
                
                print(f"\n{'='*50}")
                print(f"▶️ PROCESO REANUDADO")
                print(f"{'='*50}")
                print(f"📍 Continuando desde registro: {i + 1}/{total_registros}\n")
                
                if progress_queue:
                    progress_queue.put(("log", f"\n▶️ Reanudado desde registro {i + 1}/{total_registros}\n"))
            
            # ===== ENVIAR PROGRESO =====
            if progress_queue:
                progress_queue.put(("progress", (i + 1, total_registros)))

            # Inicializar variables
            nombres = "Sin nombre"
            apellidos = "Sin apellido"
            tipo_doc = ""
            num_doc = ""
            celular = ""
            correo = ""
            estado = ""
            perfil_ocupacional = ""
            
            try:
                # ===== COLOREAR COMO PROCESANDO =====
                for col_name, col_idx in column_indices.items():
                    try:
                        if col_name in df.columns:
                            valor = fila[col_name]
                        else:
                            valor = read_sheet.cell_value(excel_row, col_idx)
                        sheet.write(excel_row, col_idx, valor, style_procesando)
                    except Exception as e:
                        print(f"⚠️ Error al colorear celda {col_name}: {str(e)}")
                
                wb.save(RUTA_EXCEL)
                print(f"📝 Excel actualizado: marcando fila {excel_row + 1} como 'procesando'")
                
                # ===== EXTRAER DATOS =====
                tipo_doc = str(fila[COLUMNA_TIPO_DOC])
                num_doc = str(fila[COLUMNA_NUM_DOC])
                nombres = str(fila[COLUMNA_NOMBRES])
                apellidos = str(fila[COLUMNA_APELLIDOS])
                celular = str(fila[COLUMNA_CELULAR])
                correo = str(fila[COLUMNA_CORREO])
                estado = str(fila[COLUMNA_ESTADO])
                perfil_ocupacional = str(fila[COLUMNA_PERFIL])
                
                logging.info(f"Procesando estudiante {i+1}/{total_registros}: {nombres} {apellidos}")
                print(f"\n{'='*50}")
                print(f"📋 Procesando {i+1}/{total_registros}: {nombres} {apellidos}")
                print(f"{'='*50}\n")
                
                # ===== VERIFICAR DETENCIÓN ANTES DE VERIFICACIÓN =====
                if stop_event.is_set():
                    print("🛑 Detención solicitada antes de verificar estudiante")
                    break
                
                # ===== VERIFICAR SI EXISTE =====
                existe = verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido)
                
                if existe is None:
                    logging.warning(f"Saltando estudiante debido a error en verificación: {nombres} {apellidos}")
                    print(f"⚠️ Saltando estudiante debido a error en verificación")
                    
                    for col_name, col_idx in column_indices.items():
                        try:
                            if col_name in df.columns:
                                valor = fila[col_name]
                            else:
                                valor = read_sheet.cell_value(excel_row, col_idx)
                            sheet.write(excel_row, col_idx, valor, style_error)
                        except:
                            pass
                    wb.save(RUTA_EXCEL)
                    contador_saltados += 1
                    continue
                
                if existe:
                    logging.info(f"El estudiante {nombres} {apellidos} ya existe en el sistema")
                    print(f"✅ Estudiante ya existe en el sistema")
                    
                    for col_name, col_idx in column_indices.items():
                        try:
                            if col_name in df.columns:
                                valor = fila[col_name]
                            else:
                                valor = read_sheet.cell_value(excel_row, col_idx)
                            sheet.write(excel_row, col_idx, valor, style_ya_existe)
                        except:
                            pass
                    wb.save(RUTA_EXCEL)
                    contador_ya_existentes += 1
                    continue
                
                # ===== VERIFICAR DETENCIÓN ANTES DE REGISTRO =====
                if stop_event.is_set():
                    print("🛑 Detención solicitada antes de registrar estudiante")
                    break
                
                logging.info(f"El estudiante {nombres} {apellidos} no existe. Procediendo con el registro.")
                print(f"📝 Estudiante no existe. Procediendo con registro...")
                
                # ===== PROCESO DE REGISTRO (con verificaciones intercaladas) =====
                if URL_VERIFICACION in driver.current_url:
                    print("📄 Formulario de pre-inscripción detectado")
                    
                    if stop_event.is_set():
                        break
                    
                    if llenar_datos_antes_de_inscripcion(nombres, apellidos, driver):
                        print("✅ Pre-inscripción completada")
                        
                        wait.until(EC.invisibility_of_element_located((By.ID, "content-load")))
                        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "well")))

                        if stop_event.is_set():
                            break
                        
                        ya_registrado = verificar_meses_busqueda(driver)
                        
                        if ya_registrado:
                            logging.info(f"El estudiante {nombres} {apellidos} ya está registrado")
                            print(f"✅ Estudiante ya registrado (meses de búsqueda > 1)")
                            
                            for col_name, col_idx in column_indices.items():
                                try:
                                    if col_name in df.columns:
                                        valor = fila[col_name]
                                    else:
                                        valor = read_sheet.cell_value(excel_row, col_idx)
                                    sheet.write(excel_row, col_idx, valor, style_ya_existe)
                                except:
                                    pass
                            
                            wb.save(RUTA_EXCEL)
                            contador_ya_existentes += 1
                            continue
                        
                        if stop_event.is_set():
                            break
                        
                        # ===== LLENAR FORMULARIO CON VERIFICACIONES =====
                        if driver.current_url == URL_FORMULARIO or "formulario" in driver.current_url.lower():
                            print("📋 Llenando formulario completo...")
                            
                            # Cada función de llenado podría verificar stop_event internamente
                            llenar_formulario_ubicaciones(driver)
                            if stop_event.is_set(): break
                            
                            llenar_formulario_ubicaciones_nacimiento(driver, wait)
                            if stop_event.is_set(): break
                            
                            llenar_formulario_ubicacion_residencia(driver)
                            if stop_event.is_set(): break
                            
                            llenar_formulario_estado_civil(driver)
                            if stop_event.is_set(): break
                            
                            llenar_formulario_sueldo(driver)
                            if stop_event.is_set(): break
                            
                            llenar_estrato(driver)
                            if stop_event.is_set(): break
                            
                            llenar_formulario_telefono_correo(celular, correo, driver)
                            if stop_event.is_set(): break
                            
                            llenar_input_perfil_ocupacional(estado, driver)
                            if stop_event.is_set(): break
                            
                            print("✅ Formulario con datos básicos llenado correctamente")
                            
                            # Guardar información
                            botonGuardar = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, 'submitNewOft'))
                            )
                            botonGuardar.click()
                            print("✅ Click en botón Guardar")
                            logging.info("Se hizo click en el boton de Guardar")
                            
                            if stop_event.is_set():
                                break
                            
                            resultado_experiencia_laboral = experiencia_laboral(driver, perfil_ocupacional)
                            
                            if resultado_experiencia_laboral == False:
                                for col_name, col_idx in column_indices.items():
                                    try:
                                        if col_name in df.columns:
                                            valor = fila[col_name]
                                        else:
                                            valor = read_sheet.cell_value(excel_row, col_idx)
                                        sheet.write(excel_row, col_idx, valor, style_error)
                                    except:
                                        pass
                                wb.save(RUTA_EXCEL)
                                print(f"❌ Error en experiencia laboral")
                                contador_errores += 1
                                driver.get(URL_VERIFICACION)
                                time.sleep(2)
                                continue
                            
                            # ===== ÉXITO =====
                            for col_name, col_idx in column_indices.items():
                                try:
                                    if col_name in df.columns:
                                        valor = fila[col_name]
                                    else:
                                        valor = read_sheet.cell_value(excel_row, col_idx)
                                    sheet.write(excel_row, col_idx, valor, style_procesado)
                                except:
                                    pass
                            wb.save(RUTA_EXCEL)
                            print(f"✅ Registro completado exitosamente")
                            contador_procesados_exitosamente += 1

                        else:
                            print(f"⚠️ No se detectó formulario completo. URL: {driver.current_url}")
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except:
                                    pass
                            wb.save(RUTA_EXCEL)
                            contador_errores += 1
                    else:
                        print("❌ No se pudo completar la pre-inscripción")
                        for col_name, col_idx in column_indices.items():
                            try:
                                if col_name in df.columns:
                                    valor = fila[col_name]
                                else:
                                    valor = read_sheet.cell_value(excel_row, col_idx)
                                sheet.write(excel_row, col_idx, valor, style_error)
                            except:
                                pass
                        wb.save(RUTA_EXCEL)
                        contador_errores += 1
                else:
                    print(f"⚠️ No se redirigió al formulario. URL: {driver.current_url}")
                    driver.get(URL_VERIFICACION)
                    print("🔄 Reintentando verificación...")
                    existe = verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido)
                    if not existe:
                        print("🔄 Reintentando llenar datos...")
                        if llenar_datos_antes_de_inscripcion(nombres, apellidos, driver, wait):
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_procesado)
                                except:
                                    pass
                            contador_procesados_exitosamente += 1
                        else:
                            for col_name, col_idx in column_indices.items():
                                try:
                                    sheet.write(excel_row, col_idx, read_sheet.cell_value(excel_row, col_idx), style_error)
                                except:
                                    pass
                            contador_errores += 1
                        wb.save(RUTA_EXCEL)
                    
            except Exception as e:
                logging.error(f"Error procesando estudiante {nombres} {apellidos}: {str(e)}")
                print(f"❌ Error procesando estudiante: {str(e)}")
                
                for col_name, col_idx in column_indices.items():
                    try:
                        if col_name in df.columns:
                            valor = fila[col_name]
                        else:
                            valor = read_sheet.cell_value(excel_row, col_idx)
                        sheet.write(excel_row, col_idx, valor, style_error)
                    except:
                        pass
                
                try:
                    wb.save(RUTA_EXCEL)
                    print(f"📝 Excel actualizado: marcando como error")
                except Exception as save_error:
                    print(f"⚠️ Error al guardar Excel: {str(save_error)}")
                
                contador_errores += 1
                
                try:
                    driver.get(URL_VERIFICACION)
                    time.sleep(2)
                except Exception as nav_error:
                    print(f"⚠️ Error al navegar: {str(nav_error)}")
        
        # ===== RESUMEN FINAL =====
        try:
            fila_resumen = total_registros + header_row + 3
            
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
            
            sheet.write(fila_resumen, 0, "RESUMEN DE PROCESAMIENTO", style_titulo_resumen)
            sheet.write(fila_resumen + 1, 0, f"Aprendices procesados exitosamente:", style_resumen)
            sheet.write(fila_resumen + 1, 1, contador_procesados_exitosamente, style_resumen)
            sheet.write(fila_resumen + 2, 0, f"Ya existentes:", style_resumen)
            sheet.write(fila_resumen + 2, 1, contador_ya_existentes, style_resumen)
            sheet.write(fila_resumen + 3, 0, f"Errores:", style_resumen)
            sheet.write(fila_resumen + 3, 1, contador_errores, style_resumen)
            sheet.write(fila_resumen + 4, 0, f"Saltados:", style_resumen)
            sheet.write(fila_resumen + 4, 1, contador_saltados, style_resumen)
            sheet.write(fila_resumen + 5, 0, "Total procesados:", style_resumen)
            sheet.write(fila_resumen + 5, 1, total_registros, style_resumen)
            
            wb.save(RUTA_EXCEL)
            print("\n📊 Resumen guardado en Excel")
            
        except Exception as e:
            print(f"⚠️ Error al escribir resumen: {str(e)}")
            logging.error(f"Error al escribir resumen: {str(e)}")
        
        # ===== MENSAJE FINAL =====
        if stop_event.is_set():
            logging.info("🛑 Proceso detenido por el usuario")
            print("\n🛑 Proceso detenido por el usuario")
        else:
            logging.info("✅ Proceso completado exitosamente")
            print("\n✅ Proceso completado exitosamente")
            
    except Exception as e:
        logging.error(f"Error general en el proceso: {str(e)}")
        print(f"❌ Error general: {str(e)}")
        
    finally:
        sys.stdout = original_stdout
        driver.quit()
        logging.info("Navegador cerrado")
        print("🔒 Navegador cerrado")
        
        log_path = get_log_path()
        if log_path:
            print(f"\n📝 Archivo de log guardado en:\n{log_path}")
        close_logger()
        
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ruta_archivo_excel = sys.argv[1]
        main(ruta_archivo_excel)
