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
from openpyxl.styles import PatternFill
import xlrd
import xlwt
from xlutils.copy import copy

# Configurar logging
log_filename = f"automatizacion_aprendices_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Cargar variables de entorno
load_dotenv()



# --- Configuración inicial ---


RUTA_EXCEL = 'C:/Users/SENA/Documents/Reporte de Aprendices Ficha 3199341.xls'

URL_LOGIN = 'https://agenciapublicadeempleo.sena.edu.co/spe-web/spe/login'
URL_VERIFICACION = 'https://agenciapublicadeempleo.sena.edu.co/spe-web/spe/funcionario/oferta'
URL_FORMULARIO = 'https://agenciapublicadeempleo.sena.edu.co/spe-web/spe/oferta/new'

# --- Credenciales de login desde variables de entorno ---
USUARIO_LOGIN = os.getenv('USUARIO_LOGIN')
CONTRASENA_LOGIN = os.getenv('CONTRASENA_LOGIN')
if not USUARIO_LOGIN or not CONTRASENA_LOGIN:
    error_msg = "Error: Las credenciales de login no están configuradas en el archivo .env"
    logging.error(error_msg)
    print(error_msg)
    exit()

# Mapeo de tipos de documento
TIPOS_DOCUMENTO = {
    "CC": "1",
    "TI": "2", 
    "CE": "3",
    "Otro Documento de Identidad": "5",
    "PEP": "8",
    "PPT": "9"
}


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


# --- Inicializar el navegador Selenium con opciones ---
chrome_options = Options()

# Descomentar la siguiente línea para modo headless (sin interfaz gráfica), para visualizar el funcionamiento del aplicativo
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)  # Espera explícita de 10 segundos

# -- Funcion que Permite el Logueo del Funcionario APE (Agencia Publica de Empleo)
def login():
    """Realiza el proceso de login en la aplicación"""
    try:

        driver.get(URL_LOGIN)
        logging.info("Abriendo página de login...")
        
        # Esperar a que el radio button esté disponible
        radio_persona = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='radio' and @name='tipousuario' and @value='0']")
        ))
        radio_persona.click()
        
        # Esperar y completar campos de login
        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, 'username')))
        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO_LOGIN)
        
        campo_contrasena = driver.find_element(By.ID, 'password')
        campo_contrasena.clear()
        campo_contrasena.send_keys(CONTRASENA_LOGIN)
        
        boton_entrar = driver.find_element(By.ID, 'entrar')
        boton_entrar.click()
        
        # Esperar a que se complete el login
        wait.until(EC.url_changes(URL_LOGIN))
        logging.info("Login exitoso")
        return True
        
    except Exception as e:
        logging.error(f"Error durante el login: {str(e)}")
        print(f"Error durante el login: {str(e)}")
        return False
    
def verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos):
    """
    Intenta verificar al estudiante primero con CC, y luego con el tipo de documento
    original solo si no se encontró con CC.
    """
    if tipo_doc == "CC":
        # Si ya es CC, verificamos directamente
        print("Tipo de documento ya es CC, verificando directamente...")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos)
    else:
        # Primero intentamos con CC
        print(f"Intentando verificar primero con CC para documento {num_doc}...")
        encontrado_con_cc = verificar_estudiante("CC", num_doc, nombres, apellidos)
        
        # Si lo encontramos con CC, retornamos True
        if encontrado_con_cc is True:
            print(f"✅ Estudiante {num_doc} encontrado con CC, aunque el tipo original era {tipo_doc}")
            return True
            
        # Si la verificación con CC es None (error) o False (no encontrado),
        # intentamos con el tipo de documento original
        print(f"No se encontró con CC, intentando con tipo original {tipo_doc}")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos)

def verificar_estudiante(tipo_doc, num_doc, nombres, apellidos):
    """Verifica si un estudiante ya está registrado en el sistema"""
    max_intentos = 3
    
    for intento in range(1, max_intentos + 1):
        try:
            print(f"Intento {intento} - Abriendo URL: {URL_VERIFICACION}")
            driver.get(URL_VERIFICACION)
            
            # Esperar a que la página cargue completamente
            print("Esperando que la página cargue...")
            wait = WebDriverWait(driver, 20)
            wait.until(EC.visibility_of_element_located((By.ID, 'dropTipoIdentificacion')))
            print("Página cargada correctamente")
            
            logging.info(f"Verificando estudiante: {nombres} {apellidos} - Documento: {num_doc} - Tipo Doc: {tipo_doc}")
            
            # Limpiar caché y cookies antes de interactuar
            # Esperar un momento después de limpiar
            time.sleep(1)
            
            # Seleccionar tipo de documento
            tipo_doc_dropdown = wait.until(EC.element_to_be_clickable((By.ID, 'dropTipoIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", tipo_doc_dropdown)
            print("Seleccionando tipo de documento...")
            
            # Crear el objeto Select
            selector = Select(tipo_doc_dropdown)
            
            # Seleccionar tipo de documento
            value_tipo_doc = TIPOS_DOCUMENTO.get(tipo_doc)
            if value_tipo_doc:
                print(f"Seleccionando valor: {value_tipo_doc}")
                selector.select_by_value(value_tipo_doc)
            else:
                print(f"Seleccionando texto visible: {tipo_doc}")
                selector.select_by_visible_text(tipo_doc)
            
            time.sleep(1.5)
            
            # Completar número de documento
            print("Ingresando número de documento...")
            campo_num_id = wait.until(EC.element_to_be_clickable((By.ID, 'numeroIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_num_id)
            campo_num_id.click()
            time.sleep(0.5)
            campo_num_id.clear()
            time.sleep(0.5)
            
            # Ingresar dígito por dígito
            for digito in str(num_doc):
                campo_num_id.send_keys(digito)
                time.sleep(0.1)
            
            print(f"Documento ingresado: {num_doc}")
            time.sleep(1)
            
            # Hacer clic en buscar con JavaScript
            print("Haciendo clic en buscar...")
            boton_buscar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@id='btnBuscar']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", boton_buscar)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", boton_buscar)
            
            # Esperar más tiempo para asegurar que los resultados se carguen
            print("Esperando resultados...")
            time.sleep(5)  # Aumentado a 5 segundos
            
            # Capturar screenshot para diagnóstico
            screenshot_path = f"busqueda_{num_doc}_{int(time.time())}.png"
            driver.save_screenshot(screenshot_path)
            print(f"Screenshot guardado en {screenshot_path}")
            
            # VERIFICACIÓN DE EXISTENCIA
            # 1. Verificar si hay una tabla de resultados con datos
            tablas = driver.find_elements(By.TAG_NAME, "table")
            for tabla in tablas:
                if tabla.is_displayed():
                    filas = tabla.find_elements(By.TAG_NAME, "tr")
                    if filas and len(filas) > 1:  # Si hay más de una fila (encabezado + datos)
                        print(f"✅ ENCONTRADO: Tabla con {len(filas)} filas tiene resultados")
                        return True
                    
            # 2. Verificar si aparece un formulario para completar datos (caso en que no existe)
            try:
                # Buscar campos típicos del formulario de pre-inscripción
                campos_formulario = [
                    driver.find_element(By.NAME, 'nombres'),
                    driver.find_element(By.NAME, 'primerApellido')
                ]
                
                if all(campo.is_displayed() for campo in campos_formulario):
                    print(f"❌ NO ENCONTRADO: Se muestra formulario para completar datos de {num_doc}")
                    return False
            except:
                pass
            
            # 3. Verificar mensajes explícitos
            page_source = driver.page_source.lower()
            if "no se encontraron resultados" in page_source:
                print(f"❌ NO ENCONTRADO: No hay resultados para el estudiante {num_doc}")
                return False
                
            # 4. Buscar indicadores positivos
            if any(indicator in page_source for indicator in ["ya existe", "encontrado", str(num_doc)]):
                print(f"✅ ENCONTRADO: Indicadores sugieren que el estudiante {num_doc} existe")
                return True
            
            # Si no se pudo determinar claramente, verificar la URL actual
            # Generalmente cuando no existe el estudiante, la página redirige a un formulario
            if "/inscripcion" in driver.current_url or "/register" in driver.current_url:
                print(f"❌ NO ENCONTRADO: Redirigido a formulario para {num_doc}")
                return False
                
            # Por defecto, asumir que no existe si no hay evidencia clara
            print(f"❓ INDETERMINADO: No hay evidencia clara sobre {num_doc}, asumiendo que no existe")
            return False
                
        except Exception as e:
            logging.error(f"Error en intento {intento}: {str(e)}")
            print(f"Error en verificación (intento {intento}): {str(e)}")
            
            if intento == max_intentos:
                print("Se agotaron los intentos. No se pudo verificar el estudiante.")
                return None
            else:
                print(f"Reintentando... ({intento}/{max_intentos})")
                time.sleep(2)
    
    return None
    
def llenar_datos_antes_de_inscripcion(nombres_excel, apellidos_excel):
    """Llena los campos de nombres, apellidos, fecha de nacimiento y género antes de la inscripción exhaustiva."""
    try:
        logging.info(f"Intentando llenar los campos antes de la inscripción con: {nombres_excel} {apellidos_excel}")
        logging.info(f"URL actual: {driver.current_url}")
        
        # Esperar a que la página de pre-inscripción cargue completamente
        print("Esperando que el formulario de pre-inscripción cargue...")
        
        # Capturar screenshot antes de interactuar con el formulario
        driver.save_screenshot(f"pre_inscripcion_{int(time.time())}.png")
        
        # --- Nombres ---
        print("Buscando campo de nombres...")
        try:
            # Esperar explícitamente a que el campo nombres esté presente
            campo_nombres_pre = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, 'nombres'))
            )
            
            # Asegurar que el elemento es visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_nombres_pre)
            time.sleep(0.5)
            
            # Interactuar con el campo
            campo_nombres_pre.click()
            campo_nombres_pre.clear()
            for letra in nombres_excel:
                campo_nombres_pre.send_keys(letra)
                time.sleep(0.05)
                
            print(f"✅ Se llenó el campo Nombres con: {nombres_excel}")
            logging.info(f"Se llenó el campo Nombres con: {nombres_excel}")
        except Exception as e:
            print(f"❌ Error al llenar el campo Nombres: {str(e)}")
            logging.error(f"Error al llenar el campo Nombres: {str(e)}")

        # --- Apellidos ---
        print("Preparando apellidos...")
        apellidos_partes = apellidos_excel.split(' ', 1)
        primer_apellido = apellidos_partes[0] if len(apellidos_partes) > 0 else ''
        segundo_apellido = apellidos_partes[1] if len(apellidos_partes) > 1 else ''

        try:
            print("Buscando campo de primer apellido...")
            campo_primer_apellido_pre = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.NAME, 'primerApellido'))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_primer_apellido_pre)
            time.sleep(0.5)
            
            campo_primer_apellido_pre.click()
            campo_primer_apellido_pre.clear()
            for letra in primer_apellido:
                campo_primer_apellido_pre.send_keys(letra)
                time.sleep(0.05)
                
            print(f"✅ Se llenó el campo Primer apellido con: {primer_apellido}")
            logging.info(f"Se llenó el campo Primer apellido con: {primer_apellido}")
        except Exception as e:
            print(f"❌ Error al llenar el campo Primer apellido: {str(e)}")
            logging.error(f"Error al llenar el campo Primer apellido: {str(e)}")

        try:
            print("Buscando campo de segundo apellido...")
            campo_segundo_apellido_pre = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.NAME, 'segundoApellido'))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_segundo_apellido_pre)
            time.sleep(0.5)
            
            campo_segundo_apellido_pre.click()
            campo_segundo_apellido_pre.clear()
            for letra in segundo_apellido:
                campo_segundo_apellido_pre.send_keys(letra)
                time.sleep(0.05)
                
            print(f"✅ Se llenó el campo Segundo apellido con: {segundo_apellido}")
            logging.info(f"Se llenó el campo Segundo apellido con: {segundo_apellido}")
        except Exception as e:
            print(f"❌ Error al llenar el campo Segundo apellido: {str(e)}")
            logging.error(f"Error al llenar el campo Segundo apellido: {str(e)}")

        # --- Fecha de Nacimiento ---
        try:
            print("Buscando campo de fecha de nacimiento...")
            campo_fecha_nacimiento = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.NAME, 'fechaNacimiento'))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_fecha_nacimiento)
            time.sleep(0.5)
            
            campo_fecha_nacimiento.click()
            campo_fecha_nacimiento.clear()
            campo_fecha_nacimiento.send_keys('01-01-2000')  # Formato DDMMYYYY
            
            print("✅ Se llenó el campo Fecha de Nacimiento con: 01-01-2000")
            logging.info("Se llenó el campo Fecha de Nacimiento con: 01-01-2000")
        except Exception as e:
            print(f"❌ Error al llenar el campo Fecha de Nacimiento: {str(e)}")
            logging.error(f"Error al llenar el campo Fecha de Nacimiento: {str(e)}")

        # --- Género ---
        try:
            print("Determinando género...")
            genero = determinar_genero(nombres_excel)
            if genero is not None:
                print(f"Género determinado: {genero} ({'Femenino' if genero == '1' else 'Masculino'})")
                
                elemento_genero = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.NAME, 'genero'))
                )
                            
                # Haz scroll hacia el elemento
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_genero)
                time.sleep(0.5)

                # Luego crea el objeto Select usando el elemento
                selector_genero = Select(elemento_genero)
                selector_genero.select_by_value(genero)
                time.sleep(0.5)
                
                selector_genero.select_by_value(genero)
                time.sleep(1)
                
                print(f"✅ Se seleccionó el género: {genero} ({'Femenino' if genero == '1' else 'Masculino'})")
                logging.info(f"Se seleccionó el género: {genero} ({'Femenino' if genero == '1' else 'Masculino'})")
            else:
                print("⚠️ No se pudo determinar el género, intentando seleccionar por defecto Masculino")
                selector_genero = Select(WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.NAME, 'genero'))
                ))
                selector_genero.select_by_value('0')  # Masculino por defecto
                logging.warning(f"No se pudo determinar el género para: {nombres_excel}, se seleccionó Masculino por defecto")
        except Exception as e:
            print(f"❌ Error al seleccionar el género: {str(e)}")
            logging.error(f"Error al seleccionar el género: {str(e)}")

        # Verificar si hay un botón para continuar y hacer clic en él
        try:
            print("Esperando a que el botón para verificar sea clickeable...")
            try:
                botonVerificar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'btnVerificar'))
                )
                
                if botonVerificar.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botonVerificar)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", botonVerificar)
                    print("✅ Se hizo clic en el botón para verificar")
                    logging.info("Se hizo clic en el botón para verificar")
                    
                    # Esperar un momento para que se procese la verificación
                    time.sleep(2)
                    
                    # Ahora buscar el botón de continuar/inscribir
                    print("Esperando a que el botón para continuar/inscribir sea clickeable...")
                    try:
                        botonContinuar = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, 'btnContinuar'))
                        )
                        
                        if botonContinuar.is_displayed():
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botonContinuar)
                            time.sleep(0.5)
                            driver.execute_script("arguments[0].click();", botonContinuar)
                            print("✅ Se hizo clic en el botón para continuar/inscribir")
                            logging.info("Se hizo clic en el botón para continuar/inscribir")
                        else:
                            print("⚠️ El botón para continuar/inscribir no está visible")
                    except Exception as e:
                        print(f"❌ Error al buscar/hacer clic en el botón continuar/inscribir: {str(e)}")
                        logging.error(f"Error al buscar/hacer clic en el botón continuar/inscribir: {str(e)}")
                else:
                    print("⚠️ El botón para verificar no está visible")
            except Exception as e:
                print(f"❌ Error al buscar/hacer clic en el botón verificar: {str(e)}")
                logging.error(f"Error al buscar/hacer clic en el botón verificar: {str(e)}")

        except Exception as e:
            print(f"❌ Error general al manejar los botones: {str(e)}")
            logging.error(f"Error general al manejar los botones: {str(e)}")
        # Capturar screenshot después de llenar el formulario
        driver.save_screenshot(f"pre_inscripcion_completada_{int(time.time())}.png")
        print("✅ Formulario de pre-inscripción completado")
        return True

    except Exception as e:
        logging.error(f"Error general al llenar los campos antes de la inscripción: {str(e)}")
        print(f"❌ Error general al llenar los campos antes de inscripción: {str(e)}")
        
        # Capturar screenshot en caso de error
        driver.save_screenshot(f"error_pre_inscripcion_{int(time.time())}.png")
        return False

def determinar_genero(nombre):
    """Intenta determinar el género a partir del nombre."""
    nombre = nombre.lower()
    nombres_femeninos = ['ana', 'maria', 'sofia', 'isabella', 'valentina', 'camila', 'mariana', 'laura', 'daniela', 'valeria', 'liz', 'ariana', 'lizeth', 'danna'] #nombres femeninos
    nombres_masculinos = ['juan', 'carlos', 'luis', 'andres', 'sebastian', 'mateo', 'santiago', 'alejandro', 'daniel', 'gabriel'] #nombres masculinos

    primer_nombre = nombre.split(' ')[0] # Tomamos el primer nombre

    if primer_nombre in nombres_femeninos:
        return '1'  # En el aplicativo 1 es femenino
    elif primer_nombre in nombres_masculinos:
        return '0'  # En el aplicativo 0 es Masculino
    else:
        return None # No se pudo determinar el género

# -- Esta funcion selcciona la ubicacion donde se saco la indentificacion --
def llenar_formulario_ubicaciones(driver):
    """Llena los campos de País, Departamento y Municipio en el formulario principal."""
    try:
        wait = WebDriverWait(driver, 10)

        # 1. Hacer clic en el desplegable "Seleccionar"
        seleccionar_ubicacion_boton = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.dropdown-toggle"))
        )

        seleccionar_ubicacion_boton.click()

        # Seleccionar País: Colombia
        pais_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'dropPais')))
        selector_pais = Select(pais_dropdown)
        selector_pais.select_by_visible_text('Colombia')
        logging.info("Se seleccionó Colombia en el desplegable de País.")
        print("✅ Se seleccionó Colombia en el desplegable de País.")
        time.sleep(1)  # Esperar a que se cargue el departamento

        # Seleccionar Departamento: Cundinamarca
        departamento_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'dropDepartamento')))
        selector_departamento = Select(departamento_dropdown)
        selector_departamento.select_by_visible_text('Cundinamarca')
        logging.info("Se seleccionó Cundinamarca en el desplegable de Departamento.")
        print("✅ Se seleccionó Cundinamarca en el desplegable de Departamento.")
        time.sleep(1)  # Esperar a que se cargue el municipio

        # Seleccionar Municipio: Mosquera
        municipio_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'dropMunicipio')))
        selector_municipio = Select(municipio_dropdown)
        selector_municipio.select_by_visible_text('Mosquera')
        logging.info("Se seleccionó Mosquera en el desplegable de Municipio.")
        print("✅ Se seleccionó Mosquera en el desplegable de Municipio.")

        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario principal (ubicación): {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False

def llenar_formulario_ubicaciones_nacimiento(driver):
    """Llena los campos de ubicación de Nacimiento evitando conflictos con el primer formulario."""
    try:
        wait = WebDriverWait(driver, 15)
        
        print("⚪ Ingresando Datos de Nacimiento")
        
        # 1. Hacer clic en el botón "Seleccionar" para el municipio de nacimiento
        # Usar un selector mucho más específico
        seleccionar_boton = wait.until(
            EC.element_to_be_clickable((By.XPATH, 
                "//*[@id='municipioNacimientoText']/following-sibling::div//button[contains(@class, 'dropdown-toggle')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", seleccionar_boton)
        seleccionar_boton.click()
        print("✅ Se hizo clic en el botón Seleccionar para datos de nacimiento")
        time.sleep(1)
        
        # 2. Verificar que el dropdown se haya abierto
        dropdown = wait.until(
            EC.visibility_of_element_located((By.ID, "dropNacimiento"))
        )
        print("✅ Se abrió el dropdown de nacimiento correctamente")
        
        # 3. Importante: CAMBIAR AL CONTEXTO DEL DROPDOWN
        # En vez de buscar por ID directamente (lo cual puede encontrar elementos duplicados),
        # vamos a buscar elementos dentro del dropdown específico
        
        # Seleccionar País
        # Usamos XPath para encontrar el select dentro del dropdown específico
        pais_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][1]//select"))
        )
        Select(pais_select).select_by_visible_text('Colombia')
        print("✅ Se seleccionó Colombia en el desplegable de País (Nacimiento).")
        time.sleep(1.5)  # Esperar más tiempo para que se carguen opciones
        
        # Seleccionar Departamento
        # Nuevamente, usamos XPath para encontrar el segundo select dentro del dropdown
        depto_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][2]//select"))
        )
        Select(depto_select).select_by_visible_text('Cundinamarca')
        print("✅ Se seleccionó Cundinamarca en el desplegable de Departamento (Nacimiento).")
        time.sleep(1.5)
        
        # Seleccionar Municipio
        # Y el tercer select dentro del dropdown
        muni_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][3]//select"))
        )
        Select(muni_select).select_by_visible_text('Mosquera')
        print("✅ Se seleccionó Mosquera en el desplegable de Municipio (Nacimiento).")
        
        # 4. Cerrar el dropdown haciendo clic en algún otro elemento para confirmar la selección
        try:
            fecha_input = driver.find_element(By.ID, "fechaNacimiento")
            fecha_input.click()
            print("✅ Se cerró el dropdown haciendo clic fuera")
        except:
            # Si el input de fecha no está disponible, intenta hacer clic en cualquier otro elemento
            try:
                # Buscar un elemento visible y hacer clic en él
                driver.find_element(By.XPATH, "//h4[contains(text(), 'Datos de nacimiento')]").click()
                print("✅ Se cerró el dropdown haciendo clic en el título")
            except:
                print("No se pudo cerrar el dropdown manualmente")
        
        return True
        
    except Exception as e:
        print(f"❌ Error al llenar la ubicación de Nacimiento: {e}")
        logging.error(f"Error detallado al llenar la ubicación de Nacimiento: {str(e)}")
        
        # Información adicional de depuración
        try:
            # Verificar si el dropdown está abierto
            dropdown_visible = len(driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento:not([style*='display: none'])")) > 0
            print(f"¿Dropdown visible? {dropdown_visible}")
            
            if dropdown_visible:
                # Contar cuántos divs de control-group hay dentro del dropdown
                control_groups = driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento .control-group")
                print(f"Grupos de control encontrados: {len(control_groups)}")
                
                # Contar cuántos selects hay dentro del dropdown
                selects = driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento select")
                print(f"Selects encontrados dentro del dropdown: {len(selects)}")
                
                # Ver qué opciones hay disponibles en el primer select
                if len(selects) > 0:
                    options = selects[0].find_elements(By.TAG_NAME, "option")
                    option_texts = [opt.text for opt in options]
                    print(f"Opciones en el primer select: {option_texts}")
            
        except Exception as debug_e:
            print(f"Error en depuración: {debug_e}")
        
        return False
    

# -- Esta funcion selcciona la ubicacion donde se saco la es la residemcia del aprendiz --
def llenar_formulario_ubicacion_residencia(driver):
    """Llena los campos de País, Departamento y Municipio en el formulario principal. (Residencia)"""
    try:
        wait = WebDriverWait(driver, 10)
        # Seleccionar País: Colombia
        pais = wait.until(EC.presence_of_element_located((By.ID, 'pais')))
        selector_pais = Select(pais)
        selector_pais.select_by_visible_text('Colombia')
        logging.info("Se seleccionó Colombia en el desplegable de País.")
        print("✅ Se seleccionó Colombia en el desplegable de País.")
        time.sleep(1)  # Esperar a que se cargue el departamento

        # Seleccionar Departamento: Cundinamarca
        departamento = wait.until(EC.presence_of_element_located((By.ID, 'departamento')))
        selector_departamento = Select(departamento)
        selector_departamento.select_by_visible_text('Cundinamarca')
        logging.info("Se seleccionó Cundinamarca en el desplegable de Departamento.")
        print("✅ Se seleccionó Cundinamarca en el desplegable de Departamento.")
        time.sleep(1)  # Esperar a que se cargue el municipio

        # Seleccionar Municipio: Mosquera
        municipio = wait.until(EC.presence_of_element_located((By.ID, 'municipio')))
        selector_municipio = Select(municipio)
        selector_municipio.select_by_visible_text('Mosquera')
        logging.info("Se seleccionó Mosquera en el desplegable de Municipio.")
        print("✅ Se seleccionó Mosquera en el desplegable de Municipio.")

        direccion_residencia_input = wait.until(EC.presence_of_element_located((By.ID, 'residenciaDireccion')))
        direccion_residencia_input.send_keys('MOSQUERA')
        logging.info("Se coloco Mosquea en el input")
        print("✅ Mosquera ingresado en el input")

        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario principal (residencia): {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False
    
# -- Esta funcion selcciona la ubicacion donde se saco la es la residemcia del aprendiz --
def llenar_formulario_estado_civil(driver):
    """Llena los campos de Estdo Civil, Cantidad de personas a cargo, Posicion familiar, Libreta militar, Migrante, donde habita."""
    try:
        wait = WebDriverWait(driver, 10)
        # Seleccionar Estado Civil: Soltero
        estado_civil = wait.until(EC.presence_of_element_located((By.ID, 'estadoCivil')))
        selector_estado = Select(estado_civil)
        selector_estado.select_by_visible_text('Soltero')
        logging.info("Se seleccionó Soltero en el desplegable de Estado Civil.")
        print("✅ Se seleccionó Soltero en el desplegable de Estado Civil.")
        time.sleep(1)  # Esperar a que se cargue el estado civil

        # Seleccionar Cantidad de Personas a Cargo: 0
        personas_a_cargo_input = wait.until(EC.presence_of_element_located((By.ID, 'personasCargo')))
        personas_a_cargo_input.send_keys('0')
        logging.info("Se ingreso 0 en el input de Personas a Cargo.")
        print("✅ Se ingreso 0 en el input de Personas a Cargo.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Posicion Familiar: Otro
        posicion_familiar = wait.until(EC.presence_of_element_located((By.ID, 'posicionFamiliar')))
        selector_familiar = Select(posicion_familiar)
        selector_familiar.select_by_visible_text('Otro')
        logging.info("Se seleccionó Otro en el desplegable de Posicion Familiar.")
        print("✅ Se seleccionó Otro en el desplegable de Posicion Familiar.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Libreta Militar: No tiene libreta militar
        libreta_militar = wait.until(EC.presence_of_element_located((By.ID, 'libretaMilitar')))
        selector_militar = Select(libreta_militar)
        selector_militar.select_by_visible_text('No tiene libreta militar')
        logging.info("Se seleccionó 'No tiene libreta militar' en el desplegable de libreta militar.")
        print("✅ Se seleccionó 'No tiene libreta militar' en el desplegable de libreta militar.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Migrante: No
        personas_a_cargo_input = wait.until(EC.presence_of_element_located((By.ID, 'migrante')))
        personas_a_cargo_input.send_keys('No')
        logging.info("Se ingreso 'No' en el input de Migrante.")
        print("✅ Se ingreso 'No' en el input de Migrante.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar en Zona de residencia: Urbana
        zona_residencia = wait.until(EC.presence_of_element_located((By.ID, 'esRural')))
        selector_zona_residencia = Select(zona_residencia)
        selector_zona_residencia.select_by_visible_text('Urbana')
        logging.info("Se seleccionó 'Urbano' en el desplegable de zona de resindecia.")
        print("✅ Se seleccionó 'Urbano' en el desplegable de zona de resindecia.")
        time.sleep(1)  # Esperar a que se cargue
        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario principal (residencia): {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False
    
# -- Esta funcion selcciona la ubicacion donde se saco la es la residemcia del aprendiz --
def llenar_formulario_sueldo(driver):
    """Llena los campos de sueldo que aceptaria"""
    try:
        wait = WebDriverWait(driver, 10)

        # Seleccionar Sueldo que aceptaria: $ 1.000.000  -  $1.500.000
        sueldo = wait.until(EC.presence_of_element_located((By.ID, 'sueldo')))
        selector_sueldo = Select(sueldo)
        selector_sueldo.select_by_value('3')
        logging.info("Se seleccionó '$ 1.000.000  -  $1.500.000' en el desplegable de Sueldo.")
        print("✅ Se seleccionó '$ 1.000.000  -  $1.500.000' en el desplegable de Sueldo.")
        time.sleep(1)  # Esperar a que se cargue el estado civil


    except Exception as e:
        error_msg = f"Error al llenar el formulario sueldo: {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False

# Funcion que llenara el numero de telefono y correo
def llenar_formulario_telefono_correo(telefono_excel, correo_excel):
    """Llena los campos de telefono y correo electronico en el formualario principal."""
    try:
        # --- Telefono ---
        print("Buscando campo de telefono...")
        try:
            # Esperar explícitamente a que el campo telefono esté presente
            campo_telefono_pre = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, 'telefonoCelular'))
            )
            
            # Definir el número de teléfono por defecto
            telefono_por_defecto = "3101234567"
            
            # Verificación para valores vacíos o NaN
            telefono_str = str(telefono_excel).lower().strip()
            if telefono_excel is None or pd.isna(telefono_excel) or telefono_str == '' or telefono_str == 'nan' or telefono_str == 'none':
                print(f"⚠️ El número de teléfono del Excel está vacío o es NaN. Usando el valor por defecto: {telefono_por_defecto}")
                telefono_a_usar = telefono_por_defecto
            else:
                telefono_a_usar = str(telefono_excel)
            
            # Asegurar que el elemento es visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_telefono_pre)
            time.sleep(0.5)
            
            # Interactuar con el campo
            campo_telefono_pre.click()
            campo_telefono_pre.clear()
            for numero in str(telefono_a_usar):
                campo_telefono_pre.send_keys(numero)
                time.sleep(0.02)
            
            print(f"✅ Se llenó el campo telefono con: {telefono_a_usar}")
            logging.info(f"Se llenó el campo telefono con: {telefono_a_usar}")
        except Exception as e:
            print(f"❌ Error al llenar el campo telefono: {str(e)}")
            logging.error(f"Error al llenar el campo telefono: {str(e)}")
        
        # --- Correo ---
        print("Preparando Correo...")
        try:
            print("Buscando campo de Correo...")
            campo_correo = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.NAME, 'email'))
            )
            
            # Verificar si el correo está vacío o es NaN
            correo_por_defecto = "ejemplo@correo.com"
            if correo_excel is None or pd.isna(correo_excel) or str(correo_excel).strip() == '':
                print(f"⚠️ El correo del Excel está vacío o es NaN. Usando el valor por defecto: {correo_por_defecto}")
                correo_a_usar = correo_por_defecto
            else:
                correo_a_usar = str(correo_excel)  # Convertir explícitamente a string
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_correo)
            time.sleep(0.5)
            
            campo_correo.click()
            campo_correo.clear()
            for letra in correo_a_usar:
                campo_correo.send_keys(letra)
                time.sleep(0.05)
            
            print(f"✅ Se llenó el campo Correo con: {correo_a_usar}")
            logging.info(f"Se llenó el campo Correo con: {correo_a_usar}")
        except Exception as e:
            print(f"❌ Error al llenar el campo Correo: {str(e)}")
            logging.error(f"Error al llenar el campo Correo: {str(e)}")
    except Exception as e:
        print(f"Error al llenar campos de Telefono y Correo: {str(e)}")
        logging.error(f"Error general al llenar campos de Telefono y Correo: {str(e)}")


#funcion para llenar el estrato en el formulario principal
def llenar_estrato(driver):
    """LLena el campo de estrato"""
    try:   
        wait = WebDriverWait(driver, 10)

        estrato = wait.until(EC.presence_of_element_located((By.ID, 'estrato')))
        selector_suelddo = Select(estrato)
        selector_suelddo.select_by_visible_text('Dos')
        logging.info("Se selecciono estrato Dos en el campo de Estrato.")
        print("Se selecciono Dos en el campo de Estrato.")
        time.sleep(1)

    except Exception as e:
        error_msg: f"Error al llenar el formualario estrato: {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False
    
# Funcion que llenara el numero de telefono y correo
def llenar_input_perfil_ocupacional(estado):
    """Llena los campos de Perfil Ocupacional en el formualario principal."""
    try:
        # --- Nombres ---
        print("Buscando campo de Perfil Ocupacional...")
        try:
            # Esperar explícitamente a que el campo Perfil Ocupacional esté presente
            campo_perfil_ocupacional = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'descripcion'))
            )
            
            
            # Interactuar con el campo
            campo_perfil_ocupacional.click()
            campo_perfil_ocupacional.clear()
            for letra in estado:
                campo_perfil_ocupacional.send_keys(letra)
                time.sleep(0.05)
                
            print(f"✅ Se llenó el campo Perfil Ocupacional con: {estado}")
            logging.info(f"Se llenó el campo Perfil Ocupacional con: {estado}")
        except Exception as e:
            print(f"❌ Error al llenar el campo Perfil Ocupacional: {str(e)}")
            logging.error(f"Error al llenar el campo Perfil Ocupacional: {str(e)}")
    except Exception as e:
        print(f"Error al llenar campos de Perfil Ocupacional: {str(e)}")


"""
def aceptar_mensaje_emergente(driver):
    Espera y hace clic en el botón 'Aceptar' de un mensaje emergente.
    try:
        wait = WebDriverWait(driver, 15)
        boton_aceptar = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//button[text()='Aceptar']")) #para aceptar el boton de respuesta de la pagina
        )
        print("✅ Se encontró el botón Aceptar en el mensaje emergente.")
        logging.info("Se encontró el botón Aceptar en el mensaje emergente.")
        boton_aceptar.click()
        print("✅ Se hizo clic en el botón Aceptar del mensaje emergente.")
        logging.info("Se hizo clic en el botón Aceptar del mensaje emergente.")
        time.sleep(2)  # Espera opcional después de hacer clic
        return True
    except Exception as e:
        print(f"❌ Error al interactuar con el mensaje emergente: {str(e)}")
        logging.error(f"Error al interactuar con el mensaje emergente: {str(e)}")
        return False
"""

def experiencia_laboral(driver, perfil_excel):
    try:
        # Primero verificamos si hay alguna alerta presente y la aceptamos
        try:
            WebDriverWait(driver, 1).until(EC.alert_is_present())
            alerta = driver.switch_to.alert
            mensaje_alerta = alerta.text
            print(f"✅ Se detectó una alerta: {mensaje_alerta}")
            logging.info(f"Se detectó una alerta: {mensaje_alerta}")
            alerta.accept()
            print("✅ Se aceptó la alerta")
            # Pequeña pausa para que la página se estabilice después de cerrar la alerta
            time.sleep(1)
        except:
            print("No hay alertas pendientes, continuando con el flujo normal")
        
        # Aumentar el tiempo de espera para asegurar que la página esté completamente cargada
        driver.implicitly_wait(5)
        
        # Esperar a que la página esté completamente cargada
        WebDriverWait(driver, 8).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        
        # Estrategia más robusta para localizar el tab de experiencia laboral
        print("Intentando localizar la pestaña de Experiencia Laboral...")
        
        # Intentar diferentes estrategias para localizar el elemento
        try:
            # Intento 1: XPath específico
            li_elemento_tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//li[@data-toggle="tab"]/a[contains(@href, "#tab5")]'))
            )
            print("✅ Tab localizado con XPath específico")
        except:
            try:
                # Intento 2: Buscar por el texto del enlace si está visible
                li_elemento_tab = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Experiencia") or contains(text(), "Laboral")]'))
                )
                print("✅ Tab localizado por texto del enlace")
            except:
                try:
                    # Intento 3: JavaScript para buscar por atributos
                    li_elemento_tab = driver.execute_script("""
                        return document.querySelector('li[data-toggle="tab"] a[href*="#tab5"]') || 
                               document.querySelector('a[href*="experiencia"], a[href*="laboral"]');
                    """)
                    print("✅ Tab localizado con JavaScript")
                except:
                    # Si todo falla, intentar hacer clic directamente en el elemento 'add-int'
                    print("⚠️ No se pudo localizar el tab. Intentando acceder directamente al formulario...")
                    try:
                        # Verificar si el elemento 'add-int' ya está disponible
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, 'add-int'))
                        )
                        print("✅ El formulario de intereses ya está accesible. Continuando...")
                        li_elemento_tab = None  # No necesitamos el tab
                    except:
                        raise Exception("No se pudo acceder al tab ni al formulario de intereses directamente")
        
        # Si se encontró el elemento tab, hacer clic en él
        if li_elemento_tab:
            # Ejecutar scroll hacia el elemento para asegurarnos de que sea visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", li_elemento_tab)
            time.sleep(1)  # Pequeña pausa para asegurar que el elemento esté visible
            
            # Intentar hacer clic de forma regular
            try:
                print("Intentando hacer clic en el tab...")
                li_elemento_tab.click()
                print("✅ Se hizo clic en el tab de forma normal")
            except:
                # Si falla, intentar con JavaScript
                try:
                    print("Intentando hacer clic con JavaScript...")
                    driver.execute_script("arguments[0].click();", li_elemento_tab)
                    print("✅ Se hizo clic en el tab usando JavaScript")
                except Exception as e:
                    print(f"❌ Error al hacer clic en el tab: {str(e)}")
            
            # Esperar un momento para que la interfaz responda
            time.sleep(2)
                
        # Resto del código para intereses ocupacionales
        try:
            print("Buscando el botón 'add-int'...")
            intereses_ocupacionales = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'add-int'))
            )
            intereses_ocupacionales.click()
            print("✅ Se hizo clic en el botón 'add-int'")
            # --- Intereses Ocupacionales ---
            print("Iniciando proceso de agregar interés ocupacional...")

            # Esperamos a que el contenedor de búsqueda esté visible
            try:
                contenedor = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, 'int-ocu-search-group'))
                )
                is_visible = driver.execute_script(
                    "return (arguments[0].offsetWidth > 0 || arguments[0].offsetHeight > 0) && window.getComputedStyle(arguments[0]).display !== 'none'",
                    contenedor
                )
                if not is_visible:
                    print("El contenedor 'int-ocu-search-group' no está visible, haciéndolo visible...")
                    driver.execute_script("""
                        var container = document.getElementById('int-ocu-search-group');
                        if (container) {
                            container.style.display = 'block';
                            container.classList.remove('hide');
                        }
                    """)
                    time.sleep(1)
                    print("✅ Se hizo visible el contenedor de búsqueda")
            except Exception as e:
                print(f"Advertencia al manipular el contenedor: {str(e)}")

            # Ahora buscar específicamente el campo de entrada con el XPath más preciso
            print("Buscando el campo de entrada 'ocu-nombre' dentro del formulario 'int-ocu-search'...")
            campo_nombre = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//form[@id='int-ocu-search']//input[@id='ocu-nombre']"))
            )
            print("✅ Campo 'ocu-nombre' encontrado y visible")

            # Agregar el valor como texto
            print(f"Intentando ingresar texto: '{perfil_excel}'")
            campo_nombre.send_keys(perfil_excel)
            time.sleep(1)  # Añade una espera después de enviar las teclas
            print("Buscando el botón de búsqueda 'ocu-nombre-btn'...")
            boton_busqueda = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//form[@id='int-ocu-search']//button[@id='ocu-nombre-btn']"))
            )
            boton_busqueda.click()
            driver.execute_script("arguments[0].click();", boton_busqueda)
            print("✅ Se hizo clic en el botón de búsqueda")
            time.sleep(2)

            # Buscar botones "Seleccionar" en los resultados
            print(f"Buscando el botón 'Seleccionar' para el perfil: '{perfil_excel}'...")
            try:
                # Primero intentar encontrar la fila que contiene el texto del perfil
                fila_perfil = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, 
                        f"//table[@id='int-ocu-table']/tbody/tr[td[contains(text(), '{perfil_excel}')]]"))
                )
                boton = fila_perfil.find_element(By.CSS_SELECTOR, "button.btn-primary")
                driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                time.sleep(0.5)  # Dar tiempo para que el scroll termine
                driver.execute_script("arguments[0].click();", boton)
                time.sleep(1)  # Espera a que se cargue el siguiente paso
                
                print(f"✅ Se hizo clic en el botón 'Seleccionar' para: '{perfil_excel}'")

                # --- Esperar a que aparezca el formulario de agregar interés y hacer clic en "Agregar" ---
                print("Esperando a que aparezca el formulario 'Agregar interés ocupacional'...")
                formulario_agregar = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, 'int-fadd-group'))
                )
                print("✅ Formulario 'Agregar interés ocupacional' visible")

                print("Buscando y haciendo clic en el botón 'Agregar'...")
                boton_agregar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'int-fadd-add'))
                )
                boton_agregar.click()
                print("✅ Se hizo clic en el botón 'Agregar'")
                time.sleep(3) # Espera a que se complete la adición

            except Exception as e:
                print(f"❌ No se pudo encontrar o hacer clic en el botón 'Seleccionar' para '{perfil_excel}': {e}")
                # Verificar que el interés se haya agregado correctamente
                try:
                    texto_encontrado = driver.find_elements(By.XPATH, f"//*[contains(text(), '{perfil_excel}')]")
                    if texto_encontrado:
                        print(f"✅ Verificado: El interés '{perfil_excel}' se agregó correctamente")
                        return True
                    else:
                        print(f"⚠️ No se encontró confirmación visual del interés agregado")
                        return False
                except:
                    print("No se pudo verificar la adición del interés")
                    return False
        except Exception as e:
            print(f"❌ Error general en intereses_ocupacionales: {str(e)}")
            traceback.print_exc()
            return False
    except Exception as e:
        print(f"❌ Error general en experiencia_laboral: {e}")
        # Capturar una captura de pantalla para depuración
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"error_screenshot_{timestamp}.png"
        try:
            driver.save_screenshot(screenshot_path)
            print(f"Se ha guardado una captura de pantalla en: {screenshot_path}")
        except Exception as screenshot_error:
            print(f"No se pudo guardar la captura de pantalla: {screenshot_error}")



def main():
    try:
        # Realizar login
        if not login():
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
                existe = verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos)
                
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
                    if llenar_datos_antes_de_inscripcion(nombres, apellidos):
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
                            llenar_formulario_telefono_correo(celular, correo)
                            llenar_input_perfil_ocupacional(estado)
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
                            experiencia_laboral(driver, perfil_ocupacional)
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
                    existe = verificar_estudiante(tipo_doc, num_doc, nombres, apellidos)
                    if not existe:
                        print("Reintentando llenar datos...")
                        if llenar_datos_antes_de_inscripcion(nombres, apellidos):
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