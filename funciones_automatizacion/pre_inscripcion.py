import time
import os
import re
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


dir_screenchost = "Screensshot"
os.makedirs(dir_screenchost, exist_ok=True)


def limpiar_texto(texto):
    """Normaliza el texto eliminando espacios múltiples y conservando un solo espacio entre palabras."""
    if not isinstance(texto, str):
        return ""
    return ' '.join(texto.strip().split())

def limpiar_alfanumerico(texto):
    """Elimina espacios y símbolos, dejando solo caracteres alfanuméricos."""
    if not isinstance(texto, str):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', texto)

def llenar_datos_antes_de_inscripcion(nombres_excel, apellidos_excel, driver):
    """Llena los campos de nombres, apellidos, fecha de nacimiento y género antes de la inscripción exhaustiva."""
    try:
        # Imprimos los datos que se van a llenra
        logging.info(f"Intentando llenar los campos antes de la inscripción con: {nombres_excel} {apellidos_excel}")
        
        # Esperar a que la página de pre-inscripción cargue completamente
        print("Esperando que el formulario de pre-inscripción cargue...")
        
        # Capturar screenshot antes de interactuar con el formulario
        ruta_completa = os.path.join(dir_screenchost,f"pre_inscripcion_{int(time.time())}.png")
        driver.save_screenshot(ruta_completa)
        
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
            campo_nombres_pre.send_keys(nombres_excel)

                
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
                
            print(f"Se llenó el campo Primer apellido con: {primer_apellido}")
            logging.info(f"Se llenó el campo Primer apellido con: {primer_apellido}")
        except Exception as e:
            print(f"Error al llenar el campo Primer apellido: {str(e)}")
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
                
            print(f"Se llenó el campo Segundo apellido con: {segundo_apellido}")
            logging.info(f"Se llenó el campo Segundo apellido con: {segundo_apellido}")
        except Exception as e:
            print(f"Error al llenar el campo Segundo apellido: {str(e)}")
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
            
            print("Se llenó el campo Fecha de Nacimiento")
            logging.info("Se llenó el campo Fecha de Nacimiento")
        except Exception as e:
            print(f"Error al llenar el campo Fecha de Nacimiento: {str(e)}")
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
                            
                # Hacer scroll hacia el elemento
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_genero)
                time.sleep(0.5)

                # Luego crea el objeto Select usando el elemento
                selector_genero = Select(elemento_genero)
                selector_genero.select_by_value(genero)
                time.sleep(0.5)
                
                selector_genero.select_by_value(genero)
                time.sleep(1)
                
                print(f"Se seleccionó el género: {genero} ({'Femenino' if genero == '1' else 'Masculino'})")
                logging.info(f"Se seleccionó el género: {genero} ({'Femenino' if genero == '1' else 'Masculino'})")
            else:
                print("No se pudo determinar el género, intentando seleccionar por defecto Masculino")
                selector_genero = Select(WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.NAME, 'genero'))
                ))
                selector_genero.select_by_value('0')  # Masculino por defecto
                logging.warning(f"No se pudo determinar el género para: {nombres_excel}, se seleccionó Masculino por defecto")
        except Exception as e:
            print(f"Error al seleccionar el género: {str(e)}")
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
                    print("Se hizo clic en el botón para verificar")
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
                            print("Se hizo clic en el botón para continuar/inscribir")
                            logging.info("Se hizo clic en el botón para continuar/inscribir")
                        else:
                            print("El botón para continuar/inscribir no está visible")
                    except Exception as e:
                        print(f"Error al buscar/hacer clic en el botón continuar/inscribir: {str(e)}")
                        logging.error(f"Error al buscar/hacer clic en el botón continuar/inscribir: {str(e)}")
                else:
                    print("El botón para verificar no está visible")
            except Exception as e:
                print(f"Error al buscar/hacer clic en el botón verificar: {str(e)}")
                logging.error(f"Error al buscar/hacer clic en el botón verificar: {str(e)}")

        except Exception as e:
            print(f"Error general al manejar los botones: {str(e)}")
            logging.error(f"Error general al manejar los botones: {str(e)}")
        return True
    except Exception as e:
        logging.error(f"Error general al llenar los campos antes de la inscripción: {str(e)}")
        print(f"Error general al llenar los campos antes de inscripción: {str(e)}")


def determinar_genero(nombre):
    """Intenta determinar el género a partir del nombre."""
    nombre = nombre.lower()
    
    # Nombres Femeninos
    nombres_femeninos = [
        'ana', 'maria', 'sofia', 'isabella', 'valentina', 'camila', 'mariana', 'laura',
        'daniela', 'valeria', 'liz', 'ariana', 'lizeth', 'danna', 'paula', 'juliana',
        'karla', 'alejandra', 'fernanda', 'samantha', 'antonella', 'lucia', 'martina',
        'renata', 'ximena', 'gabriela', 'carolina', 'viviana', 'melany', 'abigail',
        'regina', 'andrea', 'joselyn', 'salome', 'emilia', 'angelica', 'micaela',
        'jazmin', 'alison', 'karen', 'rosario', 'estrella', 'araceli', 'joana',
        'milena', 'tania', 'yaqueline', 'noelia', 'vanessa', 'brenda',
        'helena', 'victoria', 'adriana', 'irene', 'elena', 'claudia', 'erika', 'natalia',
        'giselle', 'rocío', 'verónica', 'elisa', 'cristina', 'patricia', 'eugenia',
        'amanda', 'celia', 'ines', 'monica', 'beatriz', 'linda', 'mercedes', 'doris',
        'alicia', 'berta', 'cecilia', 'diana', 'gloria', 'hortensia', 'juana', 'martha',
        'rebeca', 'teresa', 'yolanda', 'elizabeth', 'isabel', 'aisha', 'penelope',
        'catalina', 'alondra', 'esmeralda', 'delfina', 'maite', 'carmen', 'pilar',
        'consuelo', 'esperanza', 'aura', 'iris', 'grecia', 'kiara', 'mayra', 'nayeli',
        'bianca', 'ciara', 'zoe', 'luna', 'chloe', 'sara', 'eva', 'mía', 'nora',
        'olivia', 'vera', 'maya', 'gala', 'dalia', 'kendra', 'kelly', 'kimberly',
        'leslie', 'lorena', 'margarita', 'nadia', 'nicole', 'paloma', 'sandra',
        'alana', 'alma', 'ambra', 'amira', 'anahi', 'arlet', 'belén', 'briana',
        'brisa', 'dayana', 'denisse', 'érica', 'fátima', 'fiora', 'frida', 'hayde',
        'indira', 'jana', 'kira', 'leila', 'liam', 'lisandra', 'lissette', 'lourdes',
        'maribel', 'marlene', 'melina', 'milenka', 'miranda', 'nerea', 'novia',
        'pamela', 'quinn', 'rocio', 'rosa', 'ruth', 'sibel', 'silvana', 'stella',
        'tatiana', 'thais', 'uma', 'wendy', 'yenne', 'zaira', 'zara', 'zulma',
        'bertha', 'cleopatra', 'eunice', 'felipa', 'genoveva', 'hilda', 'isolda',
        'jacinta', 'katiuska', 'lidia', 'matilda', 'norma', 'ofelia', 'prisca',
        'quiteria', 'salma', 'ursula', 'zenaida', 'arlette', 'cynthia', 'débora',
        'evelyn', 'gema', 'ivonne', 'jessica', 'kassandra', 'liliana', 'michell',
        'neida', 'rossana', 'shelby', 'susana', 'valeria', 'yessenia', 'zulema',
        'aileen', 'anya', 'aylin', 'betsy', 'danna', 'eliza', 'gemma', 'hanna',
        'ivana', 'julieta', 'katia', 'leona', 'maia', 'nala', 'perla', 'sirena',
        'tina', 'zoila', 'adela', 'amalia', 'apolonia', 'carmela', 'dominga',
        'esperanza', 'felicidad', 'graciela', 'hortensia', 'judith', 'kristel',
        'magdalena', 'minerva', 'noemi', 'petra', 'romina', 'sabrina', 'ximena',
        'yanet', 'zelia', 'ximena', 'yara', 'zafiro', 'adina', 'agatha', 'ariel',
        'ayla', 'bárbara', 'clara', 'dánae', 'dorotea', 'edith', 'eugenia',
        'francisca', 'heidi', 'iliana', 'iris', 'jana', 'kora', 'luisa',
        'mabel', 'nadine', 'odette', 'paz', 'rita', 'talia', 'ursula', 'vicky',
        'wanda', 'yaretzi', 'zoraida', 'annette', 'dina', 'glenda', 'karenza',
        'leidy', 'marina', 'roxana', 'sharom', 'ximena', 'yuliana', 'zianya'
    ]

    # Nombres Masculinos
    nombres_masculinos = [
        'juan', 'carlos', 'luis', 'andres', 'sebastian', 'mateo', 'santiago', 'alejandro',
        'daniel', 'gabriel', 'miguel', 'jose', 'diego', 'tomás', 'manuel', 'antonio',
        'adrian', 'julian', 'felipe', 'fernando', 'ricardo', 'rafael', 'pedro', 'joel',
        'nicolas', 'emiliano', 'marcos', 'david', 'lucas', 'cristian', 'axel', 'isaac',
        'eduardo', 'hugo', 'benjamin', 'matias', 'victor', 'francisco', 'gael',
        'esteban', 'jeronimo', 'ian', 'rodrigo', 'bryan', 'elias', 'mauricio', 'raul',
        'alvaro', 'omar', 'julio', 'alberto', 'arturo', 'javier', 'sergio', 'oscar', 
        'héctor', 'enrique', 'pablo',
        'ramón', 'roberto', 'alfonso', 'camilo', 'dario', 'erick', 'gonzalo', 'guillermo',
        'horacio', 'ignacio', 'jaime', 'jorge', 'leonardo', 'marcelo', 'néstor', 'oreste',
        'pascual', 'ramiro', 'rené', 'rubén', 'salvador', 'teodoro', 'ulises', 'vicente',
        'walter', 'xavier', 'yair', 'abel', 'alan', 'benito', 'braulio', 'césar',
        'demian', 'denis', 'efrén', 'fabian', 'gregorio', 'iker', 'israel', 'jesús',
        'kevin', 'leo', 'maximiliano', 'neil', 'oliver', 'pablo', 'quique', 'samuel',
        'tito', 'toni', 'troy', 'unai', 'yago', 'zenón', 'saúl', 'josué', 'noé',
        'amadeo', 'apolonio', 'basilio', 'cipriano', 'cleto', 'domenico', 'fabio',
        'federico', 'hernan', 'leonel',
        'aaron', 'abraham', 'adán', 'agustín', 'alex', 'alfredo', 'amaro', 'amir',
        'anibal', 'ariel', 'armando', 'bernal', 'boris', 'bruno', 'calixto', 'cosme',
        'damián', 'eloy', 'ermilo', 'esau', 'eugenio', 'ezequiel', 'fabrizio',
        'fausto', 'florian', 'gervasio', 'gilberto', 'guido', 'humberto', 'iker',
        'isaías', 'ismael', 'iván', 'jacob', 'jairo', 'jeremías', 'jonathan',
        'lorenzo', 'luciano', 'macario', 'máximo', 'melchor', 'montserrat', 'moses',
        'nataniel', 'octavio', 'orlando', 'pascal', 'patricio', 'pío', 'plinio',
        'policarpo', 'próspero', 'quintín', 'quiroz', 'raimundo', 'rodolfo',
        'roquefort', 'simón', 'taimur', 'teófilo', 'tristán', 'ulises', 'valerio',
        'venancio', 'wenceslao', 'wilfredo', 'yunes', 'zacarías', 'zian', 'abdon',
        'agapito', 'alano', 'anacleto', 'antelmo', 'baltasar', 'benedicto', 'blanco',
        'ceferino', 'conrado', 'doroteo', 'edmundo', 'emilio', 'epifanio', 'erasto',
        'galileo', 'gandalf', 'german', 'gisleno', 'godofrido', 'hipólito', 'hilario',
        'homer', 'isidro', 'jofre', 'leandro', 'león', 'luther', 'martín',
        'melitón', 'modesto', 'nahuel', 'nicandro', 'nicanor', 'nilo', 'nubiel',
        'pascual', 'pietro', 'primitivo', 'rafael', 'reinaldo', 'romualdo', 'rufo',
        'sabiniano', 'silvano', 'silverio', 'tadeo', 'teobaldo', 'tibercio', 'toribio',
        'umberto', 'valente', 'vito', 'willy', 'yaser', 'yusef', 'zebulón',
        'aurelio', 'casimiro', 'celestino', 'dionisio', 'eleuterio', 'evangelista',
        'florentino', 'genaro', 'hilario', 'leopoldo', 'napoleón', 'raimundo',
        'silvestre', 'valeriano', 'vladimir', 'anaximandro', 'aristóteles', 'dante',
        'platón', 'sócrates', 'virgilio', 'zenon', 'adriel', 'bautista', 'darío',
        'ezequiel', 'ian', 'javier', 'liam', 'milán', 'nehemías', 'osiel', 'pablo',
        'thiago', 'uziel', 'york', 'zair'
    ]

    primer_nombre = nombre.split(' ')[0] # Tomamos el primer nombre

    if primer_nombre in nombres_femeninos:
        return '1'  # En el aplicativo 1 es femenino
    elif primer_nombre in nombres_masculinos:
        return '0'  # En el aplicativo 0 es Masculino
    else:
        return None # No se pudo determinar el género