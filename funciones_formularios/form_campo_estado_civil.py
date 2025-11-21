import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

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
        print("✓ Se seleccionó Soltero en el desplegable de Estado Civil.")
        time.sleep(1)  # Esperar a que se cargue el estado civil

        # Seleccionar Cantidad de Personas a Cargo: 0
        personas_a_cargo_input = wait.until(EC.presence_of_element_located((By.ID, 'personasCargo')))
        personas_a_cargo_input.send_keys('0')
        logging.info("Se ingreso 0 en el input de Personas a Cargo.")
        print("✓ Se ingreso 0 en el input de Personas a Cargo.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Posicion Familiar: Otro
        posicion_familiar = wait.until(EC.presence_of_element_located((By.ID, 'posicionFamiliar')))
        selector_familiar = Select(posicion_familiar)
        selector_familiar.select_by_visible_text('Otro')
        logging.info("Se seleccionó Otro en el desplegable de Posicion Familiar.")
        print("✓ Se seleccionó Otro en el desplegable de Posicion Familiar.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Libreta Militar: No tiene libreta militar
        libreta_militar = wait.until(EC.presence_of_element_located((By.ID, 'libretaMilitar')))
        selector_militar = Select(libreta_militar)
        selector_militar.select_by_visible_text('No tiene libreta militar')
        logging.info("Se seleccionó 'No tiene libreta militar' en el desplegable de libreta militar.")
        print("✓ Se seleccionó 'No tiene libreta militar' en el desplegable de libreta militar.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar Migrante: No
        personas_a_cargo_input = wait.until(EC.presence_of_element_located((By.ID, 'migrante')))
        personas_a_cargo_input.send_keys('No')
        logging.info("Se ingreso 'No' en el input de Migrante.")
        print("✓ Se ingreso 'No' en el input de Migrante.")
        time.sleep(1)  # Esperar a que se cargue

        # Seleccionar en Zona de residencia: Urbana
        zona_residencia = wait.until(EC.presence_of_element_located((By.ID, 'esRural')))
        selector_zona_residencia = Select(zona_residencia)
        selector_zona_residencia.select_by_visible_text('Urbana')
        logging.info("Se seleccionó 'Urbano' en el desplegable de zona de resindecia.")
        print("✓ Se seleccionó 'Urbano' en el desplegable de zona de resindecia.")
        time.sleep(1)  # Esperar a que se cargue
        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario principal (residencia): {str(e)}"
        logging.error(error_msg)
        print(f"✗ {error_msg}")
        return False
