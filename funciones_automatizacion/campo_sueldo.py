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