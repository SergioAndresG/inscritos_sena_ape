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
def llenar_formulario_ubicacion_residencia(driver):
    """Llena los campos de País, Departamento y Municipio en el formulario principal. (Residencia)"""
    try:
        wait = WebDriverWait(driver, 10)
        # Seleccionar País: Colombia
        pais = wait.until(EC.presence_of_element_located((By.ID, 'pais')))
        selector_pais = Select(pais)
        selector_pais.select_by_visible_text('Colombia')
        logging.info("Se seleccionó Colombia en el desplegable de País.")
        print("✓ Se seleccionó Colombia en el desplegable de País.")
        time.sleep(1)  # Esperar a que se cargue el departamento

        # Seleccionar Departamento: Cundinamarca
        departamento = wait.until(EC.presence_of_element_located((By.ID, 'departamento')))
        selector_departamento = Select(departamento)
        selector_departamento.select_by_visible_text('Cundinamarca')
        logging.info("Se seleccionó Cundinamarca en el desplegable de Departamento.")
        print("✓ Se seleccionó Cundinamarca en el desplegable de Departamento.")
        time.sleep(1)  # Esperar a que se cargue el municipio

        # Seleccionar Municipio: Mosquera
        municipio = wait.until(EC.presence_of_element_located((By.ID, 'municipio')))
        selector_municipio = Select(municipio)
        selector_municipio.select_by_visible_text('Mosquera')
        logging.info("Se seleccionó Mosquera en el desplegable de Municipio.")
        print("✓ Se seleccionó Mosquera en el desplegable de Municipio.")

        direccion_residencia_input = wait.until(EC.presence_of_element_located((By.ID, 'residenciaDireccion')))
        direccion_residencia_input.send_keys('MOSQUERA')
        logging.info("Se coloco Mosquea en el input")
        print("✓ Mosquera ingresado en el input")

        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario principal (residencia): {str(e)}"
        logging.error(error_msg)
        print(f"✗ {error_msg}")
        return False
    
