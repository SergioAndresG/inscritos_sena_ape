import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options



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