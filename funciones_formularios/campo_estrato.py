import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options



#funcion para llenar el estrato en el formulario principal
def llenar_estrato(driver):
    """Llena el campo de estrato"""
    try:   
        wait = WebDriverWait(driver, 10)
        
        estrato = wait.until(EC.presence_of_element_located((By.ID, 'estrato')))
        selector_suelddo = Select(estrato)
        selector_suelddo.select_by_visible_text('Dos')
        logging.info("Se seleccionó estrato Dos en el campo de Estrato.")
        print("Se seleccionó Dos en el campo de Estrato.")
        time.sleep(1)
        return True

    except Exception as e:
        error_msg = f"Error al llenar el formulario estrato: {str(e)}"
        logging.error(error_msg)
        print(f"❌ {error_msg}")
        return False