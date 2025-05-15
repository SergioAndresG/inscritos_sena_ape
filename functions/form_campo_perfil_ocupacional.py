import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options



# Funcion que llenara el numero de telefono y correo
def llenar_input_perfil_ocupacional(estado,driver):
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
