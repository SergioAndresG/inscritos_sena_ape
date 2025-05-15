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
from URLS.urls import URL_LOGIN

# --- Credenciales de login desde variables de entorno ---
USUARIO_LOGIN = os.getenv('USUARIO_LOGIN')
CONTRASENA_LOGIN = os.getenv('CONTRASENA_LOGIN')
if not USUARIO_LOGIN or not CONTRASENA_LOGIN:
    error_msg = "Error: Las credenciales de login no están configuradas en el archivo .env"
    logging.error(error_msg)
    print(error_msg)
    exit()


# -- Funcion que Permite el Logueo del Funcionario APE (Agencia Publica de Empleo)
def login(driver, wait):
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