import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



# Funcion que llenara el numero de telefono y correo
def llenar_formulario_telefono_correo(telefono_excel, correo_excel, driver):
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