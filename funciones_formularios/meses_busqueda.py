import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


def verificar_meses_busqueda(driver):
    """
    Verifica si el campo mesesBusqueda tiene valor mayor a 1
    Retorna True si ya está registrado (valor > 1), False si no está registrado
    """
    try:
        
        wait = WebDriverWait(driver, 10)
        
        # Buscar el campo mesesBusqueda
        campo_meses = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "mesesBusqueda"))
        )
        
        # Obtener el valor actual del campo
        valor_meses = campo_meses.get_attribute("value")
        
        # Convertir a número y verificar
        if valor_meses and valor_meses.isdigit():
            meses_valor = int(valor_meses)
            print(f"Valor actual de mesesBusqueda: {meses_valor}")
            
            if meses_valor > 1:
                print("✓ El estudiante ya está registrado (mesesBusqueda > 1)")
                return True
            else:
                print(" El estudiante no está registrado (mesesBusqueda <= 1)")
                return False
        else:
            print("Campo mesesBusqueda no tiene valor válido")
            return False
            
    except Exception as e:
        print(f"✗ Error al verificar mesesBusqueda: {str(e)}")
        return False