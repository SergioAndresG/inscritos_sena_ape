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


URL_VERIFICACION = 'https://agenciapublicadeempleo.sena.edu.co/spe-web/spe/funcionario/oferta'

# Mapeo de tipos de documento
TIPOS_DOCUMENTO = {
    "CC": "1",
    "TI": "2", 
    "CE": "3",
    "Otro Documento de Identidad": "5",
    "PEP": "8",
    "PPT": "9"
}


def verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos, driver, wait):
    """
    Intenta verificar al estudiante primero con CC, y luego con el tipo de documento
    original solo si no se encontró con CC."""
    if tipo_doc == "CC":
        # Si ya es CC, verificamos directamente
        print("Tipo de documento ya es CC, verificando directamente...")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait)
    else:
        # Primero intentamos con CC
        print(f"Intentando verificar primero con CC para documento {num_doc}...")
        encontrado_con_cc = verificar_estudiante("CC", num_doc, nombres, apellidos, driver, wait)
        
        # Si lo encontramos con CC, retornamos True
        if encontrado_con_cc is True:
            print(f"✅ Estudiante {num_doc} encontrado con CC, aunque el tipo original era {tipo_doc}")
            return True
            
        # Si la verificación con CC es None (error) o False (no encontrado),
        # intentamos con el tipo de documento original
        print(f"No se encontró con CC, intentando con tipo original {tipo_doc}")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait)

def verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait):
    """Verifica si un estudiante ya está registrado en el sistema"""
    max_intentos = 3
    
    for intento in range(1, max_intentos + 1):
        try:
            print(f"Intento {intento} - Abriendo URL: {URL_VERIFICACION}")
            driver.get(URL_VERIFICACION)
            
            # Esperar a que la página cargue completamente
            print("Esperando que la página cargue...")
            wait.until(EC.invisibility_of_element_located((By.ID, 'content-load')))
            wait.until(EC.visibility_of_element_located((By.ID, 'dropTipoIdentificacion')))
            print("Página cargada correctamente")
            
            logging.info(f"Verificando estudiante: {nombres} {apellidos} - Documento: {num_doc} - Tipo Doc: {tipo_doc}")
            
            
            # Seleccionar tipo de documento
            tipo_doc_dropdown = wait.until(EC.element_to_be_clickable((By.ID, 'dropTipoIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", tipo_doc_dropdown)
            print("Seleccionando tipo de documento...")
            
            # Crear el objeto Select
            selector = Select(tipo_doc_dropdown)
            # Seleccionar tipo de documento
            value_tipo_doc = TIPOS_DOCUMENTO.get(tipo_doc)

            if value_tipo_doc:
                print(f"Seleccionando valor: {value_tipo_doc}")
                selector.select_by_value(value_tipo_doc)
            else:
                print(f"Seleccionando texto visible: {tipo_doc}")
                selector.select_by_visible_text(tipo_doc)
                
            campo_num_id = wait.until(EC.element_to_be_clickable((By.ID, 'numeroIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_num_id)
            
            # Completar número de documento
            print("Ingresando número de documento...")
            campo_num_id.click()

            wait.until(lambda d: campo_num_id.get_attribute('value') == '' or True)
            campo_num_id.clear()
            # Ingresar el documento
            campo_num_id.send_keys(str(num_doc))

            wait.until(lambda d: campo_num_id.get_attribute('value') == '' or str(num_doc))
            print(f"Documento ingresado: {num_doc}")
            

            # Hacer clic en buscar con JavaScript
            print("Haciendo clic en buscar...")
            boton_buscar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@id='btnBuscar']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", boton_buscar)
            driver.execute_script("arguments[0].click();", boton_buscar)
            
           # Esperar que aparezca el loader y luego desaparezca
            try:
                wait.until(EC.visibility_of_element_located((By.ID, 'content-load')))
            except:
                pass  # Si no aparece el loader, continuar
            
            wait.until(EC.invisibility_of_element_located((By.ID, 'content-load')))
            print("Esperando resultados...")

            # Esperar que aparezca algún contenido (tabla o formulario)
            wait.until(lambda d: len(d.find_elements(By.TAG_NAME, 'table')) > 0 or 
                                len(d.find_elements(By.NAME, 'nombres')) > 0)

            # VERIFICAR LA EXSITENCIA DEL USUARIO
            # 1. Verificar si hay una tabla de resultados con datos
            tablas = driver.find_elements(By.TAG_NAME, "table")
            for tabla in tablas:
                if tabla.is_displayed():
                    filas = tabla.find_elements(By.TAG_NAME, "tr")
                    if filas and len(filas) > 1:  # Si hay más de una fila (encabezado + datos)
                        print(f"✅ ENCONTRADO: Tabla con {len(filas)} filas tiene resultados")
                        return True
                    
            # 2. Verificar si aparece un formulario para completar datos (caso en que no existe)
            try:
                # Buscar campos típicos del formulario de pre-inscripción
                campos_formulario = driver.find_elements(By.CSS_SELECTOR, 'input[name="nombres], input[name="primerApellido"]')
                
                if len(campos_formulario) >= 2 and all(campo.is_displayed() for campo in campos_formulario):
                    print(f"❌ NO ENCONTRADO: Se muestra formulario para completar datos de {num_doc}")
                    return False
            except:
                pass
            
            # 3. Verificar mensajes explícitos
            page_source = driver.page_source.lower()
            if "no se encontraron resultados" in page_source:
                print(f"❌ NO ENCONTRADO: No hay resultados para el estudiante {num_doc}")
                return False
                
            # 4. Buscar indicadores positivos
            if any(indicator in page_source for indicator in ["ya existe", "encontrado", str(num_doc)]):
                print(f"✅ ENCONTRADO: Indicadores sugieren que el estudiante {num_doc} existe")
                return True
            
            # Si no se pudo determinar claramente, verificar la URL actual
            # Generalmente cuando no existe el estudiante, la página redirige a un formulario
            if "/inscripcion" in driver.current_url or "/register" in driver.current_url:
                print(f"❌ NO ENCONTRADO: Redirigido a formulario para {num_doc}")
                return False
                
            # Por defecto, asumir que no existe si no hay evidencia clara
            print(f"❓ INDETERMINADO: No hay evidencia clara sobre {num_doc}, asumiendo que no existe")
            return False
                
        except Exception as e:
            logging.error(f"Error en intento {intento}: {str(e)}")
            print(f"Error en verificación (intento {intento}): {str(e)}")
            
            if intento == max_intentos:
                print("Se agotaron los intentos. No se pudo verificar el estudiante.")
                return None
            else:
                print(f"Reintentando... ({intento}/{max_intentos})")
                time.sleep(3)
    
    return None