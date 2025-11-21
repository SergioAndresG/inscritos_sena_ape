import os
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC


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


def verificar_estudiante_con_CC_primero(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido=None):
    """
    Intenta verificar al estudiante primero con CC, y luego con el tipo de documento
    original solo si no se encontró con CC."""
    if tipo_doc == "CC":
        # Si ya es CC, verificamos directamente
        print(f"Verificando con CC: {num_doc}")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido)
    else:
        # Primero intentamos con CC
        print(f"Verificando primero con CC: {num_doc} (tipo original: {tipo_doc})")
        encontrado_con_cc = verificar_estudiante("CC", num_doc, nombres, apellidos, driver, wait, wait_rapido)
        
        # Si lo encontramos con CC, retornamos True
        if encontrado_con_cc is True:
            print(f"✓ Estudiante {num_doc} encontrado con CC, aunque el tipo original era {tipo_doc}")
            return True
            
        # Si la verificación con CC es None (error) o False (no encontrado),
        # intentamos con el tipo de documento original
        print(f"No se encontró con CC, intentando con tipo original {tipo_doc}")
        return verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido)

def verificar_estudiante(tipo_doc, num_doc, nombres, apellidos, driver, wait, wait_rapido=None):
    """Verifica si un estudiante ya está registrado en el sistema"""
    max_intentos = 3
    
    for intento in range(1, max_intentos + 1):
        try:
            print(f"Intento {intento}")
            driver.get(URL_VERIFICACION)
            
            # Esperar a que la página cargue completamente
            wait_rapido.until(EC.invisibility_of_element_located((By.ID, 'content-load')))
            wait_rapido.until(EC.visibility_of_element_located((By.ID, 'dropTipoIdentificacion')))
            print("Página cargada correctamente")
            
            logging.info(f"Verificando estudiante: {nombres} {apellidos} - Documento: {num_doc} - Tipo Doc: {tipo_doc}")
            
            
            # Seleccionar tipo de documento
            tipo_doc_dropdown = wait_rapido.until(EC.element_to_be_clickable((By.ID, 'dropTipoIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", tipo_doc_dropdown)
            print("Seleccionando tipo de documento...")
            
            # Crear el objeto Select
            selector = Select(tipo_doc_dropdown)
            # Seleccionar tipo de documento
            value_tipo_doc = TIPOS_DOCUMENTO.get(tipo_doc)

            if value_tipo_doc:
                selector.select_by_value(value_tipo_doc)
                print(f"Tipo de documento: {tipo_doc} (valor: {value_tipo_doc})")
            else:
                selector.select_by_visible_text(tipo_doc)
                print(f"Tipo de documento: {tipo_doc}")
                
            campo_num_id = wait.until(EC.element_to_be_clickable((By.ID, 'numeroIdentificacion')))
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_num_id)
            
            # Completar número de documento
            print("Ingresando número de documento...")
            campo_num_id.click()
            campo_num_id.clear()
            # Ingresar el documento
            campo_num_id.send_keys(str(num_doc))

            # Verificar que el valor se ingresó correctamente
            wait_rapido.until(lambda d: campo_num_id.get_attribute('value') == str(num_doc))
            print(f"Documento ingresado: {num_doc}")

            # Hacer clic en buscar
            print("Haciendo clic en buscar...")
            boton_buscar = wait_rapido.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='btnBuscar']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", boton_buscar)
            driver.execute_script("arguments[0].click();", boton_buscar)

            # Esperar que aparezca el loader y luego desaparezca
            try:
                wait_rapido.until(EC.visibility_of_element_located((By.ID, 'content-load')))
            except:
                pass

            wait_rapido.until(EC.invisibility_of_element_located((By.ID, 'content-load')))
            print("Esperando resultados...")
            
            try:
                # Esperar que la tabla tenga contenido REAL
                wait_rapido.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#bus-table_wrapper tbody tr"))
                )

                filas = driver.find_elements(By.CSS_SELECTOR, "#bus-table_wrapper tbody tr")
                        
                # Verificar que no sea una fila de "sin resultados"
                if len(filas) > 0:
                    primera_fila_texto = filas[0].text.lower() 
                    
                    if 'no se encontraron' in primera_fila_texto or 'sin resultados' in primera_fila_texto:
                        print(f"✗ NO ENCONTRADO: Tabla vacía o sin resultados para {num_doc}")
                        return False
                    
                    print(f"✓ ENCONTRADO: Tabla con {len(filas)} fila(s) de resultados")
                    return True
                else:
                    print(f"✗ NO ENCONTRADO: Tabla sin filas para {num_doc}")
                    return False
                    
            except Exception as e:
                print(f"NO ENCONTRADO")
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