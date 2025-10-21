import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options



def llenar_formulario_ubicaciones_nacimiento(driver,wait):
    """Llena los campos de ubicación de Nacimiento evitando conflictos con el primer formulario."""
    try:
        
        print("⚪ Ingresando Datos de Nacimiento")
        
        # 1. Hacer clic en el botón "Seleccionar" para el municipio de nacimiento
        # Usar un selector mucho más específico
        seleccionar_boton = wait.until(
            EC.element_to_be_clickable((By.XPATH, 
                "//*[@id='municipioNacimientoText']/following-sibling::div//button[contains(@class, 'dropdown-toggle')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", seleccionar_boton)
        seleccionar_boton.click()
        print("✅ Se hizo clic en el botón Seleccionar para datos de nacimiento")
        time.sleep(1)
        
        # 2. Verificar que el dropdown se haya abierto
        dropdown = wait.until(
            EC.visibility_of_element_located((By.ID, "dropNacimiento"))
        )
        print("✅ Se abrió el dropdown de nacimiento correctamente")
        
        # 3. Importante: CAMBIAR AL CONTEXTO DEL DROPDOWN
        # En vez de buscar por ID directamente (lo cual puede encontrar elementos duplicados),
        # vamos a buscar elementos dentro del dropdown específico
        
        # Seleccionar País
        # Usamos XPath para encontrar el select dentro del dropdown específico
        pais_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][1]//select"))
        )
        Select(pais_select).select_by_visible_text('Colombia')
        print("✅ Se seleccionó Colombia en el desplegable de País (Nacimiento).")
        time.sleep(1.5)  # Esperar más tiempo para que se carguen opciones
        
        # Seleccionar Departamento
        # Nuevamente, usamos XPath para encontrar el segundo select dentro del dropdown
        depto_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][2]//select"))
        )
        Select(depto_select).select_by_visible_text('Cundinamarca')
        print("✅ Se seleccionó Cundinamarca en el desplegable de Departamento (Nacimiento).")
        time.sleep(1.5)
        
        # Seleccionar Municipio
        # Y el tercer select dentro del dropdown
        muni_select = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='dropNacimiento']//div[contains(@class,'control-group')][3]//select"))
        )
        Select(muni_select).select_by_visible_text('Mosquera')
        print("✅ Se seleccionó Mosquera en el desplegable de Municipio (Nacimiento).")
        
        # 4. Cerrar el dropdown haciendo clic en algún otro elemento para confirmar la selección
        try:
            fecha_input = driver.find_element(By.ID, "fechaNacimiento")
            fecha_input.click()
            print("✅ Se cerró el dropdown haciendo clic fuera")
        except:
            # Si el input de fecha no está disponible, intenta hacer clic en cualquier otro elemento
            try:
                # Buscar un elemento visible y hacer clic en él
                driver.find_element(By.XPATH, "//h4[contains(text(), 'Datos de nacimiento')]").click()
                print("✅ Se cerró el dropdown haciendo clic en el título")
            except:
                print("No se pudo cerrar el dropdown manualmente")
        
        return True
        
    except Exception as e:
        print(f"❌ Error al llenar la ubicación de Nacimiento: {e}")
        logging.error(f"Error detallado al llenar la ubicación de Nacimiento: {str(e)}")
        
        # Información adicional de depuración
        try:
            # Verificar si el dropdown está abierto
            dropdown_visible = len(driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento:not([style*='display: none'])")) > 0
            print(f"¿Dropdown visible? {dropdown_visible}")
            
            if dropdown_visible:
                # Contar cuántos divs de control-group hay dentro del dropdown
                control_groups = driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento .control-group")
                print(f"Grupos de control encontrados: {len(control_groups)}")
                
                # Contar cuántos selects hay dentro del dropdown
                selects = driver.find_elements(By.CSS_SELECTOR, "#dropNacimiento select")
                print(f"Selects encontrados dentro del dropdown: {len(selects)}")
                
                # Ver qué opciones hay disponibles en el primer select
                if len(selects) > 0:
                    options = selects[0].find_elements(By.TAG_NAME, "option")
                    option_texts = [opt.text for opt in options]
                    print(f"Opciones en el primer select: {option_texts}")
            
        except Exception as debug_e:
            print(f"Error en depuración: {debug_e}")
        
        return False