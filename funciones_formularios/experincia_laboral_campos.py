import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import traceback
from datetime import datetime

def experiencia_laboral(driver, perfil_excel):
    try:
        # Primero verificamos si hay alguna alerta presente y la aceptamos
        try:
            alerta = WebDriverWait(driver, 2).until(EC.alert_is_present())
            texto_alerta = alerta.text
            print("Alerta detectada")
            alerta.accept() # Aceptamos la alerta
            print("Alerta aceptada")
            time.sleep(1)

            # Clasificar el tipo de alerta
            texto_lower = texto_alerta.lower()

            if "error" in texto_lower or "intente nuevamente" in texto_lower or "ha ocurrido" in texto_lower:
                logging.error(f"Alerta de error detectada: {texto_alerta}")
                print(f"❌ ERROR DETECTADO EN ALERTA: {texto_alerta}")
                return False  # Error crítico
                
            elif "guardados correctamente" in texto_lower or "éxito" in texto_lower:
                logging.info(f"Alerta de éxito: {texto_alerta}")
                print(f"✅ Datos guardados correctamente")
                # Continuar normalmente
            else:
                # Alerta desconocida, por precaución tratarla como advertencia
                logging.warning(f"Alerta desconocida: {texto_alerta}")
                print(f"⚠️ Alerta no reconocida: {texto_alerta}")
        except:
            print("No hay alertas pendientes, continuando con el flujo normal")
        
        # Aumentar el tiempo de espera para asegurar que la página esté completamente cargada
        driver.implicitly_wait(5)
        
        # Esperar a que la página esté completamente cargada
        WebDriverWait(driver, 8).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        
        # Estrategia más robusta para localizar el tab de experiencia laboral
        print("Intentando localizar la pestaña de Experiencia Laboral...")
        
        # Intentar diferentes estrategias para localizar el elemento
        try:
            # Intento 1: XPath específico
            li_elemento_tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//li[@data-toggle="tab"]/a[contains(@href, "#tab5")]'))
            )
            print("✅ Tab localizado con XPath específico")
        except:
            try:
                # Intento 2: Buscar por el texto del enlace si está visible
                li_elemento_tab = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "Experiencia") or contains(text(), "Laboral")]'))
                )
                print("✅ Tab localizado por texto del enlace")
            except:
                try:
                    # Intento 3: JavaScript para buscar por atributos
                    li_elemento_tab = driver.execute_script("""
                        return document.querySelector('li[data-toggle="tab"] a[href*="#tab5"]') || 
                               document.querySelector('a[href*="experiencia"], a[href*="laboral"]');
                    """)
                    print("✅ Tab localizado con JavaScript")
                except:
                    # Si todo falla, intentar hacer clic directamente en el elemento 'add-int'
                    print("⚠️ No se pudo localizar el tab. Intentando acceder directamente al formulario...")
                    try:
                        # Verificar si el elemento 'add-int' ya está disponible
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, 'add-int'))
                        )
                        print("✅ El formulario de intereses ya está accesible. Continuando...")
                        li_elemento_tab = None  # No necesitamos el tab
                    except:
                        raise Exception("No se pudo acceder al tab ni al formulario de intereses directamente")
        
        # Si se encontró el elemento tab, hacer clic en él
        if li_elemento_tab:
            # Ejecutar scroll hacia el elemento para asegurarnos de que sea visible
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", li_elemento_tab)
            time.sleep(1)  # Pequeña pausa para asegurar que el elemento esté visible
            
            # Intentar hacer clic de forma regular
            try:
                print("Intentando hacer clic en el tab...")
                li_elemento_tab.click()
                print("✅ Se hizo clic en el tab de forma normal")
            except:
                # Si falla, intentar con JavaScript
                try:
                    print("Intentando hacer clic con JavaScript...")
                    driver.execute_script("arguments[0].click();", li_elemento_tab)
                    print("✅ Se hizo clic en el tab usando JavaScript")
                except Exception as e:
                    print(f"❌ Error al hacer clic en el tab: {str(e)}")
            
            # Esperar un momento para que la interfaz responda
            time.sleep(2)
                
        # Resto del código para intereses ocupacionales
        try:
            print("Buscando el botón 'add-int'...")
            intereses_ocupacionales = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'add-int'))
            )
            intereses_ocupacionales.click()
            print("✅ Se hizo clic en el botón 'add-int'")
            # --- Intereses Ocupacionales ---
            print("Iniciando proceso de agregar interés ocupacional...")

            # Esperamos a que el contenedor de búsqueda esté visible
            try:
                contenedor = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, 'int-ocu-search-group'))
                )
                is_visible = driver.execute_script(
                    "return (arguments[0].offsetWidth > 0 || arguments[0].offsetHeight > 0) && window.getComputedStyle(arguments[0]).display !== 'none'",
                    contenedor
                )
                if not is_visible:
                    print("El contenedor 'int-ocu-search-group' no está visible, haciéndolo visible...")
                    driver.execute_script("""
                        var container = document.getElementById('int-ocu-search-group');
                        if (container) {
                            container.style.display = 'block';
                            container.classList.remove('hide');
                        }
                    """)
                    time.sleep(1)
                    print("✅ Se hizo visible el contenedor de búsqueda")
            except Exception as e:
                print(f"Advertencia al manipular el contenedor: {str(e)}")

            # Ahora buscar específicamente el campo de entrada con el XPath más preciso
            print("Buscando el campo de entrada 'ocu-nombre' dentro del formulario 'int-ocu-search'...")
            campo_nombre = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//form[@id='int-ocu-search']//input[@id='ocu-nombre']"))
            )
            print("✅ Campo 'ocu-nombre' encontrado y visible")

            # Agregar el valor como texto
            print(f"Intentando ingresar texto: '{perfil_excel}'")
            campo_nombre.send_keys(perfil_excel)
            time.sleep(1)  # Añade una espera después de enviar las teclas
            print("Buscando el botón de búsqueda 'ocu-nombre-btn'...")
            boton_busqueda = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//form[@id='int-ocu-search']//button[@id='ocu-nombre-btn']"))
            )
            boton_busqueda.click()
            driver.execute_script("arguments[0].click();", boton_busqueda)
            print("✅ Se hizo clic en el botón de búsqueda")
            time.sleep(2)

            # Buscar botones "Seleccionar" en los resultados
            print(f"Buscando el botón 'Seleccionar' para el perfil: '{perfil_excel}'...")
            try:
                # Primero intentar encontrar la fila que contiene el texto del perfil
                fila_perfil = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, 
                        f"//table[@id='int-ocu-table']/tbody/tr[td[contains(text(), '{perfil_excel}')]]"))
                )
                boton = fila_perfil.find_element(By.CSS_SELECTOR, "button.btn-primary")
                driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                time.sleep(0.5)  # Dar tiempo para que el scroll termine
                driver.execute_script("arguments[0].click();", boton)
                time.sleep(1)  # Espera a que se cargue el siguiente paso
                
                print(f"✅ Se hizo clic en el botón 'Seleccionar' para: '{perfil_excel}'")

                # --- Esperar a que aparezca el formulario de agregar interés y hacer clic en "Agregar" ---
                print("Esperando a que aparezca el formulario 'Agregar interés ocupacional'...")
                formulario_agregar = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, 'int-fadd-group'))
                )
                print("✅ Formulario 'Agregar interés ocupacional' visible")

                print("Buscando y haciendo clic en el botón 'Agregar'...")
                boton_agregar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'int-fadd-add'))
                )
                boton_agregar.click()
                print("✅ Se hizo clic en el botón 'Agregar'")
                time.sleep(3) # Espera a que se complete la adición

            except Exception as e:
                print(f"❌ No se pudo encontrar o hacer clic en el botón 'Seleccionar' para '{perfil_excel}': {e}")
                # Verificar que el interés se haya agregado correctamente
                try:
                    texto_encontrado = driver.find_elements(By.XPATH, f"//*[contains(text(), '{perfil_excel}')]")
                    if texto_encontrado:
                        print(f"✅ Verificado: El interés '{perfil_excel}' se agregó correctamente")
                        return True
                    else:
                        print(f"⚠️ No se encontró confirmación visual del interés agregado")
                        return False
                except:
                    print("No se pudo verificar la adición del interés")
                    return False
        except Exception as e:
            print(f"❌ Error general en intereses_ocupacionales: {str(e)}")
            traceback.print_exc()
            return False
    except Exception as e:
        print(f"❌ Error general en experiencia_laboral: {e}")
        # Capturar una captura de pantalla para depuración
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"error_screenshot_{timestamp}.png"
        try:
            driver.save_screenshot(screenshot_path)
            print(f"Se ha guardado una captura de pantalla en: {screenshot_path}")
        except Exception as screenshot_error:
            print(f"No se pudo guardar la captura de pantalla: {screenshot_error}")