from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait




def chrome():
    # --- Inicializar el navegador Selenium con opciones ---
    chrome_options = Options()

    # Descomentar la siguiente línea para modo headless (sin interfaz gráfica), para visualizar el funcionamiento del aplicativo
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 10)  # Espera explícita de 10 segundos

    return driver, wait
