import os
import sys
import time
import ssl
import pandas as pd
import pytz
import re
import json
from datetime import datetime, timedelta

from dotenv import load_dotenv
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys

from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Carga de variables de entorno
load_dotenv()

def get_env(name, default=None, required=False):
    value = os.getenv(name, default)
    if required and (value is None or value == ""):
        print(f"❌ Falta la variable de entorno requerida: {name}")
        sys.exit(1)
    return value

# Slack
SLACK_TOKEN = get_env("SLACK_TOKEN", required=True)
# Acepta varios nombres por compatibilidad
SLACK_CHANNEL_ID_NOTIFICACIONES = (
    os.getenv("SLACK_CHANNEL_ID_NOTIFICACIONES")
    or os.getenv("CHANNEL_ID_NOTIFICACIONES")
    or os.getenv("CHANNEL_ID")
)
if not SLACK_CHANNEL_ID_NOTIFICACIONES:
    print("❌ Falta SLACK_CHANNEL_ID_NOTIFICACIONES (o CHANNEL_ID_NOTIFICACIONES / CHANNEL_ID)")
    sys.exit(1)

# Google Sheets
SHEET_ID = get_env("SHEET_ID", required=True)
GSHEET_WORKSHEET_NAME = get_env("GSHEET_WORKSHEET_NAME", default="Cancelar")
GSERVICE_CREDENTIALS_JSON = get_env("GSERVICE_CREDENTIALS_JSON", required=True)

# Backoffice
BACKOFFICE_URL = get_env("BACKOFFICE_URL", default="https://backoffice.nilus.co/es-AR/login")
BACKOFFICE_EMAIL = get_env("BACKOFFICE_EMAIL", required=True)
BACKOFFICE_PASSWORD = get_env("BACKOFFICE_PASSWORD", required=True)

# SSL y Selenium
SSL_CERT_PATH = get_env("SSL_CERT_PATH", default=None)
SELENIUM_HEADLESS = get_env("SELENIUM_HEADLESS", default="true").strip().lower() == "true"
SELENIUM_WINDOW_SIZE = get_env("SELENIUM_WINDOW_SIZE", default="1920,1080")

# Detectar si está ejecutándose como .exe empaquetado con PyInstaller
if getattr(sys, 'frozen', False):
    BASE_PATH = sys._MEIPASS  # Carpeta temporal de PyInstaller
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))  # Carpeta del script .py

def resolver_ruta(rel_or_abs_path: str) -> str:
    """Si la ruta no es absoluta, la resuelve relativa a BASE_PATH."""
    if rel_or_abs_path and not os.path.isabs(rel_or_abs_path):
        return os.path.join(BASE_PATH, rel_or_abs_path)
    return rel_or_abs_path

def obtener_ruta_certificado():
    # 1) Usa SSL_CERT_PATH si está definido
    if SSL_CERT_PATH:
        return resolver_ruta(SSL_CERT_PATH)
    # 2) Fallback a certifi si está instalado
    try:
        import certifi
        return certifi.where()
    except Exception:
        pass
    # 3) Fallback a bundle local (opcional)
    return os.path.join(BASE_PATH, "certificados", "cacert.pem")

# Certificado SSL para evitar problemas de conexión
ssl_context = ssl.create_default_context(cafile=obtener_ruta_certificado())

# === Conectar a Slack ===
client = WebClient(token=SLACK_TOKEN, ssl=ssl_context)

def enviar_notificacion_slack(mensaje: str):
    try:
        resp = client.chat_postMessage(channel=SLACK_CHANNEL_ID_NOTIFICACIONES, text=mensaje)
        if not resp.get("ok", False):
            print(f"❌ Slack error: {resp.get('error')}")
        else:
            print(f"Slack OK ts={resp.get('ts')}")
    except SlackApiError as e:
        print(f"SlackApiError: {e.response.get('error')}")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {repr(e)}")

enviar_notificacion_slack("PREVENCIÓN DE ARG Y MX ha comenzado 🚀")

# === Conexión a Google Sheets ===
RUTA_CREDENCIALES = resolver_ruta(GSERVICE_CREDENTIALS_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(GSHEET_WORKSHEET_NAME)
valores = worksheet.get_all_values()

# Pasar a DataFrame (primera fila son encabezados)
df = pd.DataFrame(valores[1:], columns=valores[0])

# Filtrar pedidos de hoy (solo impresión informativa)
hoy = datetime.now().strftime("%d/%m/%Y")
print(f"📌 Pedidos de hoy ({hoy}):")

# === Selenium / Inicio del script ===
options = Options()
options.add_argument(f"--window-size={SELENIUM_WINDOW_SIZE}")
if SELENIUM_HEADLESS:
    options.add_argument("--headless=new")

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

def click_button(driver, selector, by=By.CSS_SELECTOR, wait_time=10):
    try:
        button = WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((by, selector))
        )
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}': {e}")
        return False

def guardar_cambios(driver, wait_time=10):
    try:
        modal = WebDriverWait(driver, wait_time).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Cambiar el estado del pedido')]/ancestor::div[@role='dialog']"))
        )
        guardar_btn = modal.find_element(By.XPATH, ".//button[normalize-space(text())='Guardar cambios']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", guardar_btn)
        driver.execute_script("arguments[0].click();", guardar_btn)
        print("✅ 'Guardar cambios' se clickeó correctamente.")
        return True
    except Exception as e:
        print(f"❌ No se pudo hacer clic en el botón correcto: {e}")
        return False

# LOGIN
driver.get(BACKOFFICE_URL)
time.sleep(5)
driver.find_element(By.ID, "email").send_keys(BACKOFFICE_EMAIL)
driver.find_element(By.ID, "password").send_keys(BACKOFFICE_PASSWORD)
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
time.sleep(10)

for i, row in df.iterrows():
    pedido_id = str(row.iloc[1]).strip()  # Toma literalmente la columna 1 (índice 1)

    if not pedido_id:
        mensaje = f"❌ No se encontró un ID válido en: {pedido_id}"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue

    print(f"🔄 Procesando pedido {pedido_id}")
    try:
        driver.get(f"{BACKOFFICE_URL.replace('/login','')}/orders/{pedido_id}")
    except Exception as e:
        mensaje = f"⚠️ Error al abrir el pedido {pedido_id}: {e}"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue

    try:
        # Intentar hacer clic en el combobox con id='email'
        try:
            cambiar_estado = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='email']"))
            )
            cambiar_estado.click()
        except TimeoutException:
            # Intentar con id='status'
            try:
                cambiar_estado = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='combobox' and @id='status']"))
                )
                cambiar_estado.click()
            except TimeoutException:
                mensaje = f"❌ No se pudo hacer clic en cambiar estado intento 2 {pedido_id}. Continuando con el siguiente pedido..."
                print(mensaje)
                enviar_notificacion_slack(mensaje)
                continue

        # Seleccionar "Cancelado"
        cancelado_opcion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Cancelado')]"))
        )
        cancelado_opcion.click()
        time.sleep(1)

        # Motivo de cancelación
        motivo_combo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "reason_of_canceled"))
        )
        motivo_combo.click()
        time.sleep(2)
        try:
            motivo_opcion = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(text(), 'Pago anticipado')]"))
            )
            motivo_opcion.click()
        except Exception:
            mensaje = f"❌ No se pudo seleccionar el motivo por defecto para el pedido {pedido_id}"
            print(mensaje)
            enviar_notificacion_slack(mensaje)

        # Guardar cambios
        guardar_cambios(driver)
        print(f"✅ Pedido {pedido_id} cancelado con motivo 'Pago anticipado'")
        enviar_notificacion_slack(f"✅ Pedido {pedido_id} cancelado con motivo 'Pago anticipado'")
    except Exception:
        mensaje = f"❌ Error procesando pedido {pedido_id}: TAL VEZ YA ESTA ANULADO"
        print(mensaje)
        enviar_notificacion_slack(mensaje)
        continue

print("✅ Script finalizado correctamente.")
enviar_notificacion_slack("✅ Script finalizado correctamente.")
driver.quit()
sys.exit()
