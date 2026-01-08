import os
import sys
import time
import re
import requests
import pandas as pd
from datetime import datetime, timedelta

from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from oauth2client.service_account import ServiceAccountCredentials
import gspread

# =============== Entorno ===============
load_dotenv()

def get_env(name, default=None, required=False):
    v = os.getenv(name, default)
    if required and (v is None or v.strip() == ""):
        print(f"❌ Falta la variable de entorno requerida: {name}")
        sys.exit(1)
    return v

# Slack (Incoming Webhook)
SLACK_WEBHOOK_URL = get_env("SLACK_WEBHOOK_URL", required=True)

# Backoffice
BACKOFFICE_URL = get_env("BACKOFFICE_URL", default="https://backoffice.nilus.co/es-AR/login")
BACKOFFICE_EMAIL = get_env("BACKOFFICE_EMAIL", required=True)
BACKOFFICE_PASSWORD = get_env("BACKOFFICE_PASSWORD", required=True)

# Google Sheets
SHEET_ID = get_env("SHEET_ID", required=True)
GSHEET_WORKSHEET_NAME = get_env("GSHEET_WORKSHEET_NAME", default="ReclamoAI")
GSERVICE_CREDENTIALS_JSON = get_env("GSERVICE_CREDENTIALS_JSON", required=True)

# Selenium
SELENIUM_HEADLESS = get_env("SELENIUM_HEADLESS", default="true").strip().lower() == "true"
SELENIUM_WINDOW_SIZE = get_env("SELENIUM_WINDOW_SIZE", default="1920,1080")

# Base path para rutas relativas
BASE_PATH = os.path.dirname(os.path.abspath(__file__))

def resolver_ruta(p: str) -> str:
    return p if os.path.isabs(p) else os.path.join(BASE_PATH, p)

# =============== Slack ===============
def enviar_notificacion_slack(mensaje: str):
    payload = {"text": mensaje}
    try:
        resp = requests.post(SLACK_WEBHOOK_URL, json=payload, timeout=15)
        if resp.status_code != 200:
            print(f"❌ Error Slack: {resp.status_code} - {resp.text}")
        else:
            print("Slack OK")
    except Exception as e:
        print(f"❌ Excepción Slack: {repr(e)}")

enviar_notificacion_slack("🚀 El script de RECLAMOS ha comenzado.")

# =============== Utilidades ===============
def normalizar(texto):
    """Elimina símbolos, convierte a minúsculas y separa palabras clave."""
    texto = re.sub(r"[^a-zA-Z0-9áéíóúñü\s]", "", (texto or "").lower())
    return set(texto.split())

def coincidencia_parcial(nombre_excel, nombre_web, umbral=0.5):
    palabras_excel = normalizar(nombre_excel)
    palabras_web = normalizar(nombre_web)
    if not palabras_excel:
        return False
    comunes = palabras_excel & palabras_web
    return len(comunes) / len(palabras_excel) >= umbral

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

# =============== Google Sheets ===============
ruta_json = resolver_ruta(GSERVICE_CREDENTIALS_JSON)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(ruta_json, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(GSHEET_WORKSHEET_NAME)
valores = worksheet.get_all_values()

filas = valores[1:]  # Saltar encabezado
datos_limpios = []
for fila in filas:
    # Ajusta índices según tus columnas reales
    if len(fila) >= 15 and (fila[0] or "").strip():
        datos_limpios.append([
            (fila[0] or "").strip(),     # fecha
            (fila[8] or "").strip(),     # Estado
            (fila[15] or "").strip(),    # url
            (fila[9] or "").strip(),     # Producto_Reclamado
            (fila[10] if len(fila) > 15 else "0").strip()  # Cantidad
        ])

df = pd.DataFrame(datos_limpios, columns=["fecha", "Estado", "url", "Producto_Reclamado", "Cantidad"])
df["Estado"] = df["Estado"].fillna("").astype(str).str.strip()

# Fechas
df.dropna(subset=["fecha"], inplace=True)
df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")
hoy_date = (datetime.now() - timedelta(days=0)).date()
df_hoy = df[df["fecha"].dt.date == hoy_date]

print(f"📆 Fecha actual: {datetime.now().strftime('%d/%m/%Y')}")
print(f"🔎 Filas encontradas hoy: {len(df_hoy)}")

# Filtrar estados relevantes
df_hoy = df_hoy[
    df_hoy["Estado"].str.lower().str.contains("faltante|mal estado", na=False) |
    df_hoy["Estado"].isna() |
    (df_hoy["Estado"].str.strip() == "")
]

# =============== Selenium ===============
opts = Options()
opts.add_argument(f"--window-size={SELENIUM_WINDOW_SIZE}")
if SELENIUM_HEADLESS:
    opts.add_argument("--headless=new")
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)

# Login
driver.get(BACKOFFICE_URL)
time.sleep(5)
driver.find_element(By.ID, "email").send_keys(BACKOFFICE_EMAIL)
driver.find_element(By.ID, "password").send_keys(BACKOFFICE_PASSWORD)
click_button(driver, "//button[text()='INGRESAR']", By.XPATH)
time.sleep(5)

def procesar_producto_en_pedido(driver, producto_buscado, cantidad_deseada, estado):
    from difflib import SequenceMatcher
    def similitud(nombre1, nombre2):
        return SequenceMatcher(None, (nombre1 or "").lower(), (nombre2 or "").lower()).ratio()

    try:
        productos = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "css-1es96wk"))
        )

        mejor_match, mayor_similitud = None, 0.0
        for producto in productos:
            nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
            score = similitud(producto_buscado, nombre)
            if score > mayor_similitud:
                mayor_similitud = score
                mejor_match = (producto, nombre)

        if not (mejor_match and mayor_similitud >= 0.8):
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre):
                    mejor_match = (producto, nombre)
                    break

        if not mejor_match:
            print(f"❌ Producto '{producto_buscado}' no encontrado por ningún método.")
            enviar_notificacion_slack(f"❌ Producto '{producto_buscado}' no encontrado en pedido.")
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_svg)
        try:
            boton_svg.click()
        except Exception:
            driver.execute_script("arguments[0].click();", boton_svg)

        estado_norm = (str(estado) if estado is not None else "").strip().lower()
        if re.search(r"mal[\s/-]?estado", estado_norm):
            motivo_texto = "Support - DTC - Delivery Point - Product in bad condition"
        elif re.search(r"faltante[s]?", estado_norm):
            motivo_texto = "Support - DTC - Delivery Point - Missing Product"
        else:
            motivo_texto = "Support - DTC - Delivery Point - Missing Product"

        motivo_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Selecciona un motivo")]'))
        )
        combo_div = motivo_label.find_element(By.XPATH, './ancestor::div[contains(@class, "MuiFormControl-root")]')
        combo_clickable = combo_div.find_element(By.XPATH, './/div[@role="button" or @role="combobox"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combo_clickable)
        combo_clickable.click()

        motivo_lower = motivo_texto.lower()
        opcion_xpath = f"//li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), \"{motivo_lower}\")]"
        opcion = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, opcion_xpath)))
        opcion.click()

        input_cantidad = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "quantity")))
        input_cantidad.clear()
        input_cantidad.send_keys(str(cantidad_deseada))

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Solicitar']"))).click()
        print(f"✅ Producto '{nombre}' procesado en pedido.")
    except Exception as e:
        mensaje_error = f"❌ Error procesando producto '{producto_buscado}': {str(e)}"
        print(mensaje_error)
        enviar_notificacion_slack(mensaje_error)

# Procesar por URL
grupos = df_hoy.groupby("url")
for url, grupo in grupos:
    try:
        fila_real = grupo.index[0] + 65
        valor_celda = worksheet.cell(fila_real, 13).value
        if valor_celda and valor_celda.strip():
            print(f"⏭️ Pedido {url} ya procesado. Saltando...")
            continue

        pedido_url = f"{BACKOFFICE_URL.replace('/login','')}/orders/{url}"
        print(f"\n🔄 Procesando pedido: {url}")
        driver.get(pedido_url)
        time.sleep(5)

        for _, row in grupo.iterrows():
            producto_reclamado = str(row["Producto_Reclamado"]).strip()
            estado = row["Estado"]
            try:
                modificada = int(row["Cantidad"]) if str(row["Cantidad"]).strip().isdigit() else 0
            except:
                modificada = 0
            cantidad_deseada = modificada
            procesar_producto_en_pedido(driver, producto_reclamado, cantidad_deseada, estado)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))).click()

        for idx in grupo.index:
            worksheet.update_cell(idx + 65, 13, "✅")
    except Exception as e:
        print(f"⚠️ Error en pedido {url}: {e}")

time.sleep(5)
driver.quit()
print("✅ Proceso finalizado y navegador cerrado.")
enviar_notificacion_slack("✅ Proceso de RECLAMOS finalizado correctamente.")
