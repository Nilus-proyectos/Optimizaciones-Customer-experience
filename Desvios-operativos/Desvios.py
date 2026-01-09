import os
import sys
import re
import time
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

# ================== Entorno ==================
load_dotenv()

def get_env(name, default=None, required=False):
    v = os.getenv(name, default)
    if required and (v is None or str(v).strip() == ""):
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
GSHEET_WORKSHEET_NAME = get_env("GSHEET_WORKSHEET_NAME", default="check_nueva_info_desvios")
GSERVICE_CREDENTIALS_JSON = os.getenv("GSERVICE_CREDENTIALS_JSON", "")
GSERVICE_CREDENTIALS_JSON_CONTENT = os.getenv("GSERVICE_CREDENTIALS_JSON_CONTENT", "")

# Selenium
SELENIUM_HEADLESS = get_env("SELENIUM_HEADLESS", default="true").strip().lower() == "true"
SELENIUM_WINDOW_SIZE = get_env("SELENIUM_WINDOW_SIZE", default="1920,1080")

# Fechas
TIMEZONE = get_env("TIMEZONE", default="America/Argentina/Buenos_Aires")
DAYS_OFFSET = int(get_env("DAYS_OFFSET", default="1"))  # por defecto "ayer"

# PyInstaller vs script
if getattr(sys, 'frozen', False):
    BASE_PATH = sys._MEIPASS
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# ================== Utilidades ==================
def enviar_notificacion_slack(mensaje: str):
    try:
        resp = requests.post(SLACK_WEBHOOK_URL, json={"text": mensaje}, timeout=15)
        if resp.status_code != 200:
            print(f"❌ Error Slack: {resp.status_code} - {resp.text}")
        else:
            print("Slack OK")
    except Exception as e:
        print(f"❌ Excepción enviando mensaje a Slack: {repr(e)}")

def get_gservice_credentials_path() -> str:
    """
    1) Si GSERVICE_CREDENTIALS_JSON existe (ruta absoluta o relativa al script) y está en disco, úsalo.
    2) Si no existe y tenemos GSERVICE_CREDENTIALS_JSON_CONTENT, lo materializa en BASE_PATH/credenciales/service-account.json
    """
    if GSERVICE_CREDENTIALS_JSON:
        path = GSERVICE_CREDENTIALS_JSON if os.path.isabs(GSERVICE_CREDENTIALS_JSON) else os.path.join(BASE_PATH, GSERVICE_CREDENTIALS_JSON)
        if os.path.exists(path):
            return path

    if GSERVICE_CREDENTIALS_JSON_CONTENT:
        dest_dir = os.path.join(BASE_PATH, "credenciales")
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, "service-account.json")
        with open(dest, "w") as f:
            f.write(GSERVICE_CREDENTIALS_JSON_CONTENT)
        print(f"Materialicé credenciales en: {dest}")
        return dest

    raise FileNotFoundError("No se encontró archivo de credenciales ni contenido en env (GSERVICE_CREDENTIALS_JSON_CONTENT).")

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
        button = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, selector)))
        button.click()
        return True
    except Exception as e:
        print(f"❌ Error al hacer clic en el selector '{selector}'")
        return False

def procesar_pedido(driver, datos_pedido, producto_buscado, cantidad_deseada):
    from difflib import SequenceMatcher

    def similitud(n1, n2):
        return SequenceMatcher(None, (n1 or "").lower(), (n2 or "").lower()).ratio()

    pedido_url = f"{BACKOFFICE_URL.replace('/login','')}/orders/{datos_pedido}"
    print(f"\n🔄 Procesando pedido: {datos_pedido}")
    driver.get(pedido_url)
    time.sleep(5)

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
            mejor_match = None
            for producto in productos:
                nombre = producto.find_element(By.TAG_NAME, "span").text.strip()
                if coincidencia_parcial(producto_buscado, nombre):
                    mejor_match = (producto, nombre)
                    print(f"🔍 Coincidencia parcial → Producto: {nombre}")
                    break
        else:
            print(f"🔍 Coincidencia por similitud ({mayor_similitud:.2f}) → Producto: {mejor_match[1]}")

        if not mejor_match:
            msg = f"❌ Producto '{producto_buscado}' no encontrado en pedido {datos_pedido}"
            print(msg)
            enviar_notificacion_slack(msg)
            return

        producto, nombre = mejor_match
        boton_svg = producto.find_element(By.CSS_SELECTOR, 'svg[data-testid="DoNotDisturbOnIcon"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_svg)
        try:
            boton_svg.click()
        except Exception:
            driver.execute_script("arguments[0].click();", boton_svg)

        motivo_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//label[contains(text(), "Selecciona un motivo")]'))
        )
        combo_div = motivo_label.find_element(By.XPATH, './ancestor::div[contains(@class, "MuiFormControl-root")]')
        combo_clickable = combo_div.find_element(By.XPATH, './/div[@role="button" or @role="combobox"]')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", combo_clickable)
        combo_clickable.click()

        opcion = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//li[contains(text(), "Support - DTC - Operations - Missing Product")]'))
        )
        opcion.click()

        input_cantidad = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "quantity")))
        input_cantidad.clear()
        input_cantidad.send_keys(str(cantidad_deseada))

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Solicitar']"))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar cambios']"))).click()

        print(f"✅ Pedido {datos_pedido} procesado con éxito.")
    except Exception as e:
        msg = f"❌ Error procesando el pedido {datos_pedido}"
        print(msg)
        enviar_notificacion_slack(msg)

# ================== Inicio ==================
enviar_notificacion_slack("🚀 El script de procesamiento de desvíos ha comenzado EN AMBOS PAÍSES.")

# Selenium
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
time.sleep(10)

# Google Sheets
ruta_json = get_gservice_credentials_path()
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciales = ServiceAccountCredentials.from_json_keyfile_name(ruta_json, scope)
cliente = gspread.authorize(credenciales)

sheet = cliente.open_by_key(SHEET_ID)
worksheet = sheet.worksheet(GSHEET_WORKSHEET_NAME)
valores = worksheet.get_all_values()
filas = valores[1:]  # Saltar encabezado

# Filtrar filas con país AR/MX (columna 1 -> índice 1)
filas_mx = [fila for fila in filas if len(fila) > 1 and str(fila[1]).strip().lower() in ["mx", "ar"]]

df = pd.DataFrame([
    [fila[0], fila[2], fila[7], fila[9], fila[11], (fila[12] if len(fila) > 12 else "0")]
    for fila in filas_mx if len(fila) >= 12
], columns=["fecha", "type_desvio", "datos_pedido", "producto_afectado", "cantidad_original", "cantidad_modificada"])

# Limpieza y fechas
df["fecha"] = pd.to_datetime(df["fecha"].astype(str).str.strip(), dayfirst=True, errors="coerce")
df.dropna(subset=["fecha"], inplace=True)

def extraer_id(texto):
    match = re.search(r"\b[a-f0-9]{32}\b", str(texto))
    return match.group(0) if match else None

df["datos_pedido"] = df["datos_pedido"].apply(extraer_id)

# Limpieza de datos nulos primero
df.dropna(subset=["fecha"], inplace=True)

# Convierte la columna fecha a datetime real
df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")

# Obtener la fecha de ayer como objeto datetime sin hora
ayer = (datetime.now() - timedelta(days=1)).date()
print("Ayer:", ayer)
print("Fechas únicas en df:", df["fecha"].dt.date.unique())

# Filtrar por fecha usando .dt.date
df_ayer = df[df["fecha"].dt.date == ayer]

print(f"📆 Fecha actual: {datetime.now().strftime('%d/%m/%Y')}")
print(f"📆 Fecha filtrada (ayer): {ayer}")
print(f"🔎 Filas encontradas: {len(df_ayer)}")

# Filtrar solo pedidos con estado 'faltante' o 'faltante_parcial'
df_ayer = df_ayer[df_ayer["type_desvio"].isin(["faltante", "faltante_parcial"])]

if df_ayer.empty:
    print("⚠️ No se encontraron pedidos del día anterior.")
    driver.quit()
    exit()

# Procesar pedidos
for i, row in df_ayer.iterrows():
    try:
        datos_pedido = str(row["datos_pedido"]).strip()
        producto_afectado = str(row["producto_afectado"]).strip()

        try:
            original = int(row.get("cantidad_original", 0) or 0)
        except ValueError:
            original = 0

        try:
            modificada = int(row["cantidad_modificada"]) if str(row["cantidad_modificada"]).strip().isdigit() else 0
        except Exception:
            modificada = 0

        tipo_desvio = str(row.get("type_desvio", "")).strip().lower()
        cantidad_deseada = max(original - modificada, 0) if tipo_desvio == "faltante_parcial" else original

        procesar_pedido(driver, datos_pedido, producto_afectado, cantidad_deseada)
    except Exception as e:
        print(f"⚠️ Error en fila {i}: {e}")

time.sleep(5)
driver.quit()
print("✅ Proceso finalizado y navegador cerrado.")
