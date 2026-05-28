"""Configuración del paquete `robot`: variables de entorno, URLs y mapeos.

Centraliza todo lo que era "variables de módulo" en `downloader.py`:

- Parsing de variables de entorno (timeouts, flags, paths de Playwright).
- URLs del portal del SRI y selectores comunes.
- Mapeos canónicos de tipos de comprobante y estados.

Importar este módulo tiene **side effects** (setea `PLAYWRIGHT_BROWSERS_PATH`
y `PYPPETEER_HOME` si no estaban definidos). Eso es intencional: replica
el comportamiento que tenía `downloader.py` cuando estas constantes vivían
ahí. Ningún otro lugar del repo debería hacer ese setup.

El paquete `downloader.py` re-exporta lo que externamente se consume
(`MANUAL_CONSULTA_RECIBIDOS`, `TIPOS_MAP`, `ESTADOS_EMITIDOS_MAP`) para
mantener compatible la API previa.
"""
from __future__ import annotations

import os
import sys


# --------------------------------------------------------------------------- #
# Side effect: setear paths de Playwright si no estaban definidos.
# Esto debe correr ANTES de importar playwright en cualquier módulo.
# --------------------------------------------------------------------------- #
if os.name == "nt":
    local_app = os.getenv("LOCALAPPDATA")
    if local_app:
        pw_path = os.path.join(local_app, "ms-playwright")
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", pw_path)
        os.environ.setdefault("PYPPETEER_HOME", pw_path)
else:
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "/root/.cache/ms-playwright")
    os.environ.setdefault("PYPPETEER_HOME", "/root/.cache/ms-playwright")


# --------------------------------------------------------------------------- #
# Flags de Playwright (headless, slowmo, devtools, persistent profile)
# --------------------------------------------------------------------------- #
# Default de headless según plataforma:
# - Windows / macOS: SIEMPRE hay entorno gráfico → default visible ("0").
#   La heurística `not DISPLAY` NO sirve acá: DISPLAY es una variable de X11
#   que nunca existe en Windows, así que daría headless por error.
# - Linux: headless si es un servidor (RENDER) o no hay servidor gráfico
#   (DISPLAY ausente). En un escritorio Linux con X corre visible.
# Siempre se puede forzar con la env var PLAYWRIGHT_HEADLESS.
if sys.platform in ("win32", "darwin"):
    DEFAULT_HEADLESS = "0"
else:
    DEFAULT_HEADLESS = "1" if (os.getenv("RENDER") or not os.getenv("DISPLAY")) else "0"
HEADLESS_ENV = os.getenv("PLAYWRIGHT_HEADLESS", DEFAULT_HEADLESS).strip().lower()
HEADLESS = HEADLESS_ENV not in {"0", "false", "no", "off"}

try:
    SLOW_MO = int(os.getenv("PLAYWRIGHT_SLOWMO", "0"))
except ValueError:
    SLOW_MO = 0

DEVTOOLS_ENV = os.getenv("PLAYWRIGHT_DEVTOOLS", "0").strip().lower()
DEVTOOLS = DEVTOOLS_ENV in {"1", "true", "yes", "on"}

PERSISTENT_PROFILE_ENV = os.getenv("PLAYWRIGHT_PERSISTENT_PROFILE", "1").strip().lower()
USE_PERSISTENT_PROFILE = PERSISTENT_PROFILE_ENV in {"1", "true", "yes", "on"}
USER_DATA_DIR = os.getenv("PLAYWRIGHT_USER_DATA_DIR", "browser_profile").strip()

# Preferir Chrome del sistema en lugar de Chromium bundled. Confirmado que
# reCAPTCHA Enterprise del SRI valida la versión moderna de Chrome: la app
# de referencia que el usuario validó usa Chrome 148 del sistema y pasa el
# captcha sin problemas. Chromium bundled de Playwright suele estar varias
# versiones atrás (~127 en Playwright 1.47) y eso baja el score.
#
# Si esta var está activa, `_abrir_navegador` intenta abrir Chrome del sistema
# (`channel="chrome"`) PRIMERO, en cualquier modo (persistente o no). Si Chrome
# no está instalado, cae al Chromium bundled de Playwright.
#
# Si por alguna razón Chrome del sistema rompe algo en Emitidos, podés volver
# al comportamiento previo con: PREFER_SYSTEM_CHROME=0
PREFER_SYSTEM_CHROME = (
    os.getenv("PREFER_SYSTEM_CHROME", "1").strip().lower()
    in {"1", "true", "yes", "on", "si"}
)


# --------------------------------------------------------------------------- #
# Timeouts y pausas del scraping
# --------------------------------------------------------------------------- #
try:
    DOWNLOAD_TIMEOUT = int(os.getenv("SRI_DOWNLOAD_TIMEOUT_MS", "120000"))
except (TypeError, ValueError):
    DOWNLOAD_TIMEOUT = 120000

PAUSE_AT_LOGIN_ENV = os.getenv("PAUSE_BEFORE_INGRESAR", "0").strip().lower()
PAUSE_AT_LOGIN = PAUSE_AT_LOGIN_ENV in {"1", "true", "yes", "on"}
PAUSE_PROMPT = os.getenv(
    "PAUSE_BEFORE_INGRESAR_PROMPT",
    "Pausa antes de hacer clic en 'Ingresar'. "
    "Realiza los cambios necesarios y presiona Enter para continuar.",
).strip()

PAUSE_BEFORE_CONSULTAR_ENV = os.getenv("PAUSE_BEFORE_CONSULTAR", "0").strip().lower()
PAUSE_BEFORE_CONSULTAR = PAUSE_BEFORE_CONSULTAR_ENV in {"1", "true", "yes", "on"}
try:
    PAUSE_BEFORE_CONSULTAR_SECONDS = int(os.getenv("PAUSE_BEFORE_CONSULTAR_SECONDS", "0"))
except ValueError:
    PAUSE_BEFORE_CONSULTAR_SECONDS = 0


# --------------------------------------------------------------------------- #
# Robot SRI — Recibidos
# --------------------------------------------------------------------------- #
try:
    RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS = int(
        os.getenv("RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS", "10000")
    )
except ValueError:
    RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS = 10000

try:
    RECIBIDOS_CONSULTA_INTENTOS = max(1, int(os.getenv("RECIBIDOS_CONSULTA_INTENTOS", "5")))
except ValueError:
    RECIBIDOS_CONSULTA_INTENTOS = 5

try:
    RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC = max(
        0.0,
        float(os.getenv("RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC", "1.2")),
    )
except ValueError:
    RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC = 1.2

try:
    RECIBIDOS_AUTO_PRE_EXECUTE_MS = max(
        0, int(os.getenv("RECIBIDOS_AUTO_PRE_EXECUTE_MS", "450"))
    )
except ValueError:
    RECIBIDOS_AUTO_PRE_EXECUTE_MS = 450

try:
    RECIBIDOS_AUTO_POST_EXECUTE_MS = max(
        0, int(os.getenv("RECIBIDOS_AUTO_POST_EXECUTE_MS", "300"))
    )
except ValueError:
    RECIBIDOS_AUTO_POST_EXECUTE_MS = 300

try:
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS = max(
        10000, int(os.getenv("RECIBIDOS_AUTO_RESULT_TIMEOUT_MS", "60000"))
    )
except ValueError:
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS = 60000

try:
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO = max(
        2, int(os.getenv("RECIBIDOS_REHIDRATAR_DESDE_INTENTO", "3"))
    )
except ValueError:
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO = 3

RECIBIDOS_REHIDRATAR_ON_CAPTCHA = (
    os.getenv("RECIBIDOS_REHIDRATAR_ON_CAPTCHA", "1").strip().lower()
    in {"1", "true", "yes", "on", "si"}
)

MANUAL_CONSULTA_RECIBIDOS_ENV = os.getenv("RECIBIDOS_MANUAL_CONSULTA", "0").strip().lower()
MANUAL_CONSULTA_RECIBIDOS = MANUAL_CONSULTA_RECIBIDOS_ENV in {"1", "true", "yes", "on"}

# Humanizar la interacción antes de "Consultar" en Recibidos. El reCAPTCHA
# Enterprise asigna score por trayectoria de mouse, pausas y tiempo en página;
# sin estos eventos el score baja a ~0.0 y el SRI rechaza con "Captcha
# incorrecta". Generamos mouse moves con `steps` intermedios, pausas
# aleatorias y hover sobre el botón antes del click.
RECIBIDOS_HUMANIZAR_PRE_CLICK = (
    os.getenv("RECIBIDOS_HUMANIZAR_PRE_CLICK", "1").strip().lower()
    in {"1", "true", "yes", "on", "si"}
)
try:
    # Default subido a 5000 ms (5s): el test manual del usuario confirmó que
    # con ~1.8s no alcanza para que reCAPTCHA Enterprise considere humano al
    # bot; con tiempo de página más largo (rango 4-7s) sí pasa.
    RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS = max(
        0, int(os.getenv("RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS", "5000"))
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS = 5000

# Clics reales en una zona "blanca" segura del viewport (lejos de inputs/
# selects/botones) antes de presionar Consultar. El reCAPTCHA Enterprise
# valora positivamente eventos `click` reales con timing humano. Test manual
# del usuario: 6 clics rápidos en una zona vacía sí hacen que el captcha
# pase, donde solo mover el mouse no era suficiente.
try:
    RECIBIDOS_HUMANIZAR_CLICKS = max(
        0, int(os.getenv("RECIBIDOS_HUMANIZAR_CLICKS", "6"))
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_CLICKS = 6

try:
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN = max(
        0, int(os.getenv("RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN", "80"))
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN = 80

try:
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MAX = max(
        RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN,
        int(os.getenv("RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MAX", "180")),
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MAX = max(
        RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN, 180
    )

# Coordenadas relativas al viewport (0.0-1.0) donde caen los clicks blancos.
# Default 0.82 / 0.48 → lado derecho-medio del viewport, donde el SRI suele
# tener espacio en blanco (el formulario está alineado a la izquierda).
try:
    RECIBIDOS_HUMANIZAR_CLICK_X_RATIO = max(
        0.05, min(0.95, float(os.getenv("RECIBIDOS_HUMANIZAR_CLICK_X_RATIO", "0.82")))
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_CLICK_X_RATIO = 0.82

try:
    RECIBIDOS_HUMANIZAR_CLICK_Y_RATIO = max(
        0.05, min(0.95, float(os.getenv("RECIBIDOS_HUMANIZAR_CLICK_Y_RATIO", "0.48")))
    )
except ValueError:
    RECIBIDOS_HUMANIZAR_CLICK_Y_RATIO = 0.48


# --------------------------------------------------------------------------- #
# Robot SRI — Emitidos
# --------------------------------------------------------------------------- #
try:
    EMITIDOS_RESET_AFTER_DAY_DOCS = max(
        1, int(os.getenv("EMITIDOS_RESET_AFTER_DAY_DOCS", "51"))
    )
except ValueError:
    EMITIDOS_RESET_AFTER_DAY_DOCS = 51

try:
    EMITIDOS_RESET_PAUSE_MS = max(0, int(os.getenv("EMITIDOS_RESET_PAUSE_MS", "1800")))
except ValueError:
    EMITIDOS_RESET_PAUSE_MS = 1800


# --------------------------------------------------------------------------- #
# URLs del portal SRI
# --------------------------------------------------------------------------- #
URLS = {
    "Recibidos": "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recibidos/comprobantesRecibidos.jsf",
    "Emitidos":  "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/emitidos/comprobantesEmitidos.jsf",
}
RECIBIDOS_DIRECT_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recibidos/comprobantesRecibidos.jsf"
RECUPERAR_COMPROBANTES_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recuperarComprobantes.jsf"
MENU_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf"
MENU_URL_ALT = (
    "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf"
    "?&contextoMPT=https://srienlinea.sri.gob.ec/tuportal-internet"
    "&pathMPT=Facturacion%20Electronica%20%2F%20Produccion"
    "&actualMPT=Consultas%20"
    "&linkMPT=%2Fcomprobantes-electronicos-internet%2Fpages%2Fconsultas%2Fmenu.jsf%3F"
    "&esFavorito=S#"
)
MENU_EMITIDOS_TRIGGER_URL = (
    "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf"
    "?&contextoMPT=https://srienlinea.sri.gob.ec/tuportal-internet"
)
PORTAL_HOME = (
    "https://srienlinea.sri.gob.ec/auth/realms/Internet/protocol/openid-connect/auth"
    "?client_id=app-sri-claves-angular"
    "&redirect_uri=https%3A%2F%2Fsrienlinea.sri.gob.ec%2Fsri-en-linea%2F%2Fcontribuyente%2Fperfil"
    "&state=34e5716b-3474-46e7-8c52-ddfe62a2404c"
    "&nonce=46d2f0a2-ce75-4cec-856d-987329a6f17e"
    "&response_mode=fragment&response_type=code&scope=openid"
)
AUTORIZACION_COMPROBANTES_SOAP_URL = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"


# --------------------------------------------------------------------------- #
# Selectores CSS / XPath comunes
# --------------------------------------------------------------------------- #
MENU_TOGGLE_SELECTOR = "#sri-menu, button#sri-menu, button.menu-button, button[aria-label*='menu']"
FACTURACION_MENU_SELECTOR = (
    "xpath=//a[.//span[contains(@class,'ui-menuitem-text') and "
    "normalize-space()='FACTURACIÓN ELECTRÓNICA']]"
)
MODULO_PRODUCCION_SELECTOR = (
    "xpath=//span[contains(@class,'ui-menuitem-text') and "
    "normalize-space()='Producción']/ancestor::a[1]"
)
CONSULTAS_SELECTOR = (
    "xpath=//span[contains(@class,'ui-menuitem-text') and "
    "normalize-space()='Consultas']/ancestor::a[1]"
)
OVERLAY_SELECTORS = ["#disablingDiv", "#disablingOverlay"]


# --------------------------------------------------------------------------- #
# Mensajes al usuario
# --------------------------------------------------------------------------- #
PORTAL_INDISPONIBLE_MENSAJE = (
    "El portal del SRI reporta indisponibilidad temporal. "
    "Intenta nuevamente en unos minutos."
)


# --------------------------------------------------------------------------- #
# Mapeos canónicos de tipos y estados
# --------------------------------------------------------------------------- #
TIPOS_MAP = {
    "Facturas": "Factura",
    "Retenciones": "Comprobante de Retencion",
    "Retencion": "Comprobante de Retencion",
    "Retención": "Comprobante de Retencion",
    "Notas de credito": "Notas de Credito",
    "Notas de debito": "Notas de Debito",
    "Liquidacion de compra": "Liquidacion de compra de bienes y prestacion de servicios",
    "Guia de Remision": "Guia de Remision",
    "Guias de Remision": "Guia de Remision",
    "Guia de remision": "Guia de Remision",
    "Guias de remision": "Guia de Remision",
}

ESTADOS_EMITIDOS_MAP = {
    "Autorizados": "Autorizados",
    "Autorizado": "Autorizados",
    "No Autorizados": "No Autorizados",
    "No autorizados": "No Autorizados",
    "Por Procesar": "Por Procesar",
    "Por procesar": "Por Procesar",
}

DOC_LABELS = {
    "01": "Factura",
    "03": "Liquidacion_de_Compra",
    "04": "NotaCredito",
    "05": "NotaDebito",
    "06": "GuiaRemision",
    "07": "Retencion",
}
