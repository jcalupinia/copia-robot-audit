from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from pathlib import Path
from typing import Optional
import threading
from urllib.parse import urlencode
import pandas as pd
import csv, re, json, os, time, unicodedata, html, calendar
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
try:
    import pdfplumber
except Exception:
    pdfplumber = None
import xml.etree.ElementTree as ET
from robot import parser as xml_parser

from robot.parser import construir_reporte
try:
    from robot.pdf_layout.main import extract_pdf_fields as _extract_pdf_layout_fields
except Exception:
    _extract_pdf_layout_fields = None
from robot.download_resume import (
    delete_checkpoint as _delete_download_checkpoint,
    load_checkpoint as _load_download_checkpoint,
    mark_checkpoint_running as _mark_download_checkpoint_running,
    update_checkpoint_progress as _update_download_checkpoint_progress,
)

from typing import Callable

from robot.captcha_solver import (
    CaptchaSolverError,
    MAX_ATTEMPTS as CAPTCHA_MAX_ATTEMPTS,
    is_enabled as captcha_solver_enabled,
    solve_image as solve_captcha_image,
)

USER_NOTIFICATION_CALLBACK: Optional[Callable[[str], None]] = None
CANCEL_EVENT = threading.Event()

def set_user_notifier(callback: Optional[Callable[[str], None]]):
    global USER_NOTIFICATION_CALLBACK
    USER_NOTIFICATION_CALLBACK = callback

def request_cancel():
    CANCEL_EVENT.set()

def clear_cancel():
    CANCEL_EVENT.clear()

def cancel_requested() -> bool:
    return CANCEL_EVENT.is_set()

def _check_cancel(paso: str = "") -> None:
    if CANCEL_EVENT.is_set():
        raise RuntimeError("Proceso cancelado por el usuario.")

def _extraer_clave_fila(celdas) -> str:
    def _buscar_clave(texto: str) -> str:
        texto = (texto or "").strip()
        if not texto:
            return ""
        match = re.search(r"\d{49}", texto)
        if match:
            return match.group(0)
        solo_digitos = re.sub(r"\D", "", texto)
        return solo_digitos if len(solo_digitos) == 49 else ""

    try:
        total = celdas.count()
    except Exception:
        total = 0

    for idx_celda in range(min(total, 6)):
        try:
            texto = celdas.nth(idx_celda).inner_text().strip()
        except Exception:
            continue
        clave = _buscar_clave(texto)
        if clave:
            return clave
    return ""

# ====== Configuracion global ======
if os.name == "nt":
    local_app = os.getenv("LOCALAPPDATA")
    if local_app:
        pw_path = os.path.join(local_app, "ms-playwright")
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", pw_path)
        os.environ.setdefault("PYPPETEER_HOME", pw_path)
else:
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "/root/.cache/ms-playwright")
    os.environ.setdefault("PYPPETEER_HOME", "/root/.cache/ms-playwright")

DEFAULT_HEADLESS = "1" if (os.getenv("RENDER") or not os.getenv("DISPLAY")) else "0"
HEADLESS_ENV = os.getenv("PLAYWRIGHT_HEADLESS", DEFAULT_HEADLESS).strip().lower()
HEADLESS = HEADLESS_ENV not in {"0", "false", "no", "off"}
try:
    SLOW_MO = int(os.getenv("PLAYWRIGHT_SLOWMO", "0"))
except ValueError:
    SLOW_MO = 0
try:
    DOWNLOAD_TIMEOUT = int(os.getenv("SRI_DOWNLOAD_TIMEOUT_MS", "120000"))
except (TypeError, ValueError):
    DOWNLOAD_TIMEOUT = 120000
PAUSE_AT_LOGIN_ENV = os.getenv("PAUSE_BEFORE_INGRESAR", "0").strip().lower()
PAUSE_AT_LOGIN = PAUSE_AT_LOGIN_ENV in {"1", "true", "yes", "on"}
PAUSE_PROMPT = os.getenv(
    "PAUSE_BEFORE_INGRESAR_PROMPT",
    "Pausa antes de hacer clic en 'Ingresar'. Realiza los cambios necesarios y presiona Enter para continuar.",
).strip()
PAUSE_BEFORE_CONSULTAR_ENV = os.getenv("PAUSE_BEFORE_CONSULTAR", "0").strip().lower()
PAUSE_BEFORE_CONSULTAR = PAUSE_BEFORE_CONSULTAR_ENV in {"1", "true", "yes", "on"}
try:
    PAUSE_BEFORE_CONSULTAR_SECONDS = int(os.getenv("PAUSE_BEFORE_CONSULTAR_SECONDS", "0"))
except ValueError:
    PAUSE_BEFORE_CONSULTAR_SECONDS = 0
try:
    RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS = int(os.getenv("RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS", "10000"))
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
        0,
        int(os.getenv("RECIBIDOS_AUTO_PRE_EXECUTE_MS", "450")),
    )
except ValueError:
    RECIBIDOS_AUTO_PRE_EXECUTE_MS = 450
try:
    RECIBIDOS_AUTO_POST_EXECUTE_MS = max(
        0,
        int(os.getenv("RECIBIDOS_AUTO_POST_EXECUTE_MS", "300")),
    )
except ValueError:
    RECIBIDOS_AUTO_POST_EXECUTE_MS = 300
try:
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS = max(
        10000,
        int(os.getenv("RECIBIDOS_AUTO_RESULT_TIMEOUT_MS", "60000")),
    )
except ValueError:
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS = 60000
try:
    EMITIDOS_RESET_AFTER_DAY_DOCS = max(
        1,
        int(os.getenv("EMITIDOS_RESET_AFTER_DAY_DOCS", "51")),
    )
except ValueError:
    EMITIDOS_RESET_AFTER_DAY_DOCS = 51
try:
    EMITIDOS_RESET_PAUSE_MS = max(
        0,
        int(os.getenv("EMITIDOS_RESET_PAUSE_MS", "1800")),
    )
except ValueError:
    EMITIDOS_RESET_PAUSE_MS = 1800
try:
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO = max(
        2,
        int(os.getenv("RECIBIDOS_REHIDRATAR_DESDE_INTENTO", "3")),
    )
except ValueError:
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO = 3
RECIBIDOS_REHIDRATAR_ON_CAPTCHA = (
    os.getenv("RECIBIDOS_REHIDRATAR_ON_CAPTCHA", "1").strip().lower()
    in {"1", "true", "yes", "on", "si"}
)
DEVTOOLS_ENV = os.getenv("PLAYWRIGHT_DEVTOOLS", "0").strip().lower()
DEVTOOLS = DEVTOOLS_ENV in {"1", "true", "yes", "on"}
PERSISTENT_PROFILE_ENV = os.getenv("PLAYWRIGHT_PERSISTENT_PROFILE", "1").strip().lower()
USE_PERSISTENT_PROFILE = PERSISTENT_PROFILE_ENV in {"1", "true", "yes", "on"}
USER_DATA_DIR = os.getenv("PLAYWRIGHT_USER_DATA_DIR", "browser_profile").strip()
MANUAL_CONSULTA_RECIBIDOS_ENV = os.getenv("RECIBIDOS_MANUAL_CONSULTA", "0").strip().lower()
MANUAL_CONSULTA_RECIBIDOS = MANUAL_CONSULTA_RECIBIDOS_ENV in {"1", "true", "yes", "on"}

URLS = {
    "Recibidos": "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recibidos/comprobantesRecibidos.jsf",
    "Emitidos":  "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/emitidos/comprobantesEmitidos.jsf",
}
RECIBIDOS_DIRECT_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recibidos/comprobantesRecibidos.jsf"
RECUPERAR_COMPROBANTES_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/recuperarComprobantes.jsf"
MENU_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf"
MENU_URL_ALT = (
    "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf?&contextoMPT=https://srienlinea.sri.gob.ec/tuportal-internet&pathMPT=Facturacion%20Electronica%20%2F%20Produccion&actualMPT=Consultas%20&linkMPT=%2Fcomprobantes-electronicos-internet%2Fpages%2Fconsultas%2Fmenu.jsf%3F&esFavorito=S#"
)
MENU_EMITIDOS_TRIGGER_URL = "https://srienlinea.sri.gob.ec/comprobantes-electronicos-internet/pages/consultas/menu.jsf?&contextoMPT=https://srienlinea.sri.gob.ec/tuportal-internet"
PORTAL_HOME = "https://srienlinea.sri.gob.ec/auth/realms/Internet/protocol/openid-connect/auth?client_id=app-sri-claves-angular&redirect_uri=https%3A%2F%2Fsrienlinea.sri.gob.ec%2Fsri-en-linea%2F%2Fcontribuyente%2Fperfil&state=34e5716b-3474-46e7-8c52-ddfe62a2404c&nonce=46d2f0a2-ce75-4cec-856d-987329a6f17e&response_mode=fragment&response_type=code&scope=openid"
MENU_TOGGLE_SELECTOR = "#sri-menu, button#sri-menu, button.menu-button, button[aria-label*='menu']"
FACTURACION_MENU_SELECTOR = "xpath=//a[.//span[contains(@class,'ui-menuitem-text') and normalize-space()='FACTURACIÓN ELECTRÓNICA']]"
MODULO_PRODUCCION_SELECTOR = "xpath=//span[contains(@class,'ui-menuitem-text') and normalize-space()='Producción']/ancestor::a[1]"
CONSULTAS_SELECTOR = "xpath=//span[contains(@class,'ui-menuitem-text') and normalize-space()='Consultas']/ancestor::a[1]"
OVERLAY_SELECTORS = ["#disablingDiv", "#disablingOverlay"]
AUTORIZACION_COMPROBANTES_SOAP_URL = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"
PORTAL_INDISPONIBLE_MENSAJE = (
    "El portal del SRI reporta indisponibilidad temporal. Intenta nuevamente en unos minutos."
)


TIPOS_MAP = {
    "Facturas": "Factura",
    "Retenciones": "Comprobante de Retencion",
    "Retencion": "Comprobante de Retencion",
    "Retención": "Comprobante de Retencion",
    "Notas de credito": "Notas de Credito",
    "Notas de credito": "Notas de Credito",
    "Notas de debito": "Notas de Debito",
    "Notas de debito": "Notas de Debito",
    "Liquidacion de compra": "Liquidacion de compra de bienes y prestacion de servicios",
    "Liquidacion de compra": "Liquidacion de compra de bienes y prestacion de servicios",
    "Guia de Remision": "Guia de Remision",
    "Guia de Remision": "Guia de Remision",
    "Guias de Remision": "Guia de Remision",
    "Guias de Remision": "Guia de Remision",
    "Guia de remision": "Guia de Remision",
    "Guia de remision": "Guia de Remision",
    "Guias de remision": "Guia de Remision",
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


PDF_REPORT_COLUMNS = [
    "tipoDocumento",
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "contribuyenteEspecial",
    "agenteRetencion",
    "obligadoContabilidad",
    "tipoContribuyenteRIMPE",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "fechaAutorizacion",
    "razonSocialComprador",
    "identificacionComprador",
    "direccionComprador",
    "placa",
    "guia",
    "comprobanteModificado",
    "fechaEmisionModificado",
    "razonModificacion",
    "valorModificacion",
    "descripcionesProductos",
    "subtotalTarifaEspecial",
    "subtotal15",
    "subtotal12",
    "subtotal8",
    "subtotal5",
    "subtotal0",
    "subtotalNoObjetoIVA",
    "subtotalExentoIVA",
    "subtotalSinImpuestos",
    "totalDescuento",
    "ivaTarifaEspecial",
    "iva15",
    "iva12",
    "iva8",
    "iva5",
    "ice",
    "irbpnr",
    "propina",
    "valorTotal",
    "valorTotalSinSubsidio",
    "formaPago",
    "formaPagoMonto",
    "ambiente",
    "emision",
    "claveAcceso",
    "informacionAdicional",
]

RETENCION_REPORT_COLUMNS = [
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "obligadoContabilidad",
    "numeroContribuyenteEspecial",
    "numeroAgenteRetencion",
    "fechaAutorizacion",
    "ambiente",
    "emision",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "razonSocialSujetoRetenido",
    "identificacionSujetoRetenido",
    "claveAcceso",
    "Comprobante_Sustento",
    "Numero_Sustento",
    "Fecha_Emision_Sustento",
    "Ejercicio_Fiscal",
    "Base_Imponible_Ret_IVA",
    "Impuesto_Ret_IVA",
    "Porcentaje_Ret_IVA",
    "Valor_Retenido_IVA",
    "Base_Imponible_Ret_IR",
    "Impuesto_Ret_IR",
    "Porcentaje_Ret_IR",
    "Valor_Retenido_IR",
    "informacionAdicional",
    "Base_Imponible_Ret_IR_1",
    "Impuesto_Ret_IR_1",
    "Porcentaje_Ret_IR_1",
    "Valor_Retenido_IR_1",
    "Base_Imponible_Ret_IVA_1",
    "Impuesto_Ret_IVA_1",
    "Porcentaje_Ret_IVA_1",
    "Valor_Retenido_IVA_1",
    "tipoDocumento",
]

EMITIDOS_RETENCION_REPORT_COLUMNS = [
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Tarifas IVA",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Agente de Retención",
    "Contribuyente Especial",
    "Obligado a llevar Contabilidad",
    "Razón Social Sujeto Retenido",
    "Identificación Sujeto Retenido",
    "Periodo Fiscal",
    "Dirección del Establecimiento",
    "Tipo Identificación Sujeto Retenido",
    "Parte Relacionada",
    "Código de Sustento",
    "Código del Documento de Sustento",
    "Número de Documento de Sustento",
    "Fecha de Emisión del Documento de Sustento",
    "Fecha de Registro Contable",
    "Número de Autorización del Documento de Sustento",
    "Pago Local o Externo",
    "Código Impuesto Doc. Sustento",
    "Código Porcentaje",
    "Base Imponible Impuesto",
    "Tarifa",
    "Valor Impuesto",
    "Campos Adicionales",
    "RENTA - codigoRetencion",
    "RENTA - baseImponible",
    "RENTA - porcentajeRetener",
    "RENTA - valorRetenido",
    "Sustento Imp. 1 - Código",
    "Sustento Imp. 1 - Cod. Porcentaje",
    "Sustento Imp. 1 - Base Imponible",
    "Sustento Imp. 1 - Tarifa",
    "Sustento Imp. 1 - Valor",
    "Factura Ret. 1 - Codigo",
    "Factura Ret. 1 - Cod. Porcentaje",
    "Factura Ret. 1 - Tarifa",
    "Factura Ret. 1 - Valor",
    "IVA - codigoRetencion",
    "IVA - baseImponible",
    "IVA - porcentajeRetener",
    "IVA - valorRetenido",
    "Factura Ret. 2 - Codigo",
    "Factura Ret. 2 - Cod. Porcentaje",
    "Factura Ret. 2 - Tarifa",
    "Factura Ret. 2 - Valor",
]

EMITIDOS_RETENCION_FORMA_PAGO_LABEL = {
    "01": "01 - SIN UTILIZACIÓN DEL SISTEMA FINANCIERO",
    "15": "15 - COMPENSACIÓN DE DEUDAS",
    "16": "16 - TARJETA DE DÉBITO",
    "17": "17 - DINERO ELECTRÓNICO",
    "18": "18 - TARJETA PREPAGO",
    "19": "19 - TARJETA DE CRÉDITO",
    "20": "20 - OTROS CON UTILIZACIÓN DEL SISTEMA FINANCIERO",
    "21": "21 - ENDOSO DE TÍTULOS",
}

EMITIDOS_RETENCION_DOC_CODE_LABEL = {
    "01": "01 - FACTURA",
    "03": "03 - LIQUIDACIÓN DE COMPRA",
    "04": "04 - NOTA DE CRÉDITO",
    "05": "05 - NOTA DE DÉBITO",
    "06": "06 - GUÍA DE REMISIÓN",
    "07": "07 - COMPROBANTE DE RETENCIÓN",
}

EMITIDOS_RETENCION_AMBIENTE_LABEL = {
    "1": "1 - Pruebas",
    "2": "2 - Producción",
}

EMITIDOS_RETENCION_TIPO_EMISION_LABEL = {
    "1": "1 - Emisión normal",
    "2": "2 - Emisión por indisponibilidad del sistema",
}

EMITIDOS_RETENCION_TEXT_FORCE_COLUMNS = {
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Agente de Retención",
    "Contribuyente Especial",
    "Obligado a llevar Contabilidad",
    "Razón Social Sujeto Retenido",
    "Identificación Sujeto Retenido",
    "Periodo Fiscal",
    "Dirección del Establecimiento",
    "Tipo Identificación Sujeto Retenido",
    "Parte Relacionada",
    "Código de Sustento",
    "Código del Documento de Sustento",
    "Número de Documento de Sustento",
    "Fecha de Emisión del Documento de Sustento",
    "Fecha de Registro Contable",
    "Número de Autorización del Documento de Sustento",
    "Pago Local o Externo",
    "Campos Adicionales",
}

EMITIDOS_RETENCION_NUMERIC_COLUMNS = {
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Base Imponible Impuesto",
    "Tarifa",
    "Valor Impuesto",
    "RENTA - baseImponible",
    "RENTA - porcentajeRetener",
    "RENTA - valorRetenido",
    "Sustento Imp. 1 - Base Imponible",
    "Sustento Imp. 1 - Tarifa",
    "Sustento Imp. 1 - Valor",
    "Factura Ret. 1 - Tarifa",
    "Factura Ret. 1 - Valor",
    "IVA - baseImponible",
    "IVA - porcentajeRetener",
    "IVA - valorRetenido",
    "Factura Ret. 2 - Tarifa",
    "Factura Ret. 2 - Valor",
}

EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS = [
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Tarifas IVA",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Código Documento Modificado",
    "Número Documento Modificado",
    "Fecha Emisión Doc. Sustento",
    "Motivo",
    "Valor Modificación",
    "Campos Adicionales",
    "Base Gravada 15%",
    "Monto IVA 15%",
]

EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS = {
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Código Documento Modificado",
    "Número Documento Modificado",
    "Fecha Emisión Doc. Sustento",
    "Motivo",
    "Campos Adicionales",
}

EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS = {
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Valor Modificación",
    "Base Gravada 15%",
    "Monto IVA 15%",
}

EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS = EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS
EMITIDOS_NOTA_DEBITO_TEXT_FORCE_COLUMNS = EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS
EMITIDOS_NOTA_DEBITO_NUMERIC_COLUMNS = EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS

EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL = {
    "04": "04 - RUC",
    "05": "05 - CÉDULA",
    "06": "06 - PASAPORTE",
    "07": "07 - CONSUMIDOR FINAL",
    "08": "08 - IDENTIFICACIÓN DEL EXTERIOR",
    "09": "09 - PLACA",
}

EMITIDOS_FACTURA_REPORT_COLUMNS = [
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Tarifas IVA",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Campos Adicionales",
    "Base No Gravada 0%",
]

EMITIDOS_FACTURA_TEXT_FORCE_COLUMNS = {
    "Estado",
    "Número de Autorización",
    "Fecha de Autorización",
    "Ambiente",
    "Razón Social Emisor",
    "Dir. Establecimiento",
    "Obligado Contabilidad",
    "Tipo Identificación Comprador",
    "Identificación Comprador",
    "Tipo Emisión",
    "Nombre Comercial",
    "Código del Documento",
    "Establecimiento",
    "Punto de Emisión",
    "Secuencial",
    "Dirección Matriz",
    "Contribuyente RIMPE",
    "RUC Emisor",
    "Clave de Acceso",
    "Fecha de Emisión",
    "Razón Social Comprador",
    "Dirección Comprador",
    "Moneda",
    "Plazo Pago",
    "Unidad Tiempo Pago",
    "Descripciones",
    "Forma Pago",
    "Campos Adicionales",
    "Tarifas IVA",
}

EMITIDOS_FACTURA_NUMERIC_COLUMNS = {
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Base No Gravada 0%",
}


def _emitidos_retencion_default_row() -> dict:
    row = {col: "" for col in EMITIDOS_RETENCION_REPORT_COLUMNS}
    for col in (
        "Tipo Identificación Comprador",
        "Identificación Comprador",
        "Contribuyente RIMPE",
        "Razón Social Comprador",
        "Dirección Comprador",
        "Moneda",
        "Plazo Pago",
        "Unidad Tiempo Pago",
        "Contribuyente Especial",
        "Razón Social Sujeto Retenido",
        "Identificación Sujeto Retenido",
        "Periodo Fiscal",
        "Dirección del Establecimiento",
        "Tipo Identificación Sujeto Retenido",
        "Parte Relacionada",
        "Código de Sustento",
        "Código del Documento de Sustento",
        "Número de Documento de Sustento",
        "Fecha de Emisión del Documento de Sustento",
        "Fecha de Registro Contable",
        "Número de Autorización del Documento de Sustento",
        "Pago Local o Externo",
        "Código Impuesto Doc. Sustento",
        "Código Porcentaje",
        "Forma Pago",
        "Campos Adicionales",
        "Agente de Retención",
        "Dir. Establecimiento",
        "Obligado Contabilidad",
        "Obligado a llevar Contabilidad",
    ):
        row[col] = "No Disponible"
    row["Tipo Identificación Comprador"] = "No Disponible - No Disponible"
    for col in (
        "Base Gravada",
        "Base No Gravada",
        "Monto IVA",
        "Total Descuento",
        "Propina",
        "RENTA - codigoRetencion",
        "RENTA - baseImponible",
        "RENTA - porcentajeRetener",
        "RENTA - valorRetenido",
        "Sustento Imp. 1 - Código",
        "Sustento Imp. 1 - Cod. Porcentaje",
        "Sustento Imp. 1 - Base Imponible",
        "Sustento Imp. 1 - Tarifa",
        "Sustento Imp. 1 - Valor",
        "Factura Ret. 1 - Codigo",
        "Factura Ret. 1 - Cod. Porcentaje",
        "Factura Ret. 1 - Tarifa",
        "Factura Ret. 1 - Valor",
        "IVA - codigoRetencion",
        "IVA - baseImponible",
        "IVA - porcentajeRetener",
        "IVA - valorRetenido",
        "Factura Ret. 2 - Codigo",
        "Factura Ret. 2 - Cod. Porcentaje",
        "Factura Ret. 2 - Tarifa",
        "Factura Ret. 2 - Valor",
    ):
        row[col] = 0
    row["Descripciones"] = ""
    row["Tarifas IVA"] = ""
    return row


def _texto_emitidos_retencion(valor, default: str = "") -> str:
    if valor is None:
        return default
    texto = re.sub(r"\s+", " ", str(valor).strip())
    return texto or default


def _texto_emitidos_retencion_na(valor) -> str:
    return _texto_emitidos_retencion(valor, "No Disponible")


def _numero_emitidos_retencion(valor, default=0):
    if valor in ("", None):
        return default
    if isinstance(valor, (int, float)):
        return valor
    parsed = _parse_decimal(str(valor))
    return parsed if parsed is not None else default


def _nota_credito_emitidos_default_row() -> dict:
    row = {col: "" for col in EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS}
    for col in (
        "Dir. Establecimiento",
        "Obligado Contabilidad",
        "Tipo Identificación Comprador",
        "Identificación Comprador",
        "Nombre Comercial",
        "Contribuyente RIMPE",
        "Razón Social Comprador",
        "Dirección Comprador",
        "Moneda",
        "Plazo Pago",
        "Unidad Tiempo Pago",
        "Forma Pago",
        "Código Documento Modificado",
        "Número Documento Modificado",
        "Fecha Emisión Doc. Sustento",
        "Motivo",
        "Campos Adicionales",
    ):
        row[col] = "No Disponible"
    for col in (
        "Total Sin Impuestos",
        "Base Gravada",
        "Base No Gravada",
        "Monto IVA",
        "Total Descuento",
        "Propina",
        "Importe Total",
        "Total Pago",
        "Valor Modificación",
        "Base Gravada 15%",
        "Monto IVA 15%",
    ):
        row[col] = 0
    row["Tarifas IVA"] = ""
    row["Descripciones"] = ""
    return row


def _factura_emitidos_default_row() -> dict:
    row = {col: "" for col in EMITIDOS_FACTURA_REPORT_COLUMNS}
    for col in (
        "Estado",
        "Número de Autorización",
        "Fecha de Autorización",
        "Ambiente",
        "Razón Social Emisor",
        "Dir. Establecimiento",
        "Obligado Contabilidad",
        "Tipo Identificación Comprador",
        "Identificación Comprador",
        "Tipo Emisión",
        "Nombre Comercial",
        "Código del Documento",
        "Establecimiento",
        "Punto de Emisión",
        "Secuencial",
        "Dirección Matriz",
        "Contribuyente RIMPE",
        "RUC Emisor",
        "Clave de Acceso",
        "Fecha de Emisión",
        "Razón Social Comprador",
        "Dirección Comprador",
        "Moneda",
        "Plazo Pago",
        "Unidad Tiempo Pago",
        "Forma Pago",
        "Campos Adicionales",
    ):
        row[col] = "No Disponible"
    row["Tarifas IVA"] = "0%"
    for col in (
        "Total Sin Impuestos",
        "Base Gravada",
        "Base No Gravada",
        "Monto IVA",
        "Total Descuento",
        "Propina",
        "Importe Total",
        "Total Pago",
        "Base No Gravada 0%",
    ):
        row[col] = 0
    row["Descripciones"] = ""
    return row


def _nota_debito_emitidos_default_row() -> dict:
    row = {col: "" for col in EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS}
    for col in (
        "Dir. Establecimiento",
        "Obligado Contabilidad",
        "Tipo Identificación Comprador",
        "Identificación Comprador",
        "Nombre Comercial",
        "Contribuyente RIMPE",
        "Razón Social Comprador",
        "Dirección Comprador",
        "Moneda",
        "Plazo Pago",
        "Unidad Tiempo Pago",
        "Forma Pago",
        "Código Documento Modificado",
        "Número Documento Modificado",
        "Fecha Emisión Doc. Sustento",
        "Motivo",
        "Campos Adicionales",
    ):
        row[col] = "No Disponible"
    for col in (
        "Total Sin Impuestos",
        "Base Gravada",
        "Base No Gravada",
        "Monto IVA",
        "Total Descuento",
        "Propina",
        "Importe Total",
        "Total Pago",
        "Valor Modificación",
        "Base Gravada 15%",
        "Monto IVA 15%",
    ):
        row[col] = 0
    row["Tarifas IVA"] = ""
    row["Descripciones"] = ""
    return row


def _label_tipo_ident_emitidos_nota_credito(valor: str) -> str:
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL:
        return EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL[valor]
    return valor or "No Disponible"


def _label_ambiente_emitidos_retencion(valor: str) -> str:
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_RETENCION_AMBIENTE_LABEL:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL[valor]
    valor_norm = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii").upper()
    if "PRODUCCION" in valor_norm:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL["2"]
    if "PRUEBA" in valor_norm:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL["1"]
    return valor


def _label_emision_emitidos_retencion(valor: str) -> str:
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_RETENCION_TIPO_EMISION_LABEL:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL[valor]
    valor_norm = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii").upper()
    if "NORMAL" in valor_norm:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL["1"]
    if "INDISPONIBILIDAD" in valor_norm or "CONTINGENCIA" in valor_norm:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL["2"]
    return valor


def _label_forma_pago_emitidos_retencion(valor: str) -> str:
    valor = _texto_emitidos_retencion(valor)
    if not valor:
        return "No Disponible"
    if valor in EMITIDOS_RETENCION_FORMA_PAGO_LABEL:
        return EMITIDOS_RETENCION_FORMA_PAGO_LABEL[valor]
    return valor


def _extraer_xml_emitidos_autorizacion(xml_path: Path) -> tuple[ET.Element | None, dict]:
    try:
        contenido = xml_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None, {}
    if not contenido:
        return None, {}
    try:
        root = ET.fromstring(contenido)
    except ET.ParseError:
        return None, {}
    meta = {}
    comprobante_xml = contenido
    if root.tag.lower().endswith("autorizacion"):
        meta = {
            "estado": _texto_emitidos_retencion(root.findtext("estado")),
            "numero_autorizacion": _texto_emitidos_retencion(root.findtext("numeroAutorizacion")),
            "fecha_autorizacion": _texto_emitidos_retencion(root.findtext("fechaAutorizacion")),
            "ambiente": _texto_emitidos_retencion(root.findtext("ambiente")),
        }
        comprobante_xml = root.findtext("comprobante") or ""
    try:
        comprobante_root = ET.fromstring(comprobante_xml)
    except ET.ParseError:
        return None, meta
    _strip_xml_namespaces(comprobante_root)
    return comprobante_root, meta


def _extraer_datos_xml_nota_credito_emitido(xml_path: Path) -> dict:
    row = _nota_credito_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_nc = root.find("infoNotaCredito")
    detalles = root.findall(".//detalles/detalle")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(info_trib.findtext("nombreComercial"))
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, cod_doc or "No Disponible")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_nc is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(info_nc.findtext("dirEstablecimiento"))
        row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(info_nc.findtext("obligadoContabilidad"))
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_nc.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_nc.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_nc.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_nc.findtext("razonSocialComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_nc.findtext("moneda"))
        row["Código Documento Modificado"] = _texto_emitidos_retencion_na(info_nc.findtext("codDocModificado"))
        row["Número Documento Modificado"] = _texto_emitidos_retencion_na(info_nc.findtext("numDocModificado"))
        row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(info_nc.findtext("fechaEmisionDocSustento"))
        row["Motivo"] = _texto_emitidos_retencion_na(info_nc.findtext("motivo"))
        row["Valor Modificación"] = _numero_emitidos_retencion(info_nc.findtext("valorModificacion"))
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_nc.findtext("totalSinImpuestos"))

    detalle_textos = []
    base_gravada = 0
    base_no_gravada = 0
    monto_iva = 0
    base_gravada_15 = 0
    monto_iva_15 = 0
    tarifas = []
    for detalle in detalles:
        codigo = _texto_emitidos_retencion(detalle.findtext("codigoInterno") or detalle.findtext("codigoPrincipal"))
        descripcion = (detalle.findtext("descripcion") or "").strip()
        cantidad = _texto_emitidos_retencion(detalle.findtext("cantidad"))
        precio_unitario = _texto_emitidos_retencion(detalle.findtext("precioUnitario"))
        partes = []
        if codigo:
            partes.append(f"Código: {codigo}")
        if descripcion:
            partes.append(f"Desc: {descripcion}")
        if cantidad:
            partes.append(f"Cant: {cantidad}")
        if precio_unitario:
            partes.append(f"P.Unit: {precio_unitario}")
        if partes:
            detalle_textos.append(", ".join(partes))

        for imp in detalle.findall("./impuestos/impuesto"):
            codigo = _texto_emitidos_retencion(imp.findtext("codigo"))
            codigo_pct = _texto_emitidos_retencion(imp.findtext("codigoPorcentaje"))
            tarifa = _numero_emitidos_retencion(imp.findtext("tarifa"))
            base = _numero_emitidos_retencion(imp.findtext("baseImponible"))
            valor = _numero_emitidos_retencion(imp.findtext("valor"))
            if codigo == "2":
                if codigo_pct == "0":
                    base_no_gravada += base
                else:
                    base_gravada += base
                    monto_iva += valor
                    if codigo_pct == "4":
                        base_gravada_15 += base
                        monto_iva_15 += valor
                if tarifa:
                    etiqueta = f"{int(tarifa) if float(tarifa).is_integer() else tarifa}%"
                    if etiqueta not in tarifas:
                        tarifas.append(etiqueta)

    row["Descripciones"] = " | ".join(detalle_textos)
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Total Sin Impuestos"] = row["Total Sin Impuestos"] or base_gravada or base_no_gravada
    row["Base Gravada"] = base_gravada
    row["Base No Gravada"] = base_no_gravada
    row["Tarifas IVA"] = ", ".join(tarifas)
    row["Monto IVA"] = monto_iva
    row["Importe Total"] = row["Valor Modificación"] or (row["Total Sin Impuestos"] + row["Monto IVA"])
    row["Total Pago"] = 0
    row["Base Gravada 15%"] = base_gravada_15
    row["Monto IVA 15%"] = monto_iva_15

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)

    return row


def _extraer_datos_xml_nota_debito_emitido(xml_path: Path) -> dict:
    row = _nota_debito_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_nd = root.find("infoNotaDebito")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(
            info_trib.findtext("nombreComercial") or info_trib.findtext("razonSocial")
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, "05 - NOTA DE DÉBITO")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_nd is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
            info_nd.findtext("dirEstablecimiento") or row["Dirección Matriz"]
        )
        obligado = _texto_emitidos_retencion(info_nd.findtext("obligadoContabilidad"))
        row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_nd.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_nd.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_nd.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_nd.findtext("razonSocialComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_nd.findtext("moneda") or "DOLAR")
        row["Código Documento Modificado"] = _texto_emitidos_retencion_na(info_nd.findtext("codDocModificado"))
        row["Número Documento Modificado"] = _texto_emitidos_retencion_na(info_nd.findtext("numDocModificado"))
        row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(info_nd.findtext("fechaEmisionDocSustento"))
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_nd.findtext("totalSinImpuestos"))
        row["Importe Total"] = _numero_emitidos_retencion(info_nd.findtext("valorTotal"))

        pago = info_nd.find("./pagos/pago")
        if pago is not None:
            forma = _label_forma_pago_emitidos_retencion(pago.findtext("formaPago"))
            row["Forma Pago"] = f"{forma} - {forma}" if forma != "No Disponible" else "No Disponible - No Disponible"
            row["Total Pago"] = _numero_emitidos_retencion(
                pago.findtext("total") or info_nd.findtext("valorTotal")
            )
            row["Plazo Pago"] = _texto_emitidos_retencion_na(pago.findtext("plazo"))
            row["Unidad Tiempo Pago"] = _texto_emitidos_retencion_na(pago.findtext("unidadTiempo"))
        else:
            row["Forma Pago"] = "No Disponible - No Disponible"
            row["Total Pago"] = row["Importe Total"]

    base_gravada = 0
    base_no_gravada = 0
    monto_iva = 0
    base_gravada_15 = 0
    monto_iva_15 = 0
    tarifas = []
    for imp in root.findall(".//impuestos/impuesto"):
        codigo = _texto_emitidos_retencion(imp.findtext("codigo"))
        codigo_pct = _texto_emitidos_retencion(imp.findtext("codigoPorcentaje"))
        tarifa = _numero_emitidos_retencion(imp.findtext("tarifa"))
        base = _numero_emitidos_retencion(imp.findtext("baseImponible"))
        valor = _numero_emitidos_retencion(imp.findtext("valor"))
        if codigo == "2":
            if codigo_pct == "0" or not tarifa:
                base_no_gravada += base
            else:
                base_gravada += base
                monto_iva += valor
                if codigo_pct == "4" or abs(tarifa - 15) < 0.001:
                    base_gravada_15 += base
                    monto_iva_15 += valor
            if tarifa:
                etiqueta = f"{int(tarifa) if float(tarifa).is_integer() else tarifa}%"
                if etiqueta not in tarifas:
                    tarifas.append(etiqueta)

    motivos = []
    valor_modificacion = 0
    for motivo in root.findall(".//motivos/motivo"):
        razon = _texto_emitidos_retencion_na(motivo.findtext("razon"))
        valor = _numero_emitidos_retencion(motivo.findtext("valor"))
        if razon and razon != "No Disponible":
            motivos.append(razon)
        valor_modificacion += valor

    row["Motivo"] = " | ".join(motivos) if motivos else row["Motivo"]
    row["Descripciones"] = row["Motivo"] if row["Motivo"] != "No Disponible" else row["Descripciones"]
    row["Valor Modificación"] = valor_modificacion or row["Importe Total"]
    row["Base Gravada"] = base_gravada
    row["Base No Gravada"] = base_no_gravada
    row["Tarifas IVA"] = ", ".join(tarifas)
    row["Monto IVA"] = monto_iva
    row["Base Gravada 15%"] = base_gravada_15
    row["Monto IVA 15%"] = monto_iva_15
    if not row["Total Pago"]:
        row["Total Pago"] = row["Importe Total"]

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)

    return row

def _normalizar_texto_pdf(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")

def _leer_texto_pdf(pdf_path: Path) -> str:
    if pdfplumber is None:
        try:
            from pdfminer.high_level import extract_text
        except Exception:
            return ""
        try:
            return extract_text(str(pdf_path)) or ""
        except Exception:
            return ""
    try:
        partes = []
        with pdfplumber.open(pdf_path) as pdf:
            for pagina in pdf.pages:
                try:
                    texto = pagina.extract_text() or ""
                except Exception:
                    texto = ""
                if texto:
                    partes.append(texto)
        return "\n".join(partes)
    except Exception:
        return ""

def _es_archivo_pdf(ruta: Path) -> bool:
    try:
        with ruta.open("rb") as f:
            return f.read(4) == b"%PDF"
    except Exception:
        return False

def _extraer_regex(texto: str, patrones: list[str]) -> str:
    for patron in patrones:
        encontrado = re.search(patron, texto, flags=re.IGNORECASE | re.MULTILINE)
        if encontrado:
            return (encontrado.group(1) or "").strip()
    return ""

def _extraer_monto(texto: str, patrones: list[str]) -> str:
    valor = _extraer_regex(texto, patrones)
    return valor

def _extraer_forma_pago(lineas: list[str]) -> tuple[str, str]:
    stop_words = ("SUBTOTAL", "TOTAL", "IVA", "DESCUENTO", "PROPINA", "ICE", "IRBPNR")
    for idx, linea in enumerate(lineas):
        if "FORMA DE PAGO" in linea.upper():
            for j in range(idx + 1, min(idx + 12, len(lineas))):
                cand = (lineas[j] or "").strip()
                if not cand:
                    continue
                cand_upper = cand.upper()
                if any(palabra in cand_upper for palabra in stop_words):
                    continue
                monto_match = re.search(r"([0-9][0-9.,]+)$", cand)
                if monto_match:
                    monto = monto_match.group(1)
                    forma = cand[:monto_match.start()].strip(" -")
                    if forma:
                        return forma, monto
                if re.match(r"\d{2}\s*[-–]\s*\S+", cand) or re.match(r"\d{2}\s+\S+", cand):
                    return cand, ""
            break
    return "", ""

def _extraer_seccion(texto: str, inicio: list[str], fin: list[str], max_lineas: int = 8) -> str:
    lineas = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    inicio_norm = [x.upper() for x in inicio]
    fin_norm = [x.upper() for x in fin]
    capturando = False
    recogidas = []
    for linea in lineas:
        linea_upper = linea.upper()
        if not capturando and any(token in linea_upper for token in inicio_norm):
            capturando = True
            continue
        if capturando:
            if any(token in linea_upper for token in fin_norm):
                break
            recogidas.append(linea)
            if len(recogidas) >= max_lineas:
                break
    return "; ".join(recogidas)

def _extraer_tipo_documento(texto: str) -> str:
    mapa = [
        (r"FACTURA", "Factura"),
        (r"NOTA\s+DE\s+CREDITO", "Nota de Credito"),
        (r"NOTA\s+DE\s+DEBITO", "Nota de Debito"),
        (r"LIQUIDACION\s+DE\s+COMPRAS?", "Liquidacion de compra"),
        (r"LIQUIDACION\s+DE\s+COMPRA\s+DE\s+BIENES\s+Y\s+PRESTACION\s+DE\s+SERVICIOS", "Liquidacion de compra"),
        (r"GUIA\s+DE\s+REMISION", "Guia de Remision"),
        (r"COMPROBANTE\s+DE\s+RETENCION", "Retencion"),
    ]
    for patron, valor in mapa:
        if re.search(patron, texto, flags=re.IGNORECASE):
            return valor
    return ""

def _extraer_datos_pdf(pdf_path: Path) -> dict:
    datos = {col: "" for col in PDF_REPORT_COLUMNS}
    texto = _leer_texto_pdf(pdf_path)
    if not texto:
        return datos
    texto_norm = _normalizar_texto_pdf(texto)
    texto_compacto = re.sub(r"(\d)\s+(?=\d)", r"\1", texto_norm)
    lineas = [ln.strip() for ln in texto_norm.splitlines() if ln.strip()]

    def _limpiar_valor(valor: str) -> str:
        if not valor:
            return ""
        texto_valor = str(valor).strip()
        texto_valor = unicodedata.normalize("NFKD", texto_valor).encode("ascii", "ignore").decode("ascii")
        texto_valor = re.sub(r"^[/\-\s]+", "", texto_valor)
        texto_valor = re.sub(
            r"(?i)^(RUC|RAZON SOCIAL|NOMBRES Y APELLIDOS|DIRECCION MATRIZ|DIRECCION SUCURSAL|DIRECCION|"
            r"DIRECCION COMPRADOR|DIRECCION ADQUIRENTE|MATRICULA|PLACA|GUIA|GUIA DE REMISION)\s*[:\-]?\s*",
            "",
            texto_valor,
        )
        texto_valor = re.sub(r"\s+", " ", texto_valor).strip(" /-")
        return texto_valor

    def _truncar_por_keywords(valor: str, keywords: list[str]) -> str:
        if not valor:
            return ""
        upper = valor.upper()
        corte = None
        for keyword in keywords:
            pos = upper.find(keyword)
            if pos != -1:
                if corte is None or pos < corte:
                    corte = pos
        if corte is not None and corte > 0:
            return valor[:corte].strip(" /-")
        return valor

    def _bloque_lineas(inicio: list[str], fin: list[str], max_lineas: int = 40) -> list[str]:
        inicio_norm = [x.upper() for x in inicio]
        fin_norm = [x.upper() for x in fin]
        for i, linea in enumerate(lineas):
            linea_upper = linea.upper()
            if any(token in linea_upper for token in inicio_norm):
                recogidas = []
                for j in range(i + 1, min(i + 1 + max_lineas, len(lineas))):
                    linea_j = lineas[j]
                    if any(token in linea_j.upper() for token in fin_norm):
                        break
                    recogidas.append(linea_j)
                return recogidas
        return []

    def _buscar_texto_por_etiqueta(lineas_src: list[str], etiquetas: list[str]) -> str:
        etiquetas_norm = [e.upper() for e in etiquetas]
        for idx, linea in enumerate(lineas_src):
            linea_upper = linea.upper()
            for etiqueta in etiquetas_norm:
                if etiqueta in linea_upper:
                    pos = linea_upper.find(etiqueta)
                    resto = linea[pos + len(etiqueta):].strip(" :/-")
                    if resto:
                        return _limpiar_valor(resto)
                    if idx + 1 < len(lineas_src):
                        return _limpiar_valor(lineas_src[idx + 1] or "")
        return ""

    def _buscar_valores_por_etiqueta(lineas_src: list[str], etiquetas: list[str]) -> list[str]:
        resultados = []
        etiquetas_norm = [e.upper() for e in etiquetas]
        for idx, linea in enumerate(lineas_src):
            linea_upper = linea.upper()
            for etiqueta in etiquetas_norm:
                if etiqueta in linea_upper:
                    pos = linea_upper.find(etiqueta)
                    resto = linea[pos + len(etiqueta):].strip(" :/-")
                    if not resto and idx + 1 < len(lineas_src):
                        resto = (lineas_src[idx + 1] or "").strip(" /-")
                    resto = _limpiar_valor(resto)
                    if resto:
                        resultados.append(resto)
        return list(dict.fromkeys(resultados))

    def _buscar_fecha_por_etiqueta(lineas_src: list[str], etiquetas: list[str]) -> str:
        etiquetas_norm = [e.upper() for e in etiquetas]
        patron_fecha = re.compile(r"(\d{2}[/-]\d{2}[/-]\d{4}(?:\s+\d{2}:\d{2}:\d{2})?)")
        for idx, linea in enumerate(lineas_src):
            linea_upper = linea.upper()
            if any(et in linea_upper for et in etiquetas_norm):
                match = patron_fecha.search(linea)
                if match:
                    return match.group(1)
                for j in range(idx + 1, min(idx + 4, len(lineas_src))):
                    match = patron_fecha.search(lineas_src[j])
                    if match:
                        return match.group(1)
                if idx + 1 < len(lineas_src):
                    combinado = f"{linea} {lineas_src[idx + 1]}"
                    match = patron_fecha.search(combinado)
                    if match:
                        return match.group(1)
        return ""

    def _normalizar_label(texto_label: str) -> str:
        base = unicodedata.normalize("NFKD", texto_label or "").encode("ascii", "ignore").decode("ascii")
        base = re.sub(r"[^A-Za-z0-9 ]+", " ", base).upper()
        base = re.sub(r"\s+", " ", base).strip()
        return base

    def _extraer_kv(lineas_src: list[str]) -> dict:
        kv = {}
        for linea in lineas_src:
            if ":" not in linea:
                continue
            segmentos = [seg.strip() for seg in re.split(r"\s*/\s*", linea) if seg.strip()]
            for seg in segmentos:
                if ":" not in seg:
                    continue
                label, valor = seg.split(":", 1)
                label_norm = _normalizar_label(label)
                valor = _limpiar_valor(valor)
                if label_norm and valor:
                    kv[label_norm] = valor
        return kv

    def _buscar_kv(kv: dict, etiquetas: list[str]) -> str:
        for etiqueta in etiquetas:
            valor = kv.get(_normalizar_label(etiqueta))
            if valor:
                return _limpiar_valor(valor)
        return ""

    def _buscar_numero_por_etiqueta(lineas_src: list[str], etiquetas: list[str], min_len: int, max_len: int) -> str:
        etiquetas_norm = [e.upper() for e in etiquetas]
        for idx, linea in enumerate(lineas_src):
            linea_upper = linea.upper()
            if any(et in linea_upper for et in etiquetas_norm):
                candidatos = [linea]
                if idx + 1 < len(lineas_src):
                    candidatos.append(lineas_src[idx + 1])
                texto_cand = " ".join(candidatos)
                texto_cand = re.sub(r"(\d)\s+(?=\d)", r"\1", texto_cand)
                match = re.search(rf"(\d{{{min_len},{max_len}}})", texto_cand)
                if match:
                    return match.group(1)
        return ""

    def _es_linea_emisor_valida(linea: str) -> bool:
        if not linea:
            return False
        upper = _normalizar_label(linea)
        bloqueos = [
            "RUC",
            "FACTURA",
            "AUTORIZACION",
            "NUMERO",
            "CLAVE",
            "FECHA",
            "AMBIENTE",
            "EMISION",
            "OBLIGADO",
            "AGENTE",
            "DIRECCION",
            "MATRIZ",
            "SUCURSAL",
            "LOGO",
            "COMPROBANTE",
            "RETENCION",
            "IDENTIFICACION",
            "RIMPE",
            "CONTRIBUYENTE",
        ]
        return not any(token in upper for token in bloqueos)

    def _buscar_razon_comercial(lineas_src: list[str]) -> tuple[str, str]:
        razon = ""
        comercial = ""
        for idx, linea in enumerate(lineas_src):
            candidato = _limpiar_valor(linea)
            if not candidato:
                continue
            if not _es_linea_emisor_valida(candidato) or _parece_direccion(candidato):
                continue
            razon = candidato
            for j in range(idx + 1, min(idx + 4, len(lineas_src))):
                cand2 = _limpiar_valor(lineas_src[j])
                if not cand2:
                    continue
                if not _es_linea_emisor_valida(cand2) or _parece_direccion(cand2):
                    continue
                if cand2.upper() == razon.upper():
                    continue
                comercial = cand2
                break
            break
        return razon, comercial

    def _buscar_direccion_lineas(lineas_src: list[str], etiqueta: str) -> str:
        etiqueta_norm = _normalizar_label(etiqueta)
        for idx, linea in enumerate(lineas_src):
            upper = _normalizar_label(linea)
            if "DIRECCION" in upper and etiqueta_norm in upper:
                val = _limpiar_valor(linea)
                if val and val.upper() not in {"MATRIZ", "SUCURSAL"}:
                    return val
                if idx + 1 < len(lineas_src):
                    val = _limpiar_valor(lineas_src[idx + 1])
                    if val and val.upper() not in {"MATRIZ", "SUCURSAL"}:
                        return val
            if "DIRECCION" in upper and idx + 1 < len(lineas_src):
                siguiente = lineas_src[idx + 1]
                siguiente_upper = _normalizar_label(siguiente)
                if etiqueta_norm in siguiente_upper:
                    if idx + 2 < len(lineas_src):
                        val = _limpiar_valor(lineas_src[idx + 2])
                        if val and val.upper() not in {"MATRIZ", "SUCURSAL"}:
                            return val
        return ""

    def _lineas_izquierda_pdf() -> list[str]:
        if pdfplumber is None:
            return []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                words = page.extract_words() or []
        except Exception:
            return []
        if not words:
            return []
        limite_x = page.width * 0.45
        izquierda = [w for w in words if (w.get("x0") or 0) < limite_x]
        if not izquierda:
            return []
        izquierda.sort(key=lambda w: (w.get("top", 0), w.get("x0", 0)))
        lineas = []
        actual = []
        top_actual = None
        tolerancia = 2
        for w in izquierda:
            top = w.get("top", 0)
            if top_actual is None or abs(top - top_actual) <= tolerancia:
                actual.append(w)
                if top_actual is None:
                    top_actual = top
            else:
                lineas.append(actual)
                actual = [w]
                top_actual = top
        if actual:
            lineas.append(actual)
        textos = []
        limite_y = page.height * 0.4
        for grupo in lineas:
            top = grupo[0].get("top", 0)
            if top > limite_y:
                continue
            ordenados = sorted(grupo, key=lambda w: w.get("x0", 0))
            texto = " ".join(w.get("text", "") for w in ordenados).strip()
            if not texto:
                continue
            upper = texto.upper()
            if "RAZON SOCIAL / NOMBRES" in upper or "IDENTIFICACION" in upper:
                break
            textos.append(texto)
        return textos

    def _lineas_derecha_pdf() -> list[str]:
        if pdfplumber is None:
            return []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                words = page.extract_words() or []
        except Exception:
            return []
        if not words:
            return []
        limite_x = page.width * 0.45
        derecha = [w for w in words if (w.get("x0") or 0) >= limite_x]
        if not derecha:
            return []
        derecha.sort(key=lambda w: (w.get("top", 0), w.get("x0", 0)))
        lineas = []
        actual = []
        top_actual = None
        tolerancia = 2
        for w in derecha:
            top = w.get("top", 0)
            if top_actual is None or abs(top - top_actual) <= tolerancia:
                actual.append(w)
                if top_actual is None:
                    top_actual = top
            else:
                lineas.append(actual)
                actual = [w]
                top_actual = top
        if actual:
            lineas.append(actual)
        textos = []
        limite_y = page.height * 0.45
        for grupo in lineas:
            top = grupo[0].get("top", 0)
            if top > limite_y:
                continue
            ordenados = sorted(grupo, key=lambda w: w.get("x0", 0))
            texto = " ".join(w.get("text", "") for w in ordenados).strip()
            if texto:
                textos.append(texto)
        return textos

    def _parece_direccion(texto: str) -> bool:
        upper = _normalizar_label(texto)
        if re.search(r"\d", texto):
            return True
        if re.search(r"\bS\s*N\b", upper):
            return True
        return bool(
            re.search(
                r"\b(AV|AVENIDA|CALLE|KM|NUM|NRO|SECTOR|URB|URBANIZACION|BARRIO|CIUDADELA)\b",
                upper,
            )
        )

    def _extraer_direccion_por_tipo(lineas_src: list[str], tipo: str) -> str:
        if not lineas_src:
            return ""
        tipo_upper = _normalizar_label(tipo)
        for idx, linea in enumerate(lineas_src):
            upper = _normalizar_label(linea)
            if tipo_upper not in upper:
                continue
            for offset in (-2, -1, 0, 1, 2):
                j = idx + offset
                if j < 0 or j >= len(lineas_src):
                    continue
                cand = lineas_src[j]
                cand_upper = _normalizar_label(cand)
                if "DIRECCION" in cand_upper:
                    val = _limpiar_valor(cand)
                    if val and val.upper() not in {"MATRIZ", "SUCURSAL"}:
                        return val
                if _es_linea_emisor_valida(cand) and _parece_direccion(cand):
                    return _limpiar_valor(cand)
        return ""

    def _extraer_direccion_entre_labels(lineas_src: list[str], inicio: str, fin: str) -> str:
        if not lineas_src:
            return ""
        inicio_norm = _normalizar_label(inicio)
        fin_norm = _normalizar_label(fin)
        for idx, linea in enumerate(lineas_src):
            if inicio_norm in _normalizar_label(linea):
                for j in range(idx + 1, min(idx + 6, len(lineas_src))):
                    cand_raw = lineas_src[j]
                    cand_norm = _normalizar_label(cand_raw)
                    if fin_norm in cand_norm:
                        break
                    cand = _limpiar_valor(cand_raw)
                    if cand and _parece_direccion(cand):
                        return cand
                break
        return ""

    def _extraer_direccion_por_label(lineas_src: list[str], etiqueta: str) -> str:
        if not lineas_src:
            return ""
        etiqueta_norm = _normalizar_label(etiqueta)
        for idx, linea in enumerate(lineas_src):
            if etiqueta_norm not in _normalizar_label(linea):
                continue
            for j in range(idx - 1, max(idx - 6, -1), -1):
                cand_raw = lineas_src[j]
                cand_norm = _normalizar_label(cand_raw)
                if "DIRECCION" in cand_norm and etiqueta_norm in cand_norm:
                    continue
                cand = _limpiar_valor(cand_raw)
                if cand and _parece_direccion(cand):
                    return cand
        return ""

    emisor_lineas = _bloque_lineas(
        ["DATOS DEL EMISOR", "EMISOR"],
        ["DATOS DEL COMPRADOR", "COMPRADOR", "ADQUIRIENTE", "CLIENTE"],
    )
    emisor_lineas_encontrado = bool(emisor_lineas)
    if not emisor_lineas:
        emisor_lineas = lineas
    comprador_lineas = _bloque_lineas(
        ["DATOS DEL COMPRADOR", "COMPRADOR", "ADQUIRIENTE", "CLIENTE"],
        ["DETALLE", "FORMA DE PAGO", "INFORMACION ADICIONAL", "TOTAL"],
    )
    if not comprador_lineas:
        comprador_lineas = lineas
    emisor_kv = _extraer_kv(emisor_lineas)
    comprador_kv = _extraer_kv(comprador_lineas)
    lineas_izquierda = _lineas_izquierda_pdf()
    lineas_derecha = _lineas_derecha_pdf()

    datos["tipoDocumento"] = _extraer_tipo_documento(texto_norm)

    ruc_candidatos = re.findall(r"RUC\s*[:\-]?\s*(\d{10,13})", texto_compacto, flags=re.IGNORECASE | re.MULTILINE)
    if ruc_candidatos:
        datos["rucEmisor"] = ruc_candidatos[0]
    else:
        datos["rucEmisor"] = _extraer_regex(texto_compacto, [r"RUC[^0-9]*?(\d{10,13})"])
    if not datos["rucEmisor"]:
        datos["rucEmisor"] = _buscar_kv(emisor_kv, ["RUC", "R.U.C", "R U C"])
    if not datos["rucEmisor"]:
        datos["rucEmisor"] = _buscar_numero_por_etiqueta(emisor_lineas, ["RUC", "R.U.C", "R U C"], 10, 13)
    if not datos["rucEmisor"]:
        posibles = re.findall(r"\d{10,13}", texto_compacto)
        if posibles:
            datos["rucEmisor"] = posibles[0]
    datos["rucEmisor"] = _limpiar_valor(datos["rucEmisor"])

    if emisor_lineas_encontrado:
        razon_vals = _buscar_valores_por_etiqueta(
            emisor_lineas,
            ["RAZON SOCIAL", "NOMBRES Y APELLIDOS", "RAZON SOCIAL / NOMBRES Y APELLIDOS"],
        )
        if razon_vals:
            datos["razonSocialEmisor"] = razon_vals[0]

    datos["razonSocialEmisor"] = datos["razonSocialEmisor"] or _extraer_regex(
        texto_norm,
        [r"RAZON\s+SOCIAL(?:\s+DEL\s+EMISOR|\s+EMISOR)?\s*[:\-]?\s*([^\n]+)"],
    )
    if not datos["razonSocialEmisor"]:
        datos["razonSocialEmisor"] = _buscar_kv(
            emisor_kv,
            ["RAZON SOCIAL", "NOMBRES Y APELLIDOS", "RAZON SOCIAL / NOMBRES Y APELLIDOS"],
        )
    if not datos["razonSocialEmisor"]:
        datos["razonSocialEmisor"] = _buscar_texto_por_etiqueta(
            emisor_lineas,
            ["RAZON SOCIAL", "NOMBRES Y APELLIDOS", "RAZON SOCIAL / NOMBRES Y APELLIDOS"],
        )
    datos["razonSocialEmisor"] = _limpiar_valor(datos["razonSocialEmisor"])
    razon_alt, comercial_alt = _buscar_razon_comercial(lineas_izquierda or emisor_lineas)
    if lineas_izquierda and razon_alt:
        datos["razonSocialEmisor"] = razon_alt

    datos["nombreComercial"] = _extraer_regex(texto_norm, [r"NOMBRE\s+COMERCIAL\s*[:\-]?\s*([^\n]+)"])
    if not datos["nombreComercial"]:
        datos["nombreComercial"] = _buscar_kv(emisor_kv, ["NOMBRE COMERCIAL"])
    if not datos["nombreComercial"]:
        datos["nombreComercial"] = _buscar_texto_por_etiqueta(emisor_lineas, ["NOMBRE COMERCIAL"])
    datos["nombreComercial"] = _limpiar_valor(datos["nombreComercial"])
    if not datos["nombreComercial"] and comercial_alt:
        datos["nombreComercial"] = comercial_alt

    datos["direccionMatrizEmisor"] = _extraer_regex(texto_norm, [r"DIRECCION\s+MATRIZ\s*[:\-]?\s*([^\n]+)"])
    if not datos["direccionMatrizEmisor"]:
        datos["direccionMatrizEmisor"] = _buscar_kv(emisor_kv, ["DIRECCION MATRIZ"])
    if not datos["direccionMatrizEmisor"]:
        datos["direccionMatrizEmisor"] = _buscar_texto_por_etiqueta(emisor_lineas, ["DIRECCION MATRIZ"])
    datos["direccionMatrizEmisor"] = _limpiar_valor(datos["direccionMatrizEmisor"])
    matriz_alt = _extraer_direccion_por_tipo(lineas_izquierda or lineas, "MATRIZ")
    if matriz_alt:
        datos["direccionMatrizEmisor"] = matriz_alt
    if not datos["direccionMatrizEmisor"] and not lineas_izquierda:
        datos["direccionMatrizEmisor"] = _buscar_direccion_lineas(emisor_lineas, "MATRIZ") or _buscar_direccion_lineas(lineas, "MATRIZ")

    datos["direccionSucursalEmisor"] = _extraer_regex(texto_norm, [r"DIRECCION\s+SUCURSAL\s*[:\-]?\s*([^\n]+)"])
    if not datos["direccionSucursalEmisor"]:
        datos["direccionSucursalEmisor"] = _buscar_kv(emisor_kv, ["DIRECCION SUCURSAL", "ESTABLECIMIENTO"])
    if not datos["direccionSucursalEmisor"]:
        datos["direccionSucursalEmisor"] = _buscar_texto_por_etiqueta(emisor_lineas, ["DIRECCION SUCURSAL", "ESTABLECIMIENTO"])
    datos["direccionSucursalEmisor"] = _limpiar_valor(datos["direccionSucursalEmisor"])
    sucursal_alt = _extraer_direccion_por_tipo(lineas_izquierda or lineas, "SUCURSAL")
    if not sucursal_alt:
        sucursal_alt = _extraer_direccion_entre_labels(lineas_izquierda or lineas, "MATRIZ", "SUCURSAL")
    if not sucursal_alt:
        sucursal_alt = _extraer_direccion_por_label(lineas_izquierda or lineas, "SUCURSAL")
    if sucursal_alt:
        datos["direccionSucursalEmisor"] = sucursal_alt
    elif "CLAVE DE ACCESO" in (datos["direccionSucursalEmisor"] or "").upper():
        datos["direccionSucursalEmisor"] = ""
    if not datos["direccionSucursalEmisor"] and not lineas_izquierda:
        datos["direccionSucursalEmisor"] = _buscar_direccion_lineas(emisor_lineas, "SUCURSAL") or _buscar_direccion_lineas(lineas, "SUCURSAL")
    if datos["nombreComercial"]:
        nombre_norm = _normalizar_label(datos["nombreComercial"])
        matriz_norm = _normalizar_label(datos["direccionMatrizEmisor"])
        sucursal_norm = _normalizar_label(datos["direccionSucursalEmisor"])
        if nombre_norm and (nombre_norm == matriz_norm or nombre_norm == sucursal_norm):
            datos["nombreComercial"] = ""

    datos["contribuyenteEspecial"] = _extraer_regex(
        texto_norm,
        [r"CONTRIBUYENTE\s+ESPECIAL\s*(?:NRO\.?|NO\.?)?\s*[:\-]?\s*([^\n]+)"],
    )
    if not datos["contribuyenteEspecial"]:
        datos["contribuyenteEspecial"] = _buscar_kv(emisor_kv, ["CONTRIBUYENTE ESPECIAL"])
    if not datos["contribuyenteEspecial"]:
        datos["contribuyenteEspecial"] = _buscar_texto_por_etiqueta(emisor_lineas, ["CONTRIBUYENTE ESPECIAL"])
    datos["contribuyenteEspecial"] = _limpiar_valor(datos["contribuyenteEspecial"])

    datos["agenteRetencion"] = _extraer_regex(texto_norm, [r"AGENTE\s+DE\s+RETENCION\s*[:\-]?\s*([^\n]+)"])
    if not datos["agenteRetencion"]:
        datos["agenteRetencion"] = _buscar_kv(emisor_kv, ["AGENTE DE RETENCION"])
    if not datos["agenteRetencion"]:
        datos["agenteRetencion"] = _buscar_texto_por_etiqueta(emisor_lineas, ["AGENTE DE RETENCION"])
    datos["agenteRetencion"] = _limpiar_valor(datos["agenteRetencion"])

    datos["obligadoContabilidad"] = _extraer_regex(
        texto_norm,
        [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*[:\-]?\s*([^\n]+)"],
    )
    if not datos["obligadoContabilidad"]:
        datos["obligadoContabilidad"] = _buscar_kv(emisor_kv, ["OBLIGADO A LLEVAR CONTABILIDAD"])
    if not datos["obligadoContabilidad"]:
        datos["obligadoContabilidad"] = _buscar_texto_por_etiqueta(emisor_lineas, ["OBLIGADO A LLEVAR CONTABILIDAD"])
    datos["obligadoContabilidad"] = _limpiar_valor(datos["obligadoContabilidad"])

    datos["tipoContribuyenteRIMPE"] = _extraer_regex(texto_norm, [r"RIMPE\s*[:\-]?\s*([^\n]+)"])
    if not datos["tipoContribuyenteRIMPE"]:
        datos["tipoContribuyenteRIMPE"] = _buscar_kv(emisor_kv, ["RIMPE", "TIPO CONTRIBUYENTE RIMPE"])
    if not datos["tipoContribuyenteRIMPE"]:
        datos["tipoContribuyenteRIMPE"] = _buscar_texto_por_etiqueta(emisor_lineas, ["RIMPE", "TIPO CONTRIBUYENTE RIMPE"])
    datos["tipoContribuyenteRIMPE"] = _limpiar_valor(datos["tipoContribuyenteRIMPE"])

    numero_comp = _extraer_regex(texto_compacto, [r"(\d{3}\s*-\s*\d{3}\s*-\s*\d{9})"])
    if numero_comp:
        numero_comp = re.sub(r"\\s*-\\s*", "-", numero_comp)
    datos["numeroComprobante"] = numero_comp
    if numero_comp:
        partes = numero_comp.split("-")
        if len(partes) == 3:
            datos["establecimiento"] = partes[0].zfill(3)
            datos["puntoEmision"] = partes[1].zfill(3)
            datos["secuencial"] = partes[2].zfill(9)

    datos["fechaEmision"] = _extraer_regex(texto_norm, [r"FECHA\s+EMISION\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})"])
    if not datos["fechaEmision"]:
        datos["fechaEmision"] = _buscar_fecha_por_etiqueta(
            lineas,
            ["FECHA EMISION", "FECHA DE EMISION"],
        )
    datos["fechaAutorizacion"] = _extraer_regex(
        texto_norm,
        [r"FECHA\s+AUTORIZACION\s*[:\-]?\s*([0-9/ :]+)", r"FECHA\s+Y\s+HORA\s+AUTORIZACION\s*[:\-]?\s*([0-9/ :]+)"],
    )
    if not datos["fechaAutorizacion"]:
        datos["fechaAutorizacion"] = _buscar_fecha_por_etiqueta(
            lineas,
            ["FECHA AUTORIZACION", "FECHA Y HORA AUTORIZACION", "FECHA Y HORA DE AUTORIZACION"],
        )
    if not datos["fechaAutorizacion"] and lineas_derecha:
        datos["fechaAutorizacion"] = _buscar_fecha_por_etiqueta(
            lineas_derecha,
            ["FECHA Y HORA", "FECHA Y HORA DE", "FECHA AUTORIZACION", "FECHA Y HORA DE AUTORIZACION"],
        )

    datos["razonSocialComprador"] = datos["razonSocialComprador"] or _extraer_regex(
        texto_norm,
        [r"RAZON\s+SOCIAL(?:\s+DEL\s+COMPRADOR|\s+COMPRADOR)?\s*[:\-]?\s*([^\n]+)"],
    )
    if not datos["razonSocialComprador"]:
        datos["razonSocialComprador"] = _buscar_kv(
            comprador_kv,
            ["RAZON SOCIAL", "NOMBRES Y APELLIDOS", "RAZON SOCIAL / NOMBRES Y APELLIDOS"],
        )
    if not datos["razonSocialComprador"]:
        datos["razonSocialComprador"] = _buscar_texto_por_etiqueta(
            comprador_lineas,
            ["RAZON SOCIAL", "NOMBRES Y APELLIDOS", "RAZON SOCIAL / NOMBRES Y APELLIDOS"],
        )
    datos["razonSocialComprador"] = _limpiar_valor(datos["razonSocialComprador"])

    datos["identificacionComprador"] = _extraer_regex(
        texto_compacto,
        [r"IDENTIFICACION\s+COMPRADOR\s*[:\-]?\s*(\d{5,13})", r"IDENTIFICACION\s*[:\-]?\s*(\d{5,13})"],
    )
    if not datos["identificacionComprador"]:
        datos["identificacionComprador"] = _buscar_kv(
            comprador_kv,
            ["IDENTIFICACION", "RUC", "CEDULA", "PASAPORTE"],
        )
    if not datos["identificacionComprador"]:
        datos["identificacionComprador"] = _buscar_numero_por_etiqueta(
            comprador_lineas,
            ["IDENTIFICACION", "RUC", "CEDULA", "PASAPORTE"],
            5,
            13,
        )
    if not datos["identificacionComprador"]:
        posibles = re.findall(r"\d{10,13}", texto_compacto)
        if posibles:
            datos["identificacionComprador"] = posibles[-1]
    datos["identificacionComprador"] = _limpiar_valor(datos["identificacionComprador"])

    datos["direccionComprador"] = _extraer_regex(texto_norm, [r"DIRECCION\s+COMPRADOR\s*[:\-]?\s*([^\n]+)"])
    if not datos["direccionComprador"]:
        datos["direccionComprador"] = _buscar_kv(
            comprador_kv,
            ["DIRECCION", "DIRECCION COMPRADOR", "DIRECCION ADQUIRENTE"],
        )
    if not datos["direccionComprador"]:
        datos["direccionComprador"] = _buscar_texto_por_etiqueta(
            comprador_lineas,
            ["DIRECCION", "DIRECCION COMPRADOR", "DIRECCION ADQUIRENTE"],
        )
    datos["direccionComprador"] = _truncar_por_keywords(
        _limpiar_valor(datos["direccionComprador"]),
        ["SUBTOTAL", "TOTAL", "IVA", "DESCUENTO", "P.UNIT", "P. UNIT", "P.TOTAL", "P. TOTAL"],
    )

    datos["placa"] = _extraer_regex(texto_norm, [r"PLACA\s*[:\-]?\s*([^\n]+)", r"MATRICULA\s*[:\-]?\s*([^\n]+)"])
    if not datos["placa"]:
        datos["placa"] = _buscar_kv(comprador_kv, ["PLACA", "MATRICULA"])
    if not datos["placa"]:
        datos["placa"] = _buscar_texto_por_etiqueta(comprador_lineas, ["PLACA", "MATRICULA"])
    datos["placa"] = _limpiar_valor(datos["placa"])
    if "DIRECCION" in (datos["placa"] or "").upper():
        datos["placa"] = ""
    if "GUIA" in (datos["placa"] or "").upper():
        datos["placa"] = datos["placa"].split("GUIA", 1)[0].strip(" /-")

    datos["guia"] = _extraer_regex(texto_norm, [r"GUIA\s*(?:DE\s+REMISION)?\s*[:\-]?\s*([^\n]+)"])
    if not datos["guia"]:
        datos["guia"] = _buscar_kv(comprador_kv, ["GUIA", "GUIA DE REMISION"])
    if not datos["guia"]:
        datos["guia"] = _buscar_texto_por_etiqueta(comprador_lineas, ["GUIA", "GUIA DE REMISION"])
    datos["guia"] = _limpiar_valor(datos["guia"])
    if "DIRECCION" in (datos["guia"] or "").upper():
        datos["guia"] = datos["guia"].split("DIRECCION", 1)[0].strip(" /-")
    datos["comprobanteModificado"] = _extraer_regex(texto_norm, [r"COMPROBANTE\s+MODIFICADO\s*[:\-]?\s*([^\n]+)"])
    datos["fechaEmisionModificado"] = _extraer_regex(texto_norm, [r"FECHA\s+EMISION\s+MODIFICADO\s*[:\-]?\s*([^\n]+)"])
    datos["razonModificacion"] = _extraer_regex(texto_norm, [r"RAZON\s+MODIFICACION\s*[:\-]?\s*([^\n]+)"])
    datos["valorModificacion"] = _extraer_monto(texto_norm, [r"VALOR\s+MODIFICACION\s*[:\-]?\s*([0-9.,]+)"])

    datos["descripcionesProductos"] = _extraer_seccion(
        texto_norm,
        ["DESCRIPCION", "DESCRIPCIONES", "DETALLE"],
        ["SUBTOTAL", "TOTAL", "FORMA DE PAGO", "INFORMACION ADICIONAL"],
        max_lineas=40,
    )

    datos["subtotalTarifaEspecial"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+TARIFA\s+ESPECIAL\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotal15"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+15%\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotal12"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+12%\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotal8"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+8%\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotal5"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+5%\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotal0"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+0%\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotalNoObjetoIVA"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+NO\s+OBJETO\s+IVA\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotalExentoIVA"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+EXENTO\s+IVA\s*[:\-]?\s*([0-9.,]+)"])
    datos["subtotalSinImpuestos"] = _extraer_monto(texto_norm, [r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*[:\-]?\s*([0-9.,]+)"])
    datos["totalDescuento"] = _extraer_monto(texto_norm, [r"TOTAL\s+DESCUENTO\s*[:\-]?\s*([0-9.,]+)"])
    datos["ivaTarifaEspecial"] = _extraer_monto(texto_norm, [r"IVA\s+TARIFA\s+ESPECIAL\s*[:\-]?\s*([0-9.,]+)"])
    datos["iva15"] = _extraer_monto(texto_norm, [r"IVA\s+15%\s*[:\-]?\s*([0-9.,]+)"])
    datos["iva12"] = _extraer_monto(texto_norm, [r"IVA\s+12%\s*[:\-]?\s*([0-9.,]+)"])
    datos["iva8"] = _extraer_monto(texto_norm, [r"IVA\s+8%\s*[:\-]?\s*([0-9.,]+)"])
    datos["iva5"] = _extraer_monto(texto_norm, [r"IVA\s+5%\s*[:\-]?\s*([0-9.,]+)"])
    datos["ice"] = _extraer_monto(texto_norm, [r"ICE\s*[:\-]?\s*([0-9.,]+)"])
    datos["irbpnr"] = _extraer_monto(texto_norm, [r"IRBPNR\s*[:\-]?\s*([0-9.,]+)"])
    datos["propina"] = _extraer_monto(texto_norm, [r"PROPINA\s*[:\-]?\s*([0-9.,]+)"])
    datos["valorTotal"] = _extraer_monto(texto_norm, [r"VALOR\s+TOTAL\s*[:\-]?\s*([0-9.,]+)"])
    datos["valorTotalSinSubsidio"] = _extraer_monto(texto_norm, [r"VALOR\s+TOTAL\s+SIN\s+SUBSIDIO\s*[:\-]?\s*([0-9.,]+)"])

    forma_pago, forma_pago_monto = _extraer_forma_pago(lineas)
    datos["formaPago"] = forma_pago
    if forma_pago_monto:
        monto_parsed = _parse_decimal(forma_pago_monto)
        datos["formaPagoMonto"] = monto_parsed if monto_parsed is not None else forma_pago_monto
    else:
        datos["formaPagoMonto"] = ""

    datos["ambiente"] = _extraer_regex(texto_norm, [r"AMBIENTE\s*[:\-]?\s*([^\n]+)"])
    if not datos["ambiente"]:
        datos["ambiente"] = _buscar_texto_por_etiqueta(lineas, ["AMBIENTE"])
    if datos["ambiente"].upper() == "PRODUCCION":
        datos["ambiente"] = "PRODUCCIÓN"
    datos["emision"] = _extraer_regex(texto_norm, [r"TIPO\s+EMISION\s*[:\-]?\s*([^\n]+)", r"EMISION\s*[:\-]?\s*(NORMAL|INDISPONIBILIDAD)"])
    if not datos["emision"]:
        datos["emision"] = _buscar_texto_por_etiqueta(lineas, ["EMISION", "TIPO EMISION"])
    datos["claveAcceso"] = _extraer_regex(texto_compacto, [r"CLAVE\s+DE\s+ACCESO\s*[:\-]?\s*(\d{49})"])
    if not datos["claveAcceso"]:
        datos["claveAcceso"] = _extraer_regex(texto_compacto, [r"(\d{49})"])
    if not datos["claveAcceso"]:
        solo_digitos = re.sub(r"\D+", "", texto_norm)
        match = re.search(r"(\d{49})", solo_digitos)
        if match:
            datos["claveAcceso"] = match.group(1)
    if not datos["fechaEmision"] and datos["claveAcceso"]:
        clave_digits = re.sub(r"\D+", "", datos["claveAcceso"])
        if len(clave_digits) >= 8:
            dd, mm, yyyy = clave_digits[:2], clave_digits[2:4], clave_digits[4:8]
            datos["fechaEmision"] = f"{dd}/{mm}/{yyyy}"

    datos["informacionAdicional"] = _extraer_seccion(
        texto_norm,
        ["INFORMACION ADICIONAL"],
        ["FORMA DE PAGO", "SUBTOTAL", "TOTAL"],
        max_lineas=6,
    )
    if not datos["contribuyenteEspecial"]:
        datos["contribuyenteEspecial"] = "No Disponible"
    if not datos["agenteRetencion"]:
        datos["agenteRetencion"] = "No Disponible"
    if not datos["direccionSucursalEmisor"]:
        datos["direccionSucursalEmisor"] = "No Disponible"
    if not datos["comprobanteModificado"]:
        datos["comprobanteModificado"] = "No Disponible"
    if not datos["fechaEmisionModificado"]:
        datos["fechaEmisionModificado"] = "No Disponible"
    if not datos["razonModificacion"]:
        datos["razonModificacion"] = "No Disponible"
    if not datos["valorModificacion"]:
        datos["valorModificacion"] = "No Disponible"
    if not datos["informacionAdicional"]:
        datos["informacionAdicional"] = "No Disponible"
    if datos["contribuyenteEspecial"] and datos["contribuyenteEspecial"] != "No Disponible":
        ce = datos["contribuyenteEspecial"].strip()
        if re.fullmatch(r"\d+", ce):
            datos["contribuyenteEspecial"] = str(int(ce))
    if datos["direccionSucursalEmisor"] == "No Disponible":
        sucursal_alt = _extraer_direccion_por_label(lineas, "SUCURSAL")
        if sucursal_alt:
            datos["direccionSucursalEmisor"] = sucursal_alt
    campos_ceros = [
        "subtotalTarifaEspecial",
        "subtotal15",
        "subtotal12",
        "subtotal8",
        "subtotal5",
        "subtotal0",
        "subtotalNoObjetoIVA",
        "subtotalExentoIVA",
        "subtotalSinImpuestos",
        "totalDescuento",
        "ivaTarifaEspecial",
        "iva15",
        "iva12",
        "iva8",
        "iva5",
        "ice",
        "irbpnr",
        "propina",
        "valorTotal",
        "valorTotalSinSubsidio",
    ]
    for campo in campos_ceros:
        if datos.get(campo) in ("", None):
            datos[campo] = "0"
    # Mejora con layout (cuando el texto no es fiable)
    try:
        datos_layout = _extraer_datos_pdf_layout(pdf_path)
    except Exception:
        datos_layout = {}
    if datos_layout:
        for campo, valor in datos_layout.items():
            if campo in {
                "subtotalTarifaEspecial",
                "subtotal15",
                "subtotal12",
                "subtotal8",
                "subtotal5",
                "subtotal0",
                "subtotalNoObjetoIVA",
                "subtotalExentoIVA",
                "subtotalSinImpuestos",
                "totalDescuento",
                "ivaTarifaEspecial",
                "iva15",
                "iva12",
                "iva8",
                "iva5",
                "ice",
                "irbpnr",
                "propina",
                "valorTotal",
                "valorTotalSinSubsidio",
                "formaPagoMonto",
            }:
                if valor not in ("", None):
                    datos[campo] = valor
            else:
                if not datos.get(campo) and valor not in ("", None):
                    datos[campo] = valor
    return datos

def _extraer_datos_pdf_retencion(pdf_path: Path) -> dict:
    datos = {col: "" for col in RETENCION_REPORT_COLUMNS}
    texto = _leer_texto_pdf(pdf_path)
    if not texto:
        return datos
    texto_norm = _normalizar_texto_pdf(texto)
    lineas_raw = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    lineas_norm = [_normalizar_label_simple(ln) for ln in lineas_raw]

    def _buscar_idx(token: str) -> int | None:
        for idx, norm in enumerate(lineas_norm):
            if token in norm:
                return idx
        return None

    def _linea_siguiente_valor(idx: int, saltar_tokens: list[str] | None = None) -> str:
        if idx is None:
            return ""
        saltar_tokens = saltar_tokens or []
        for j in range(idx + 1, len(lineas_raw)):
            raw = lineas_raw[j].strip()
            norm = lineas_norm[j]
            if not raw:
                continue
            if any(t in norm for t in saltar_tokens):
                continue
            return raw
        return ""

    def _linea_anterior_valor(idx: int, saltar_tokens: list[str] | None = None) -> str:
        if idx is None:
            return ""
        saltar_tokens = saltar_tokens or []
        for j in range(idx - 1, -1, -1):
            raw = lineas_raw[j].strip()
            norm = lineas_norm[j]
            if not raw:
                continue
            if any(t in norm for t in saltar_tokens):
                continue
            return raw
        return ""

    def _buscar_fecha(texto_src: str) -> str:
        match = re.search(r"(\d{2}/\d{2}/\d{4})", texto_src)
        return match.group(1) if match else ""

    def _buscar_fecha_hora(texto_src: str) -> str:
        match = re.search(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})", texto_src)
        return match.group(1) if match else ""

    def _limpiar_valor(valor: str) -> str:
        if not valor:
            return ""
        return re.sub(r"\s+", " ", valor.strip())

    datos["tipoDocumento"] = "Retencion"

    datos["rucEmisor"] = _extraer_regex(
        texto_norm,
        [r"R\.U\.C\.\s*:?(?:\s|\n)*(\d{10,13})"],
    )
    if not datos["rucEmisor"]:
        idx_ruc = _buscar_idx("R U C")
        ruc_val = _linea_siguiente_valor(idx_ruc)
        if ruc_val and re.fullmatch(r"\d{10,13}", ruc_val):
            datos["rucEmisor"] = ruc_val
        else:
            datos["rucEmisor"] = _extraer_regex(texto_norm, [r"R\.\?U\.\?C\.\?\s*[:\-]?\s*(\d{10,13})"])

    razon_emisor = ""
    idx_aut = _buscar_idx("NUMERO DE AUTORIZACION")
    if idx_aut is None:
        idx_aut = _buscar_idx("AUTORIZACION")
    idx_num_aut = None
    if idx_aut is not None:
        for j in range(idx_aut + 1, len(lineas_raw)):
            if re.fullmatch(r"\d{10,}", lineas_raw[j].strip()):
                idx_num_aut = j
                break
    if idx_num_aut is not None:
        for j in range(idx_num_aut + 1, len(lineas_raw)):
            raw = lineas_raw[j].strip()
            norm = lineas_norm[j]
            if not raw:
                continue
            if re.search(r"\d{2}/\d{2}/\d{4}", raw):
                continue
            if any(token in norm for token in ["FECHA", "AUTORIZACION", "AMBIENTE", "EMISION", "DIRECCION", "CLAVE", "OBLIGADO", "AGENTE", "COMPROBANTE", "R U C"]):
                continue
            razon_emisor = raw
            break
    datos["razonSocialEmisor"] = _limpiar_valor(razon_emisor)
    if not datos["razonSocialEmisor"]:
        idx_fecha_alt = _buscar_idx("FECHA Y HORA")
        if idx_fecha_alt is None:
            idx_fecha_alt = _buscar_idx("FECHA Y HORA DE")
        candidato = _linea_anterior_valor(
            idx_fecha_alt, saltar_tokens=["AUTORIZACION", "FECHA", "HORA", "AMBIENTE"]
        )
        if candidato and re.search(r"[A-ZÁÉÍÓÚÑ]", candidato, flags=re.IGNORECASE):
            datos["razonSocialEmisor"] = _limpiar_valor(candidato)

    idx_fecha = _buscar_idx("FECHA Y HORA")
    if idx_fecha is None:
        idx_fecha = _buscar_idx("FECHA Y HORA DE")
    if idx_fecha is not None:
        for j in range(idx_fecha + 1, len(lineas_raw)):
            fecha_hora = _buscar_fecha_hora(lineas_raw[j])
            if fecha_hora:
                datos["fechaAutorizacion"] = fecha_hora
                break
    if not datos["fechaAutorizacion"]:
        datos["fechaAutorizacion"] = _buscar_fecha_hora(texto_norm)

    idx_amb = _buscar_idx("AMBIENTE")
    if idx_amb is not None:
        nombre_comercial = _linea_anterior_valor(idx_amb, saltar_tokens=["AUTORIZACION", "FECHA", "HORA", "AMBIENTE"])
        if nombre_comercial and re.search(r"[A-ZÁÉÍÓÚÑ]", nombre_comercial, flags=re.IGNORECASE):
            if nombre_comercial != datos["razonSocialEmisor"]:
                datos["nombreComercial"] = _limpiar_valor(nombre_comercial)
        linea_amb = lineas_raw[idx_amb]
        match_amb = re.search(r"(?i)AMBIENTE\s*:?\s*(PRODUCCI[ÓO]N|PRUEBAS)", linea_amb)
        if match_amb:
            datos["ambiente"] = _limpiar_valor(match_amb.group(1))
        else:
            datos["ambiente"] = _limpiar_valor(_linea_siguiente_valor(idx_amb))

    idx_matriz = _buscar_idx("DIRECCION MATRIZ")
    if idx_matriz is not None:
        tokens_ignorar = [
            "AUTORIZACION",
            "FECHA",
            "HORA",
            "AMBIENTE",
            "EMISION",
            "DIRECCION",
            "CLAVE",
            "OBLIGADO",
            "AGENTE",
            "COMPROBANTE",
            "R U C",
            "LOGO",
        ]
        valores_ignorar = {"PRODUCCION", "PRUEBAS", "NORMAL", "CONTINGENCIA"}

        def _es_candidato_emisor(raw: str, norm: str) -> bool:
            if not raw:
                return False
            if any(token in norm for token in tokens_ignorar):
                return False
            if norm in valores_ignorar:
                return False
            if re.search(r"\d{2}/\d{2}/\d{4}", raw):
                return False
            if re.fullmatch(r"\d+", raw):
                return False
            return True

        emisor_asignado = False
        if idx_num_aut is not None and idx_num_aut + 1 < idx_matriz:
            candidatos = []
            for j in range(idx_num_aut + 1, idx_matriz):
                raw = lineas_raw[j].strip()
                norm = lineas_norm[j]
                if not _es_candidato_emisor(raw, norm):
                    continue
                candidatos.append(raw)
            if candidatos:
                datos["razonSocialEmisor"] = _limpiar_valor(candidatos[0])
                if len(candidatos) > 1 and not datos.get("nombreComercial"):
                    cand_com = _limpiar_valor(candidatos[1])
                    if cand_com and cand_com != datos["razonSocialEmisor"]:
                        datos["nombreComercial"] = cand_com
                emisor_asignado = True

        if not emisor_asignado:
            razon_idx = None
            for j in range(idx_matriz - 1, -1, -1):
                cand = lineas_raw[j].strip()
                norm = lineas_norm[j]
                if not _es_candidato_emisor(cand, norm):
                    continue
                razon_idx = j
                break
            if razon_idx is not None:
                razon_idx_prev = None
                for j in range(razon_idx - 1, -1, -1):
                    cand_prev = lineas_raw[j].strip()
                    norm_prev = lineas_norm[j]
                    if not _es_candidato_emisor(cand_prev, norm_prev):
                        continue
                    razon_idx_prev = j
                    break

                if razon_idx_prev is not None:
                    datos["razonSocialEmisor"] = _limpiar_valor(lineas_raw[razon_idx_prev])
                    if not datos.get("nombreComercial"):
                        cand_com = lineas_raw[razon_idx].strip()
                        if cand_com and cand_com != datos["razonSocialEmisor"]:
                            datos["nombreComercial"] = _limpiar_valor(cand_com)
                else:
                    datos["razonSocialEmisor"] = _limpiar_valor(lineas_raw[razon_idx])
                    if not datos.get("nombreComercial"):
                        for j in range(razon_idx + 1, idx_matriz):
                            cand = lineas_raw[j].strip()
                            norm = lineas_norm[j]
                            if not _es_candidato_emisor(cand, norm):
                                continue
                            if cand != datos["razonSocialEmisor"]:
                                datos["nombreComercial"] = _limpiar_valor(cand)
                            break

    if not datos.get("nombreComercial") and datos.get("razonSocialEmisor"):
        idx_raz = None
        for j, raw in enumerate(lineas_raw):
            if _limpiar_valor(raw) == datos["razonSocialEmisor"]:
                idx_raz = j
                break
        if idx_raz is not None:
            for j in range(idx_raz + 1, len(lineas_raw)):
                cand = lineas_raw[j].strip()
                norm = lineas_norm[j]
                if not cand:
                    continue
                if re.search(r"\d{2}/\d{2}/\d{4}", cand):
                    continue
                if any(token in norm for token in ["FECHA", "AUTORIZACION", "AMBIENTE", "EMISION", "DIRECCION", "CLAVE", "OBLIGADO", "AGENTE", "COMPROBANTE", "R U C"]):
                    continue
                if cand != datos["razonSocialEmisor"]:
                    datos["nombreComercial"] = _limpiar_valor(cand)
                break
    if not datos.get("nombreComercial"):
        datos["nombreComercial"] = "No Disponible"

    if idx_matriz is not None:
        linea_matriz = lineas_raw[idx_matriz]
        valor_matriz = ""
        match_matriz = re.search(r"(?i)direcci[oó]n\s+matriz\s*:?\s*(.+)$", linea_matriz)
        if match_matriz:
            valor_matriz = match_matriz.group(1).strip()
        if not valor_matriz:
            valor_matriz = _linea_siguiente_valor(idx_matriz)
        if valor_matriz and idx_matriz + 2 < len(lineas_raw):
            posible_cont = lineas_raw[idx_matriz + 2].strip()
            norm_cont = lineas_norm[idx_matriz + 2]
            if posible_cont and not any(
                token in norm_cont
                for token in ["DIRECCION SUCURSAL", "OBLIGADO", "AGENTE", "EMISION", "AMBIENTE", "CLAVE DE ACCESO"]
            ):
                if not re.search(r"\d{2}/\d{2}/\d{4}", posible_cont):
                    valor_matriz = f"{valor_matriz} {posible_cont}".strip()
        if idx_matriz is not None and (not valor_matriz or valor_matriz.startswith(":") or len(valor_matriz) < 8):
            piezas = []
            for j in range(idx_matriz + 1, min(idx_matriz + 5, len(lineas_raw))):
                raw = lineas_raw[j].strip()
                norm = lineas_norm[j]
                if not raw:
                    continue
                if any(
                    token in norm
                    for token in ["DIRECCION SUCURSAL", "OBLIGADO", "AGENTE", "EMISION", "AMBIENTE", "CLAVE DE ACCESO"]
                ):
                    break
                piezas.append(raw)
            if piezas:
                valor_matriz = " ".join(piezas).strip()
        datos["direccionMatrizEmisor"] = _limpiar_valor(valor_matriz)

    idx_emision = _buscar_idx("EMISION")
    if idx_emision is not None:
        linea_emision = lineas_raw[idx_emision]
        match_emision = re.search(r"(?i)EMISI[ÓO]N\s*:?\s*(NORMAL|CONTINGENCIA)", linea_emision)
        if match_emision:
            datos["emision"] = _limpiar_valor(match_emision.group(1))
        else:
            datos["emision"] = _limpiar_valor(_linea_siguiente_valor(idx_emision))

    idx_sucursal = _buscar_idx("DIRECCION SUCURSAL")
    if idx_sucursal is not None:
        linea_sucursal = lineas_raw[idx_sucursal]
        valor_sucursal = ""
        match_sucursal = re.search(r"(?i)direcci[oó]n\s+sucursal\s*:?\s*(.+)$", linea_sucursal)
        if match_sucursal:
            valor_sucursal = match_sucursal.group(1).strip()
        if not valor_sucursal:
            valor_sucursal = _linea_siguiente_valor(idx_sucursal)
        if valor_sucursal and idx_sucursal + 2 < len(lineas_raw):
            posible_cont = lineas_raw[idx_sucursal + 2].strip()
            norm_cont = lineas_norm[idx_sucursal + 2]
            if posible_cont and not any(
                token in norm_cont
                for token in ["DIRECCION MATRIZ", "OBLIGADO", "AGENTE", "EMISION", "AMBIENTE", "CLAVE DE ACCESO"]
            ):
                if not re.search(r"\d{2}/\d{2}/\d{4}", posible_cont):
                    valor_sucursal = f"{valor_sucursal} {posible_cont}".strip()
        if idx_sucursal is not None and (not valor_sucursal or valor_sucursal.startswith(":") or len(valor_sucursal) < 8):
            piezas = []
            for j in range(idx_sucursal + 1, min(idx_sucursal + 5, len(lineas_raw))):
                raw = lineas_raw[j].strip()
                norm = lineas_norm[j]
                if not raw:
                    continue
                if any(
                    token in norm
                    for token in ["DIRECCION MATRIZ", "OBLIGADO", "AGENTE", "EMISION", "AMBIENTE", "CLAVE DE ACCESO"]
                ):
                    break
                piezas.append(raw)
            if piezas:
                valor_sucursal = " ".join(piezas).strip()
        datos["direccionSucursalEmisor"] = _limpiar_valor(valor_sucursal)

    idx_obligado = _buscar_idx("OBLIGADO A LLEVAR CONTABILIDAD")
    if idx_obligado is not None:
        linea_obligado = lineas_raw[idx_obligado]
        match_obligado = re.search(r"(?i)OBLIGADO A LLEVAR CONTABILIDAD\s*(SI|NO)\b", linea_obligado)
        if match_obligado:
            datos["obligadoContabilidad"] = match_obligado.group(1).upper()
        else:
            for j in range(idx_obligado + 1, len(lineas_raw)):
                val = lineas_raw[j].strip()
                if not val:
                    continue
                val_clean = re.sub(r"[^A-Z]", "", val.upper())
                if val_clean in {"SI", "NO"}:
                    datos["obligadoContabilidad"] = val_clean
                    break

    datos["numeroContribuyenteEspecial"] = _extraer_regex(
        texto_norm,
        [r"CONTRIBUYENTE\s+ESPECIAL[^0-9]*?(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*NO\.?\s*(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*No\.?\s*(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*#\s*(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*:\s*(\d+)", r"CONTRIBUYENTE\s+ESPECIAL\s*NO\s*(\d+)"],
    )
    if datos["numeroContribuyenteEspecial"]:
        ce = datos["numeroContribuyenteEspecial"].strip()
        if re.fullmatch(r"\d+", ce):
            datos["numeroContribuyenteEspecial"] = str(int(ce))
    else:
        datos["numeroContribuyenteEspecial"] = "No Disponible"

    idx_agente = _buscar_idx("AGENTE DE RETENCION")
    if idx_agente is not None:
        linea_agente = lineas_raw[idx_agente]
        match_agente = re.search(r"(?i)AGENTE DE RETENCION[^0-9]*([0-9]+)", linea_agente)
        if match_agente:
            datos["numeroAgenteRetencion"] = match_agente.group(1)
        else:
            for j in range(idx_agente + 1, len(lineas_raw)):
                agente_val = lineas_raw[j].strip()
                if not agente_val:
                    continue
                if re.fullmatch(r"\d+", agente_val):
                    datos["numeroAgenteRetencion"] = agente_val
                    break

    if not datos.get("numeroAgenteRetencion"):
        datos["numeroAgenteRetencion"] = "No Disponible"

    datos["ambiente"] = datos.get("ambiente", "").strip()
    if not datos["ambiente"]:
        datos["ambiente"] = _extraer_regex(texto_norm, [r"AMBIENTE\s*[:\-]?\s*([A-Z??????]+)"])
    if datos["ambiente"].upper() == "PRODUCCION":
        datos["ambiente"] = "PRODUCCI?N"

    if not datos.get("emision"):
        datos["emision"] = _extraer_regex(texto_norm, [r"EMISION\s*[:\-]?\s*([A-Z??????]+)"])

    datos["numeroComprobante"] = _extraer_regex(
        texto_norm,
        [r"No\.\s*(\d{3}-\d{3}-\d{6,9})", r"(\d{3}-\d{3}-\d{6,9})"],
    )
    if datos["numeroComprobante"]:
        partes = datos["numeroComprobante"].split("-")
        if len(partes) == 3:
            datos["establecimiento"] = partes[0]
            datos["puntoEmision"] = partes[1]
            datos["secuencial"] = partes[2]

    idx_razon = _buscar_idx("RAZON SOCIAL")
    if idx_razon is not None:
        linea_razon = lineas_raw[idx_razon]
        match_razon = re.search(
            r"(?i)RAZON\s+SOCIAL\s*/\s*NOMBRES\s+Y\s+APELLIDOS\s*:?\s*(.+)$",
            linea_razon,
        )
        if match_razon and match_razon.group(1).strip():
            datos["razonSocialSujetoRetenido"] = _limpiar_valor(match_razon.group(1))
        # Extraer bloque de sujeto retenido desde la etiqueta
        bloque = []
        for j in range(idx_razon + 1, len(lineas_raw)):
            norm = lineas_norm[j]
            if any(token in norm for token in ["COMPROBANTE", "BASE IMPONIBLE", "IMPUESTO", "PORCENTAJE", "RETENCION", "VALORRETENIDO", "INFORMACION ADICIONAL"]):
                break
            bloque.append(lineas_raw[j].strip())
            if len(bloque) > 12:
                break

        razon = datos.get("razonSocialSujetoRetenido", "")
        identificacion = datos.get("identificacionSujetoRetenido", "")
        fecha_emision = datos.get("fechaEmision", "")

        for item in bloque:
            if not item:
                continue
            norm_item = _normalizar_label_simple(item)
            if "IDENTIFICACION" in norm_item or norm_item == "FECHA":
                continue
            fecha_val = _buscar_fecha(item)
            if fecha_val and not fecha_emision:
                fecha_emision = fecha_val
                continue
            if re.fullmatch(r"\d{10,13}", item) and not identificacion:
                identificacion = item
                continue
            if not razon and re.search(r"[A-ZÁÉÍÓÚÑ]", item, flags=re.IGNORECASE):
                razon = item

        if razon:
            datos["razonSocialSujetoRetenido"] = _limpiar_valor(razon)
        if identificacion:
            datos["identificacionSujetoRetenido"] = identificacion
        if fecha_emision:
            datos["fechaEmision"] = fecha_emision

    if not datos.get("identificacionSujetoRetenido"):
        idx_ident = _buscar_idx("IDENTIFICACION")
        if idx_ident is not None:
            linea_ident = lineas_raw[idx_ident]
            match_ident = re.search(r"(?i)IDENTIFICACION\s*:?\s*(\d{5,13})", linea_ident)
            if match_ident:
                datos["identificacionSujetoRetenido"] = match_ident.group(1)
            else:
                for j in range(idx_ident + 1, len(lineas_raw)):
                    raw = lineas_raw[j].strip()
                    if not raw:
                        continue
                    if re.search(r"\d{5,13}", raw):
                        datos["identificacionSujetoRetenido"] = re.search(r"\d{5,13}", raw).group(0)
                        break

    if not datos.get("fechaEmision"):
        idx_fecha_suj = _buscar_idx("FECHA")
        if idx_fecha_suj is not None:
            linea_fecha = lineas_raw[idx_fecha_suj]
            match_fecha = re.search(r"(\d{2}/\d{2}/\d{4})", linea_fecha)
            if match_fecha:
                datos["fechaEmision"] = match_fecha.group(1)
            else:
                for j in range(idx_fecha_suj + 1, len(lineas_raw)):
                    fecha_val = _buscar_fecha(lineas_raw[j])
                    if fecha_val:
                        datos["fechaEmision"] = fecha_val
                        break

    if not datos.get("razonSocialSujetoRetenido"):
        bloque_sujeto = re.search(
            r"RAZON\s+SOCIAL\s*/\s*NOMBRES\s+Y\s+APELLIDOS\s*:\s*IDENTIFICACION\s*FECHA\s*([A-Z0-9 .,/\\-]+?)\s*(\d{10,13})\s*([0-3]\d/[01]\d/\d{4})",
            texto_norm,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if bloque_sujeto:
            datos["razonSocialSujetoRetenido"] = _limpiar_valor(bloque_sujeto.group(1))
            if not datos.get("identificacionSujetoRetenido"):
                datos["identificacionSujetoRetenido"] = bloque_sujeto.group(2)
            if not datos.get("fechaEmision"):
                datos["fechaEmision"] = bloque_sujeto.group(3)

    datos["claveAcceso"] = _extraer_regex(texto_norm, [r"(\d{49})"])
    if not datos.get("rucEmisor") and datos.get("claveAcceso") and len(datos["claveAcceso"]) == 49:
        datos["rucEmisor"] = datos["claveAcceso"][10:23]

    fecha_emision = ""
    idx_fecha_emision = _buscar_idx("FECHA EMISION")
    if idx_fecha_emision is None:
        for i, norm in enumerate(lineas_norm):
            if norm == "FECHA" and i > 0 and "IDENTIFICACION" in lineas_norm[i - 1]:
                idx_fecha_emision = i
                break
    if idx_fecha_emision is not None:
        for j in range(idx_fecha_emision + 1, len(lineas_raw)):
            fecha_val = _buscar_fecha(lineas_raw[j])
            if fecha_val:
                fecha_emision = fecha_val
                break
    if fecha_emision:
        datos["fechaEmision"] = fecha_emision

    # Tabla de retenciones
    tabla_inicio = None
    tabla_fin = None
    for idx, norm in enumerate(lineas_norm):
        if norm == "COMPROBANTE":
            tabla_inicio = idx + 1
            break
    if tabla_inicio is None:
        for idx, norm in enumerate(lineas_norm):
            if "COMPROBANTE" in norm and "IMPUESTO" in norm:
                tabla_inicio = idx + 1
                break
    if tabla_inicio is not None:
        for idx in range(tabla_inicio, len(lineas_norm)):
            if "INFORMACION ADICIONAL" in lineas_norm[idx] or lineas_norm[idx].startswith("INFORMACION"):
                tabla_fin = idx
                break
    tabla_lineas = lineas_raw[tabla_inicio:tabla_fin] if tabla_inicio is not None else []

    comprobante_sustento = ""
    numero_sustento = ""
    fecha_emision_sustento = ""
    ejercicio_fiscal = ""

    tokens_header = [
        "NUMERO",
        "FECHA",
        "EJERCICIO",
        "FISCAL",
        "BASE",
        "IMPUESTO",
        "PORCENTAJE",
        "RETENCION",
        "VALORRETENIDO",
    ]
    for idx, linea in enumerate(tabla_lineas):
        norm = _normalizar_label_simple(linea)
        if any(tok in norm for tok in tokens_header):
            continue
        if re.search(r"[A-ZÁÉÍÓÚÑ]", linea, flags=re.IGNORECASE):
            comprobante_sustento = linea.strip()
            siguiente = idx + 1
            while siguiente < len(tabla_lineas):
                val = tabla_lineas[siguiente].strip()
                if re.search(r"\d{2}/\d{2}/\d{4}", val):
                    break
                if re.fullmatch(r"\d+", val):
                    numero_sustento += val
                    siguiente += 1
                    continue
                if numero_sustento:
                    break
                siguiente += 1
            break

    for linea in tabla_lineas:
        if not fecha_emision_sustento:
            fecha_emision_sustento = _buscar_fecha(linea)
        if not ejercicio_fiscal:
            match_ej = re.search(r"(\d{2}/\d{4})", linea)
            if match_ej:
                ejercicio_fiscal = match_ej.group(1)
        if fecha_emision_sustento and ejercicio_fiscal:
            break

    datos["Comprobante_Sustento"] = _limpiar_valor(comprobante_sustento)
    if numero_sustento:
        numero_sustento = numero_sustento.strip()
    datos["Numero_Sustento"] = numero_sustento
    datos["Fecha_Emision_Sustento"] = fecha_emision_sustento
    datos["Ejercicio_Fiscal"] = ejercicio_fiscal

    def _es_numero_valor(linea: str) -> bool:
        if "/" in linea:
            return False
        return re.fullmatch(r"\d+(?:[.,]\d+)?", linea.strip()) is not None

    filas = []
    usados_base = set()
    i = 0
    while i < len(tabla_lineas):
        raw = tabla_lineas[i].strip()
        norm = _normalizar_label_simple(raw)
        impuesto = ""
        if "IMPUESTO A LA" in norm:
            # Puede venir en dos líneas: "Impuesto a la" + "Renta"
            if i + 1 < len(tabla_lineas) and "RENTA" in _normalizar_label_simple(tabla_lineas[i + 1]):
                impuesto = "Renta"
                i += 1
        elif "RENTA" in norm:
            impuesto = "Renta"
        elif norm == "IVA" or " IVA" in norm:
            impuesto = "IVA"

        if not impuesto:
            i += 1
            continue

        base = ""
        for j in range(i - 1, -1, -1):
            cand = tabla_lineas[j].strip()
            if not _es_numero_valor(cand):
                continue
            if len(re.sub(r"[.,]", "", cand)) >= 9:
                continue
            if j in usados_base:
                continue
            base = cand
            usados_base.add(j)
            break

        porcentaje = ""
        valor = ""
        j = i + 1
        while j < len(tabla_lineas) and (not porcentaje or not valor):
            cand = tabla_lineas[j].strip()
            if _es_numero_valor(cand):
                if not porcentaje:
                    porcentaje = cand
                elif not valor:
                    valor = cand
            j += 1

        filas.append(
            {
                "impuesto": impuesto,
                "base": base,
                "porcentaje": porcentaje,
                "valor": valor,
            }
        )
        i = j

    iva_items = [r for r in filas if r["impuesto"] == "IVA"]
    renta_items = [r for r in filas if r["impuesto"] == "Renta"]

    def _asignar_retencion(items, base_key: str, imp_key: str, porc_key: str, val_key: str, imp_label: str):
        if items:
            datos[base_key] = items[0]["base"]
            datos[imp_key] = imp_label
            datos[porc_key] = items[0]["porcentaje"]
            datos[val_key] = items[0]["valor"]

    _asignar_retencion(iva_items, "Base_Imponible_Ret_IVA", "Impuesto_Ret_IVA", "Porcentaje_Ret_IVA", "Valor_Retenido_IVA", "IVA")
    _asignar_retencion(renta_items, "Base_Imponible_Ret_IR", "Impuesto_Ret_IR", "Porcentaje_Ret_IR", "Valor_Retenido_IR", "Impuesto a la Renta")

    if len(iva_items) > 1:
        _asignar_retencion(iva_items[1:], "Base_Imponible_Ret_IVA_1", "Impuesto_Ret_IVA_1", "Porcentaje_Ret_IVA_1", "Valor_Retenido_IVA_1", "IVA")
    elif len(iva_items) == 1:
        _asignar_retencion(iva_items, "Base_Imponible_Ret_IVA_1", "Impuesto_Ret_IVA_1", "Porcentaje_Ret_IVA_1", "Valor_Retenido_IVA_1", "IVA")
    if len(renta_items) > 1:
        _asignar_retencion(renta_items[1:], "Base_Imponible_Ret_IR_1", "Impuesto_Ret_IR_1", "Porcentaje_Ret_IR_1", "Valor_Retenido_IR_1", "Impuesto a la Renta")
    elif len(renta_items) == 1:
        _asignar_retencion(renta_items, "Base_Imponible_Ret_IR_1", "Impuesto_Ret_IR_1", "Porcentaje_Ret_IR_1", "Valor_Retenido_IR_1", "Impuesto a la Renta")

    info_idx = None
    for idx, norm in enumerate(lineas_norm):
        if "INFORMACION ADICIONAL" in norm:
            info_idx = idx
    if info_idx is not None:
        adicionales = [ln.strip() for ln in lineas_raw[info_idx + 1:] if ln.strip()]
        datos["informacionAdicional"] = "\n".join(adicionales)
    if not datos["informacionAdicional"]:
        datos["informacionAdicional"] = "No Disponible"

    campos_ceros = [
        "Base_Imponible_Ret_IVA",
        "Porcentaje_Ret_IVA",
        "Valor_Retenido_IVA",
        "Base_Imponible_Ret_IR",
        "Porcentaje_Ret_IR",
        "Valor_Retenido_IR",
        "Base_Imponible_Ret_IR_1",
        "Porcentaje_Ret_IR_1",
        "Valor_Retenido_IR_1",
        "Base_Imponible_Ret_IVA_1",
        "Porcentaje_Ret_IVA_1",
        "Valor_Retenido_IVA_1",
    ]
    for campo in campos_ceros:
        if datos.get(campo) in ("", None):
            datos[campo] = "0"
    if not datos["numeroContribuyenteEspecial"]:
        datos["numeroContribuyenteEspecial"] = "No Disponible"
    return datos


def _extraer_datos_xml_retencion(xml_path: Path) -> dict:
    datos = {col: "" for col in RETENCION_REPORT_COLUMNS}
    try:
        contenido = xml_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return datos
    if not contenido:
        return datos

    meta = {}
    comprobante_xml = contenido
    if "<autorizacion" in contenido.lower():
        try:
            comprobante_xml, meta = _extraer_comprobante_desde_autorizacion(contenido)
        except Exception:
            comprobante_xml = contenido

    try:
        root = ET.fromstring(comprobante_xml)
    except ET.ParseError:
        return datos

    _strip_xml_namespaces(root)

    def _limpiar_valor(valor: str) -> str:
        if not valor:
            return ""
        return re.sub(r"\s+", " ", valor.strip())

    def _map_ambiente(valor: str) -> str:
        if not valor:
            return ""
        v = valor.strip().upper()
        if v == "1":
            return "PRUEBAS"
        if v == "2":
            return "PRODUCCION"
        return v

    def _map_emision(valor: str) -> str:
        if not valor:
            return ""
        v = valor.strip().upper()
        if v == "1":
            return "NORMAL"
        if v == "2":
            return "CONTINGENCIA"
        return v

    info_trib = root.find("infoTributaria") or root.find(".//infoTributaria")
    info_comp = root.find("infoCompRetencion") or root.find(".//infoCompRetencion")

    if info_trib is not None:
        datos["rucEmisor"] = _limpiar_valor(info_trib.findtext("ruc", ""))
        datos["razonSocialEmisor"] = _limpiar_valor(info_trib.findtext("razonSocial", ""))
        datos["nombreComercial"] = _limpiar_valor(info_trib.findtext("nombreComercial", ""))
        datos["direccionMatrizEmisor"] = _limpiar_valor(info_trib.findtext("dirMatriz", ""))
        datos["claveAcceso"] = _limpiar_valor(info_trib.findtext("claveAcceso", ""))
        datos["ambiente"] = _map_ambiente(meta.get("ambiente") or info_trib.findtext("ambiente", ""))
        datos["emision"] = _map_emision(info_trib.findtext("tipoEmision", ""))
        datos["establecimiento"] = _limpiar_valor(info_trib.findtext("estab", ""))
        datos["puntoEmision"] = _limpiar_valor(info_trib.findtext("ptoEmi", ""))
        datos["secuencial"] = _limpiar_valor(info_trib.findtext("secuencial", ""))
        if datos["establecimiento"] and datos["puntoEmision"] and datos["secuencial"]:
            datos["numeroComprobante"] = f"{datos['establecimiento']}-{datos['puntoEmision']}-{datos['secuencial']}"

    if info_comp is not None:
        datos["direccionSucursalEmisor"] = _limpiar_valor(info_comp.findtext("dirEstablecimiento", ""))
        datos["obligadoContabilidad"] = _limpiar_valor(info_comp.findtext("obligadoContabilidad", ""))
        contrib = _limpiar_valor(info_comp.findtext("contribuyenteEspecial", ""))
        datos["numeroContribuyenteEspecial"] = contrib or "No Disponible"
        agente = _limpiar_valor(info_comp.findtext("agenteRetencion", ""))
        datos["numeroAgenteRetencion"] = agente or "No Disponible"
        datos["fechaEmision"] = _limpiar_valor(info_comp.findtext("fechaEmision", ""))
        datos["razonSocialSujetoRetenido"] = _limpiar_valor(info_comp.findtext("razonSocialSujetoRetenido", ""))
        datos["identificacionSujetoRetenido"] = _limpiar_valor(info_comp.findtext("identificacionSujetoRetenido", ""))
        datos["Ejercicio_Fiscal"] = _limpiar_valor(
            info_comp.findtext("periodoFiscal", "") or info_comp.findtext("ejercicioFiscal", "")
        )

    if not datos.get("nombreComercial"):
        datos["nombreComercial"] = "No Disponible"

    datos["fechaAutorizacion"] = _limpiar_valor(
        meta.get("fecha_autorizacion") or meta.get("fechaAutorizacion") or meta.get("fecha_autorizacion", "")
    )
    if not datos["fechaAutorizacion"]:
        datos["fechaAutorizacion"] = _limpiar_valor(meta.get("fechaAutorizacion", ""))

    if not datos.get("claveAcceso"):
        clave = root.findtext(".//claveAcceso", "")
        datos["claveAcceso"] = _limpiar_valor(clave)

    DOC_SUSTENTO_LABELS = {
        "01": "FACTURA",
        "00": "OTROS",
        "03": "LIQUIDACION DE COMPRA",
        "04": "NOTA DE CREDITO",
        "05": "NOTA DE DEBITO",
        "06": "GUIA DE REMISION",
        "07": "COMPROBANTE DE RETENCION",
    }

    # Sustento desde docsSustento/docSustento (variante principal)
    doc_sustento = root.find(".//docsSustento/docSustento")
    if doc_sustento is not None:
        cod_doc = _limpiar_valor(doc_sustento.findtext("codDocSustento", ""))
        if cod_doc in DOC_SUSTENTO_LABELS:
            datos["Comprobante_Sustento"] = DOC_SUSTENTO_LABELS[cod_doc]
        elif cod_doc:
            datos["Comprobante_Sustento"] = cod_doc
        numero_raw = _limpiar_valor(doc_sustento.findtext("numDocSustento", ""))
        if numero_raw and numero_raw.isdigit() and len(numero_raw) == 15:
            datos["Numero_Sustento"] = f"{numero_raw[:3]}-{numero_raw[3:6]}-{numero_raw[6:]}"
        else:
            datos["Numero_Sustento"] = numero_raw
        datos["Fecha_Emision_Sustento"] = _limpiar_valor(doc_sustento.findtext("fechaEmisionDocSustento", ""))
    else:
        # Variante 1: impuestos/impuesto/codDocSustento
        first_imp = root.find(".//impuestos/impuesto")
        if first_imp is not None:
            cod_doc = _limpiar_valor(
                first_imp.findtext("codDocSustento", "")
                or first_imp.findtext("tipoDocumento", "")
                or first_imp.findtext("codDoc", "")
            )
            if cod_doc in DOC_SUSTENTO_LABELS:
                datos["Comprobante_Sustento"] = DOC_SUSTENTO_LABELS[cod_doc]
            elif cod_doc:
                datos["Comprobante_Sustento"] = cod_doc

    # Retenciones (IVA/IR) desde docsSustento si existe, si no desde impuestos
    retenciones = []
    for ret in root.findall(".//docsSustento/docSustento/retenciones/retencion"):
        codigo = _limpiar_valor(ret.findtext("codigo", ""))
        base = _limpiar_valor(ret.findtext("baseImponible", ""))
        porcentaje = _limpiar_valor(ret.findtext("porcentajeRetener", ""))
        valor = _limpiar_valor(ret.findtext("valorRetenido", ""))
        retenciones.append({"codigo": codigo, "base": base, "porcentaje": porcentaje, "valor": valor})

    if not retenciones:
        for imp in root.findall(".//impuestos/impuesto"):
            codigo = _limpiar_valor(imp.findtext("codigo", ""))
            base = _limpiar_valor(imp.findtext("baseImponible", ""))
            porcentaje = _limpiar_valor(imp.findtext("porcentajeRetener", ""))
            valor = _limpiar_valor(imp.findtext("valorRetenido", ""))
            retenciones.append({"codigo": codigo, "base": base, "porcentaje": porcentaje, "valor": valor})
            if not datos.get("Comprobante_Sustento"):
                cod_doc = _limpiar_valor(
                    imp.findtext("codDocSustento", "")
                    or imp.findtext("tipoDocumento", "")
                    or imp.findtext("codDoc", "")
                )
                if cod_doc in DOC_SUSTENTO_LABELS:
                    datos["Comprobante_Sustento"] = DOC_SUSTENTO_LABELS[cod_doc]
                elif cod_doc:
                    datos["Comprobante_Sustento"] = cod_doc

    iva_items = [r for r in retenciones if r["codigo"] == "2"]
    renta_items = [r for r in retenciones if r["codigo"] == "1"]

    def _asignar_retencion(items, base_key: str, imp_key: str, porc_key: str, val_key: str, imp_label: str):
        if items:
            datos[base_key] = items[0]["base"]
            datos[imp_key] = imp_label
            datos[porc_key] = items[0]["porcentaje"]
            datos[val_key] = items[0]["valor"]

    _asignar_retencion(iva_items, "Base_Imponible_Ret_IVA", "Impuesto_Ret_IVA", "Porcentaje_Ret_IVA", "Valor_Retenido_IVA", "IVA")
    _asignar_retencion(renta_items, "Base_Imponible_Ret_IR", "Impuesto_Ret_IR", "Porcentaje_Ret_IR", "Valor_Retenido_IR", "Impuesto a la Renta")

    # Si no hay segunda ocurrencia, copiar la primera (segun plantilla)
    if len(iva_items) > 1:
        _asignar_retencion(iva_items[1:], "Base_Imponible_Ret_IVA_1", "Impuesto_Ret_IVA_1", "Porcentaje_Ret_IVA_1", "Valor_Retenido_IVA_1", "IVA")
    elif datos.get("Base_Imponible_Ret_IVA"):
        datos["Base_Imponible_Ret_IVA_1"] = datos["Base_Imponible_Ret_IVA"]
        datos["Impuesto_Ret_IVA_1"] = datos.get("Impuesto_Ret_IVA", "")
        datos["Porcentaje_Ret_IVA_1"] = datos.get("Porcentaje_Ret_IVA", "")
        datos["Valor_Retenido_IVA_1"] = datos.get("Valor_Retenido_IVA", "")

    if len(renta_items) > 1:
        _asignar_retencion(renta_items[1:], "Base_Imponible_Ret_IR_1", "Impuesto_Ret_IR_1", "Porcentaje_Ret_IR_1", "Valor_Retenido_IR_1", "Impuesto a la Renta")
    elif datos.get("Base_Imponible_Ret_IR"):
        datos["Base_Imponible_Ret_IR_1"] = datos["Base_Imponible_Ret_IR"]
        datos["Impuesto_Ret_IR_1"] = datos.get("Impuesto_Ret_IR", "")
        datos["Porcentaje_Ret_IR_1"] = datos.get("Porcentaje_Ret_IR", "")
        datos["Valor_Retenido_IR_1"] = datos.get("Valor_Retenido_IR", "")

    info_adicional = root.find(".//infoAdicional")
    if info_adicional is not None:
        campos = []
        for campo in info_adicional.findall("campoAdicional"):
            nombre = _limpiar_valor(campo.attrib.get("nombre", ""))
            valor = _limpiar_valor(campo.text or "")
            if nombre or valor:
                campos.append(f"{nombre}: {valor}" if nombre else valor)
        datos["informacionAdicional"] = "\n".join(campos)
    if not datos["informacionAdicional"]:
        datos["informacionAdicional"] = "No Disponible"

    datos["tipoDocumento"] = "Retencion"

    campos_ceros = [
        "Base_Imponible_Ret_IVA",
        "Porcentaje_Ret_IVA",
        "Valor_Retenido_IVA",
        "Base_Imponible_Ret_IR",
        "Porcentaje_Ret_IR",
        "Valor_Retenido_IR",
        "Base_Imponible_Ret_IR_1",
        "Porcentaje_Ret_IR_1",
        "Valor_Retenido_IR_1",
        "Base_Imponible_Ret_IVA_1",
        "Porcentaje_Ret_IVA_1",
        "Valor_Retenido_IVA_1",
    ]
    for campo in campos_ceros:
        if datos.get(campo) in ("", None):
            datos[campo] = "0"

    if not datos.get("numeroContribuyenteEspecial"):
        datos["numeroContribuyenteEspecial"] = "No Disponible"
    if not datos.get("numeroAgenteRetencion"):
        datos["numeroAgenteRetencion"] = "No Disponible"

    return datos


def _extraer_datos_xml_retencion_emitido(xml_path: Path) -> dict:
    row = _emitidos_retencion_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_comp = root.find("infoCompRetencion")
    doc_sustento = root.find(".//docsSustento/docSustento")

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(info_trib.findtext("nombreComercial"))
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(
            info_trib.findtext("ambiente") or meta.get("ambiente")
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, cod_doc or "No Disponible")
        row["Agente de Retención"] = _texto_emitidos_retencion_na(
            info_trib.findtext("agenteRetencion") or root.findtext(".//agenteRetencion")
        )

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(
        meta.get("numero_autorizacion") or row.get("Clave de Acceso")
    )
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_comp is not None:
        dir_est = _texto_emitidos_retencion_na(info_comp.findtext("dirEstablecimiento"))
        obligado = _texto_emitidos_retencion_na(info_comp.findtext("obligadoContabilidad"))
        row["Dir. Establecimiento"] = dir_est
        row["Obligado Contabilidad"] = obligado
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_comp.findtext("fechaEmision"))
        row["Contribuyente Especial"] = _texto_emitidos_retencion_na(info_comp.findtext("contribuyenteEspecial"))
        row["Obligado a llevar Contabilidad"] = obligado
        row["Razón Social Sujeto Retenido"] = _texto_emitidos_retencion_na(
            info_comp.findtext("razonSocialSujetoRetenido")
        )
        row["Identificación Sujeto Retenido"] = _texto_emitidos_retencion_na(
            info_comp.findtext("identificacionSujetoRetenido")
        )
        row["Periodo Fiscal"] = _texto_emitidos_retencion_na(info_comp.findtext("periodoFiscal"))
        row["Dirección del Establecimiento"] = dir_est
        row["Tipo Identificación Sujeto Retenido"] = _texto_emitidos_retencion_na(
            info_comp.findtext("tipoIdentificacionSujetoRetenido")
        )
        row["Parte Relacionada"] = _texto_emitidos_retencion_na(info_comp.findtext("parteRel"))

    if doc_sustento is not None:
        row["Código de Sustento"] = _texto_emitidos_retencion_na(doc_sustento.findtext("codSustento"))
        row["Código del Documento de Sustento"] = _texto_emitidos_retencion_na(
            doc_sustento.findtext("codDocSustento")
        )
        row["Número de Documento de Sustento"] = _texto_emitidos_retencion_na(
            doc_sustento.findtext("numDocSustento")
        )
        row["Fecha de Emisión del Documento de Sustento"] = _texto_emitidos_retencion_na(
            doc_sustento.findtext("fechaEmisionDocSustento")
        )
        row["Fecha de Registro Contable"] = _texto_emitidos_retencion_na(
            doc_sustento.findtext("fechaRegistroContable")
        )
        row["Número de Autorización del Documento de Sustento"] = _texto_emitidos_retencion_na(
            doc_sustento.findtext("numAutDocSustento")
        )
        row["Pago Local o Externo"] = _texto_emitidos_retencion_na(doc_sustento.findtext("pagoLocExt"))
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(doc_sustento.findtext("totalSinImpuestos"))
        row["Importe Total"] = _numero_emitidos_retencion(doc_sustento.findtext("importeTotal"))

        pago = doc_sustento.find("./pagos/pago")
        if pago is not None:
            row["Forma Pago"] = _label_forma_pago_emitidos_retencion(pago.findtext("formaPago"))
            row["Total Pago"] = _numero_emitidos_retencion(
                pago.findtext("total"),
                row["Importe Total"],
            )
            row["Plazo Pago"] = _texto_emitidos_retencion_na(pago.findtext("plazo"))
            row["Unidad Tiempo Pago"] = _texto_emitidos_retencion_na(pago.findtext("unidadTiempo"))
        else:
            row["Total Pago"] = row["Importe Total"]

        impuesto_doc = doc_sustento.find("./impuestosDocSustento/impuestoDocSustento")
        if impuesto_doc is not None:
            row["Código Impuesto Doc. Sustento"] = _texto_emitidos_retencion_na(
                impuesto_doc.findtext("codImpuestoDocSustento")
            )
            row["Código Porcentaje"] = _texto_emitidos_retencion_na(
                impuesto_doc.findtext("codigoPorcentaje")
            )
            row["Base Imponible Impuesto"] = _numero_emitidos_retencion(impuesto_doc.findtext("baseImponible"))
            row["Tarifa"] = _numero_emitidos_retencion(impuesto_doc.findtext("tarifa"))
            row["Valor Impuesto"] = _numero_emitidos_retencion(impuesto_doc.findtext("valorImpuesto"))
            row["Sustento Imp. 1 - Código"] = row["Código Impuesto Doc. Sustento"]
            row["Sustento Imp. 1 - Cod. Porcentaje"] = row["Código Porcentaje"]
            row["Sustento Imp. 1 - Base Imponible"] = row["Base Imponible Impuesto"]
            row["Sustento Imp. 1 - Tarifa"] = row["Tarifa"]
            row["Sustento Imp. 1 - Valor"] = row["Valor Impuesto"]

        retenciones = doc_sustento.findall("./retenciones/retencion")
    else:
        row["Total Pago"] = row["Importe Total"]
        retenciones = root.findall(".//impuestos/impuesto")

    for idx, ret in enumerate(retenciones[:2], start=1):
        codigo = _texto_emitidos_retencion(ret.findtext("codigo"))
        row[f"Factura Ret. {idx} - Codigo"] = codigo or 0
        row[f"Factura Ret. {idx} - Cod. Porcentaje"] = "N/A" if codigo else 0
        row[f"Factura Ret. {idx} - Tarifa"] = 0
        row[f"Factura Ret. {idx} - Valor"] = 0

        if codigo == "1":
            row["RENTA - codigoRetencion"] = _texto_emitidos_retencion(ret.findtext("codigoRetencion")) or 0
            row["RENTA - baseImponible"] = _numero_emitidos_retencion(ret.findtext("baseImponible"))
            row["RENTA - porcentajeRetener"] = _numero_emitidos_retencion(ret.findtext("porcentajeRetener"))
            row["RENTA - valorRetenido"] = _numero_emitidos_retencion(ret.findtext("valorRetenido"))
        elif codigo == "2":
            row["IVA - codigoRetencion"] = _texto_emitidos_retencion(ret.findtext("codigoRetencion")) or 0
            row["IVA - baseImponible"] = _numero_emitidos_retencion(ret.findtext("baseImponible"))
            row["IVA - porcentajeRetener"] = _numero_emitidos_retencion(ret.findtext("porcentajeRetener"))
            row["IVA - valorRetenido"] = _numero_emitidos_retencion(ret.findtext("valorRetenido"))

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)
    return row


def _map_retencion_legada_a_emitidos_row(legacy: dict | None) -> dict:
    row = _emitidos_retencion_default_row()
    if not isinstance(legacy, dict):
        return row
    row["Razón Social Emisor"] = _texto_emitidos_retencion_na(legacy.get("razonSocialEmisor"))
    row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(legacy.get("direccionSucursalEmisor"))
    row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(legacy.get("obligadoContabilidad"))
    row["Nombre Comercial"] = _texto_emitidos_retencion_na(legacy.get("nombreComercial"))
    row["Código del Documento"] = "07 - COMPROBANTE DE RETENCIÓN"
    row["Establecimiento"] = _texto_emitidos_retencion(legacy.get("establecimiento"))
    row["Punto de Emisión"] = _texto_emitidos_retencion(legacy.get("puntoEmision"))
    row["Secuencial"] = _texto_emitidos_retencion(legacy.get("secuencial"))
    row["Dirección Matriz"] = _texto_emitidos_retencion(legacy.get("direccionMatrizEmisor"))
    row["RUC Emisor"] = _texto_emitidos_retencion(legacy.get("rucEmisor"))
    row["Clave de Acceso"] = _texto_emitidos_retencion(legacy.get("claveAcceso"))
    row["Fecha de Emisión"] = _texto_emitidos_retencion(legacy.get("fechaEmision"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(legacy.get("fechaAutorizacion"))
    row["Importe Total"] = _numero_emitidos_retencion(legacy.get("valorTotal"))
    row["Total Sin Impuestos"] = _numero_emitidos_retencion(legacy.get("Base_Imponible_Ret_IR")) or _numero_emitidos_retencion(legacy.get("Base_Imponible_Ret_IVA"))
    row["Total Pago"] = row["Importe Total"]
    row["Agente de Retención"] = _texto_emitidos_retencion_na(legacy.get("numeroAgenteRetencion"))
    row["Contribuyente Especial"] = _texto_emitidos_retencion_na(legacy.get("numeroContribuyenteEspecial"))
    row["Obligado a llevar Contabilidad"] = row["Obligado Contabilidad"]
    row["Razón Social Sujeto Retenido"] = _texto_emitidos_retencion_na(legacy.get("razonSocialSujetoRetenido"))
    row["Identificación Sujeto Retenido"] = _texto_emitidos_retencion_na(legacy.get("identificacionSujetoRetenido"))
    row["Periodo Fiscal"] = _texto_emitidos_retencion_na(legacy.get("Ejercicio_Fiscal"))
    row["Dirección del Establecimiento"] = row["Dir. Establecimiento"]
    row["Código del Documento de Sustento"] = _texto_emitidos_retencion_na(legacy.get("Comprobante_Sustento"))
    row["Número de Documento de Sustento"] = _texto_emitidos_retencion_na(legacy.get("Numero_Sustento"))
    row["Fecha de Emisión del Documento de Sustento"] = _texto_emitidos_retencion_na(legacy.get("Fecha_Emision_Sustento"))
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(legacy.get("informacionAdicional"))
    row["RENTA - codigoRetencion"] = _texto_emitidos_retencion(legacy.get("Impuesto_Ret_IR")) or 0
    row["RENTA - baseImponible"] = _numero_emitidos_retencion(legacy.get("Base_Imponible_Ret_IR"))
    row["RENTA - porcentajeRetener"] = _numero_emitidos_retencion(legacy.get("Porcentaje_Ret_IR"))
    row["RENTA - valorRetenido"] = _numero_emitidos_retencion(legacy.get("Valor_Retenido_IR"))
    row["IVA - codigoRetencion"] = _texto_emitidos_retencion(legacy.get("Impuesto_Ret_IVA")) or 0
    row["IVA - baseImponible"] = _numero_emitidos_retencion(legacy.get("Base_Imponible_Ret_IVA"))
    row["IVA - porcentajeRetener"] = _numero_emitidos_retencion(legacy.get("Porcentaje_Ret_IVA"))
    row["IVA - valorRetenido"] = _numero_emitidos_retencion(legacy.get("Valor_Retenido_IVA"))
    if row["RENTA - codigoRetencion"]:
        row["Factura Ret. 1 - Codigo"] = "1"
        row["Factura Ret. 1 - Cod. Porcentaje"] = "N/A"
    if row["IVA - codigoRetencion"]:
        target = "1" if row["Factura Ret. 1 - Codigo"] == 0 else "2"
        row[f"Factura Ret. {target} - Codigo"] = "2"
        row[f"Factura Ret. {target} - Cod. Porcentaje"] = "N/A"
    return row


def _extraer_lineas_layout_pdf(pdf_path: Path, y_tolerance: float = 3.0) -> list[dict]:
    if pdfplumber is None:
        return []
    try:
        resultado = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
                words = sorted(words, key=lambda item: (float(item.get("top", 0)), float(item.get("x0", 0))))
                current_words = []
                current_top = None
                for word in words:
                    word_top = float(word.get("top", 0.0))
                    if current_top is None or abs(word_top - current_top) <= y_tolerance:
                        current_words.append(word)
                        current_top = word_top if current_top is None else current_top
                        continue
                    current_sorted = sorted(current_words, key=lambda item: float(item.get("x0", 0)))
                    resultado.append({
                        "page": page_number,
                        "top": current_top,
                        "text": " ".join((w.get("text", "") or "").strip() for w in current_sorted if (w.get("text", "") or "").strip()),
                        "words": current_sorted,
                    })
                    current_words = [word]
                    current_top = word_top
                if current_words:
                    current_sorted = sorted(current_words, key=lambda item: float(item.get("x0", 0)))
                    resultado.append({
                        "page": page_number,
                        "top": current_top,
                        "text": " ".join((w.get("text", "") or "").strip() for w in current_sorted if (w.get("text", "") or "").strip()),
                        "words": current_sorted,
                    })
        return resultado
    except Exception:
        return []


def _texto_linea_layout(linea: dict, min_x: float | None = None, max_x: float | None = None) -> str:
    words = []
    for word in linea.get("words", []):
        x0 = float(word.get("x0", 0.0))
        if min_x is not None and x0 < min_x:
            continue
        if max_x is not None and x0 > max_x:
            continue
        text = (word.get("text", "") or "").strip()
        if text:
            words.append(text)
    return " ".join(words).strip()


def _buscar_indice_linea_layout(lineas: list[dict], token: str, start: int = 0) -> int | None:
    token_norm = _normalizar_label_simple(token)
    for idx in range(start, len(lineas)):
        texto = _normalizar_label_simple(lineas[idx].get("text", ""))
        if token_norm in texto:
            return idx
    return None


def _buscar_indice_linea_layout_exacta(lineas: list[dict], token: str, start: int = 0) -> int | None:
    token_norm = _normalizar_label_simple(token)
    for idx in range(start, len(lineas)):
        texto = _normalizar_label_simple(lineas[idx].get("text", ""))
        if texto == token_norm:
            return idx
    return None


def _siguiente_linea_layout_no_vacia(lineas: list[dict], idx: int, min_x: float | None = None, max_x: float | None = None) -> str:
    for pos in range(idx + 1, len(lineas)):
        texto = _texto_linea_layout(lineas[pos], min_x=min_x, max_x=max_x)
        if texto:
            return texto
    return ""


def _fecha_hora_pdf_a_iso(valor: str) -> str:
    valor = (valor or "").strip()
    if not valor:
        return ""
    try:
        dt = datetime.strptime(valor, "%d/%m/%Y %H:%M:%S")
        return dt.strftime("%Y-%m-%dT%H:%M:%S-05:00")
    except Exception:
        return valor


def _codigo_tipo_identificacion_desde_numero(identificacion: str, default: str = "No Disponible") -> str:
    digits = re.sub(r"\D+", "", identificacion or "")
    if len(digits) == 13:
        return _label_tipo_ident_emitidos_nota_credito("04")
    if len(digits) == 10:
        return _label_tipo_ident_emitidos_nota_credito("05")
    return default


def _codigo_documento_sri(descripcion: str, default: str = "No Disponible") -> str:
    texto_norm = _normalizar_label_simple(descripcion)
    if "FACTURA" in texto_norm:
        return "01"
    if "LIQUIDACION" in texto_norm:
        return "03"
    if "CREDITO" in texto_norm:
        return "04"
    if "DEBITO" in texto_norm:
        return "05"
    if "GUIA" in texto_norm:
        return "06"
    if "RETENCION" in texto_norm:
        return "07"
    return default


def _combinar_rows_emitidos_especificos(primary: dict, secondary: dict | None) -> dict:
    if not isinstance(secondary, dict):
        return primary
    result = dict(primary)
    for key, value in secondary.items():
        current = result.get(key)
        if current in ("", None) and value not in ("", None):
            result[key] = value
    return result


def _extraer_campos_adicionales_por_layout(lineas: list[dict], top_min: float, left_max: float = 320.0) -> str:
    adicionales = []
    for linea in lineas:
        if float(linea.get("top", 0.0)) <= top_min:
            continue
        texto = _texto_linea_layout(linea, max_x=left_max)
        if not texto or ":" not in texto:
            continue
        etiqueta, valor = texto.split(":", 1)
        etiqueta = etiqueta.strip()
        valor = valor.strip()
        if not etiqueta or not valor:
            continue
        if _normalizar_label_simple(etiqueta) == "INFORMACION ADICIONAL":
            continue
        adicionales.append(f"{etiqueta}: {valor}")
    return "; ".join(adicionales)


def _extraer_bloque_direccion_layout(
    lineas: list[dict],
    idx_inicio: int | None,
    *,
    max_x: float = 320.0,
    stop_tokens: tuple[str, ...] = ("SUCURSAL", "CLAVE DE ACCESO", "CONTRIBUYENTE", "OBLIGADO"),
) -> str:
    if idx_inicio is None:
        return ""
    partes = []
    for pos in range(idx_inicio + 1, min(idx_inicio + 6, len(lineas))):
        texto = _texto_linea_layout(lineas[pos], max_x=max_x)
        if not texto:
            continue
        normalizado = _normalizar_label_simple(texto)
        if normalizado in {"MATRIZ", "DIRECCION", "DIRECCION MATRIZ"}:
            continue
        if any(token in normalizado for token in stop_tokens):
            break
        if "EMISION" in normalizado:
            texto = re.split(r"(?i)\bEMISION\b", texto)[0].strip(" :-")
        if texto:
            partes.append(texto)
    return " ".join(partes).strip()


def _extraer_items_emitidos_layout(
    lineas: list[dict],
    *,
    top_inicio: float,
    top_fin: float,
) -> list[dict]:
    region = [ln for ln in lineas if top_inicio < float(ln.get("top", 0.0)) < top_fin]
    items = []
    idx = 0
    while idx < len(region):
        texto = region[idx].get("text", "") or ""
        texto_norm = _normalizar_label_simple(texto)
        if "SUBTOTAL" in texto_norm:
            break
        if any(token in texto_norm for token in ("PRINCIPAL AUXILIAR", "CANTIDAD DESCRIPCION", "COD COD", "PRECIO UNITARIO")):
            idx += 1
            continue
        match = re.match(r"^([A-Z0-9.\-]+)\s+([A-Z0-9.\-]+)\s+(.+)$", texto.strip(), flags=re.IGNORECASE)
        if not match:
            idx += 1
            continue
        codigo_pref, codigo_aux_pref, descripcion = match.groups()
        if not re.search(r"\d", codigo_pref):
            idx += 1
            continue
        cantidad = ""
        precio_unitario = ""
        if idx + 1 < len(region):
            numeros = re.findall(r"\d+(?:\.\d+)?", region[idx + 1].get("text", ""))
            if numeros:
                cantidad = numeros[0]
            if len(numeros) > 1:
                precio_unitario = numeros[1]
        codigo = codigo_pref
        descripcion_partes = [descripcion.strip()]
        if idx + 2 < len(region):
            cont = region[idx + 2].get("text", "") or ""
            cont_match = re.match(r"^([A-Z0-9.\-]+)\s+([A-Z0-9.\-]+)\s+(.+)$", cont.strip(), flags=re.IGNORECASE)
            if cont_match:
                suf1, _suf2, desc_tail = cont_match.groups()
                if codigo_pref.endswith("-") and re.fullmatch(r"[A-Z0-9]+", suf1):
                    codigo = f"{codigo_pref}{suf1}"
                if desc_tail:
                    descripcion_partes.append(desc_tail.strip())
                idx += 3
            else:
                idx += 2
        else:
            idx += 1
        items.append(
            {
                "codigo": codigo.strip(),
                "auxiliar": codigo_aux_pref.strip(),
                "descripcion": " ".join(fragment for fragment in descripcion_partes if fragment).strip(),
                "cantidad": cantidad.strip(),
                "precio_unitario": precio_unitario.strip(),
            }
        )
    return items


def _formatear_descripciones_emitidos(items: list[dict], *, incluir_auxiliar: bool = False) -> str:
    partes = []
    for item in items:
        codigo = (item.get("codigo") or "").strip()
        descripcion = (item.get("descripcion") or "").strip()
        cantidad = _formatear_cantidad_emitidos(item.get("cantidad") or "")
        precio_unitario = _formatear_precio_emitidos(item.get("precio_unitario") or "")
        fragmentos = [f"Código: {codigo}"]
        if incluir_auxiliar:
            fragmentos.append(f"Aux: {(item.get('auxiliar') or '').strip()}")
        fragmentos.append(f"Desc: {descripcion}")
        if cantidad:
            fragmentos.append(f"Cant: {cantidad}")
        if precio_unitario:
            fragmentos.append(f"P.Unit: {precio_unitario}")
        partes.append(", ".join(fragmentos))
    return " ; ".join(partes)


def _formatear_cantidad_emitidos(cantidad: str) -> str:
    try:
        valor = float(str(cantidad).replace(",", "."))
    except Exception:
        return (cantidad or "").strip()
    if valor.is_integer():
        return f"{valor:.4f}"
    return (cantidad or "").strip()


def _formatear_precio_emitidos(precio: str) -> str:
    try:
        valor = float(str(precio).replace(",", "."))
    except Exception:
        return (precio or "").strip()
    return f"{valor:.5f}"


def _extraer_campos_adicionales_emitidos_desde_texto(texto_pdf: str) -> str:
    lineas = [ln.strip() for ln in (texto_pdf or "").splitlines() if ln.strip()]
    inicio = None
    for idx, linea in enumerate(lineas):
        if "INFORMACION ADICIONAL" in _normalizar_label_simple(unicodedata.normalize("NFKD", linea).encode("ascii", "ignore").decode("ascii")):
            inicio = idx + 1
            break
    if inicio is None:
        return ""
    adicionales = []
    stop_tokens = (
        "SUBTOTAL",
        "TOTAL DESCUENTO",
        "ICE",
        "IRBPNR",
        "PROPINA",
        "VALOR TOTAL",
        "FORMA DE PAGO VALOR",
        "IVA 15",
        "IVA 12",
        "IVA 8",
        "IVA 5",
    )
    for linea in lineas[inicio:]:
        texto = linea
        texto_norm = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
        normalized = _normalizar_label_simple(texto_norm)
        for token in stop_tokens:
            pos = normalized.find(token)
            if pos != -1:
                texto = texto[:pos].rstrip(" :-")
                break
        if not texto:
            if any(token in normalized for token in stop_tokens):
                break
            continue
        if ":" not in texto:
            if any(token in normalized for token in stop_tokens):
                break
            if adicionales:
                adicionales[-1] = f"{adicionales[-1]} {texto}".strip()
            continue
        etiqueta_norm = _normalizar_label_simple(texto.split(":", 1)[0])
        if etiqueta_norm == "INFORMACION ADICIONAL":
            continue
        adicionales.append(texto)
    return "; ".join(adicionales)


def _extraer_datos_pdf_retencion_emitido(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_layout_pdf(pdf_path)
    if not lineas:
        return _map_retencion_legada_a_emitidos_row(_extraer_datos_pdf_retencion(pdf_path))

    row = _emitidos_retencion_default_row()
    row["Estado"] = "AUTORIZADO"
    row["Código del Documento"] = "07 - COMPROBANTE DE RETENCIÓN"
    row["Forma Pago"] = "20 - OTROS CON UTILIZACIÓN DEL SISTEMA FINANCIERO"
    row["Pago Local o Externo"] = "01"
    row["Código de Sustento"] = "06"
    row["Moneda"] = "No Disponible"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Tipo Identificación Comprador"] = "No Disponible - No Disponible"
    row["Identificación Comprador"] = "No Disponible"
    row["Razón Social Comprador"] = "No Disponible"
    row["Dirección Comprador"] = "No Disponible"
    row["Contribuyente RIMPE"] = "No Disponible"
    row["Contribuyente Especial"] = "No Disponible"

    texto = _leer_texto_pdf(pdf_path)
    texto_norm = _normalizar_texto_pdf(texto)
    auth = _extraer_regex(texto_norm, [r"(\d{49})"])
    numero = _extraer_regex(texto_norm, [r"No\.\s*(\d{3}-\d{3}-\d{9})"])
    ruc = _extraer_regex(texto_norm, [r"R\.U\.C\.\s*:?\s*(\d{13})"])
    fecha_hora_auth = _extraer_regex(texto_norm, [r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"])

    row["Número de Autorización"] = auth or "No Disponible"
    row["Clave de Acceso"] = auth or "No Disponible"
    row["Fecha de Autorización"] = _fecha_hora_pdf_a_iso(fecha_hora_auth)
    row["RUC Emisor"] = ruc or "No Disponible"
    if numero:
        row["Establecimiento"], row["Punto de Emisión"], row["Secuencial"] = numero.split("-")
    if auth and len(auth) == 49:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(auth[23])
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(auth[47])
    else:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(_extraer_regex(texto_norm, [r"AMBIENTE\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"]))
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(_extraer_regex(texto_norm, [r"EMISI[ÓO]N\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"]))

    idx_num_aut = _buscar_indice_linea_layout(lineas, "NUMERO DE AUTORIZACION")
    if idx_num_aut is not None:
        row["Razón Social Emisor"] = _texto_emitidos_retencion_na(
            _siguiente_linea_layout_no_vacia(lineas, idx_num_aut + 1, max_x=260)
        )
    idx_aut = _buscar_indice_linea_layout_exacta(lineas, "AUTORIZACION")
    if idx_aut is not None:
        for pos in range(idx_aut + 1, len(lineas)):
            candidato = _texto_linea_layout(lineas[pos], max_x=260)
            if candidato and candidato != row["Razón Social Emisor"]:
                row["Nombre Comercial"] = _texto_emitidos_retencion_na(candidato)
                break
    idx_amb = _buscar_indice_linea_layout(lineas, "AMBIENTE")
    if idx_amb is not None:
        candidato = re.sub(r"(?i)\bAMBIENTE\s*:.*$", "", lineas[idx_amb].get("text", "")).strip()
        if not candidato and idx_amb > 0:
            candidato = _texto_linea_layout(lineas[idx_amb - 1], max_x=260)
        if candidato and candidato != row["Razón Social Emisor"]:
            row["Nombre Comercial"] = _texto_emitidos_retencion_na(candidato)

    idx_matriz = _buscar_indice_linea_layout(lineas, "DIRECCION MATRIZ")
    if idx_matriz is not None:
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(
            _texto_linea_layout(lineas[idx_matriz], min_x=120, max_x=310)
        )
    idx_sucursal = _buscar_indice_linea_layout(lineas, "DIRECCION SUCURSAL")
    if idx_sucursal is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
            _texto_linea_layout(lineas[idx_sucursal], min_x=120, max_x=310)
        )
        row["Dirección del Establecimiento"] = row["Dir. Establecimiento"]

    row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*(SI|NO)"]))
    row["Obligado a llevar Contabilidad"] = row["Obligado Contabilidad"]
    row["Agente de Retención"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"AGENTE\s+DE\s+RETENCION\s+RESOLUCION\s+NO\.\s*([0-9]+)"]))

    idx_razon = _buscar_indice_linea_layout(lineas, "RAZON SOCIAL / NOMBRES Y APELLIDOS")
    if idx_razon is not None:
        row["Razón Social Sujeto Retenido"] = _texto_emitidos_retencion_na(_texto_linea_layout(lineas[idx_razon], min_x=180))
        row["Identificación Sujeto Retenido"] = _texto_emitidos_retencion_na(_siguiente_linea_layout_no_vacia(lineas, idx_razon, min_x=80, max_x=180))
        row["Fecha de Emisión"] = _texto_emitidos_retencion_na(_siguiente_linea_layout_no_vacia(lineas, idx_razon + 1, min_x=80, max_x=180))

    ident_sujeto = re.sub(r"\D+", "", row["Identificación Sujeto Retenido"])
    tipo_id_sujeto = re.sub(r"\D+", "", row["Identificación Sujeto Retenido"])
    if len(tipo_id_sujeto) == 13:
        row["Tipo Identificación Sujeto Retenido"] = "04"
    elif len(tipo_id_sujeto) == 10:
        row["Tipo Identificación Sujeto Retenido"] = "05"

    idx_info_adic = _buscar_indice_linea_layout(lineas, "INFORMACION ADICIONAL")
    if idx_info_adic is not None:
        row["Campos Adicionales"] = _texto_emitidos_retencion_na(_extraer_campos_adicionales_por_layout(lineas, float(lineas[idx_info_adic].get("top", 0.0))))

    top_inicio_tabla = 0.0
    idx_inicio_tabla = _buscar_indice_linea_layout(lineas, "COMPROBANTE NUMERO")
    if idx_inicio_tabla is not None:
        top_inicio_tabla = float(lineas[idx_inicio_tabla].get("top", 0.0))
    top_fin_tabla = float(lineas[idx_info_adic].get("top", 9999.0)) if idx_info_adic is not None else 9999.0
    lineas_tabla = [ln for ln in lineas if top_inicio_tabla < float(ln.get("top", 0.0)) < top_fin_tabla]

    tipo_doc_sustento = ""
    numero_doc_partes = []
    fecha_doc_sustento = ""
    periodo_fiscal = ""
    iva_row = None
    renta_row = None
    renta_top = None

    for linea in lineas_tabla:
        texto_linea = linea.get("text", "")
        texto_upper = _normalizar_label_simple(texto_linea)
        izquierda = _texto_linea_layout(linea, max_x=75)
        centro_num = [w.get("text", "") for w in linea.get("words", []) if 80 <= float(w.get("x0", 0.0)) <= 150 and re.fullmatch(r"\d+", (w.get("text", "") or ""))]
        if izquierda and not tipo_doc_sustento and any(token in texto_upper for token in ("FACTURA", "LIQUIDACION", "NOTA", "GUIA")):
            tipo_doc_sustento = izquierda
        if centro_num:
            numero_doc_partes.extend(centro_num)
        if not fecha_doc_sustento:
            fecha_match = re.search(r"\d{2}/\d{2}/\d{4}", texto_linea)
            if fecha_match:
                fecha_doc_sustento = fecha_match.group(0)
        if not periodo_fiscal:
            periodo_match = re.search(r"\d{2}/\d{4}", texto_linea)
            if periodo_match:
                periodo_fiscal = periodo_match.group(0)
        if "IVA" in texto_upper:
            nums = [w.get("text", "") for w in linea.get("words", []) if float(w.get("x0", 0.0)) >= 280 and re.fullmatch(r"\d+(?:\.\d+)?", (w.get("text", "") or ""))]
            if len(nums) >= 3:
                iva_row = tuple(nums[:3]) + (float(linea.get("top", 0.0)),)
        if "RENTA" in texto_upper or "IMPUESTO A LA" in texto_upper:
            renta_top = float(linea.get("top", 0.0))

    base_candidates = []
    for linea in lineas_tabla:
        texto_upper = _normalizar_label_simple(linea.get("text", ""))
        if "IVA" in texto_upper:
            continue
        nums = [w.get("text", "") for w in linea.get("words", []) if float(w.get("x0", 0.0)) >= 280 and re.fullmatch(r"\d+(?:\.\d+)?", (w.get("text", "") or ""))]
        if len(nums) >= 3:
            base_candidates.append((float(linea.get("top", 0.0)), nums[:3]))
    if base_candidates:
        if renta_top is not None:
            base_candidates.sort(key=lambda item: abs(item[0] - renta_top))
        renta_row = base_candidates[0]

    numero_doc_sustento = "".join(numero_doc_partes)
    row["Código del Documento de Sustento"] = _codigo_documento_sri(tipo_doc_sustento)
    row["Número de Documento de Sustento"] = _texto_emitidos_retencion_na(numero_doc_sustento)
    row["Fecha de Emisión del Documento de Sustento"] = _texto_emitidos_retencion_na(fecha_doc_sustento)
    row["Fecha de Registro Contable"] = row["Fecha de Emisión"]
    row["Periodo Fiscal"] = _texto_emitidos_retencion_na(periodo_fiscal)
    row["Parte Relacionada"] = "No Disponible"
    row["Número de Autorización del Documento de Sustento"] = "No Disponible"

    renta_base = 0
    if renta_row:
        renta_base = _numero_emitidos_retencion(renta_row[1][0])
        row["RENTA - baseImponible"] = renta_base
        row["RENTA - porcentajeRetener"] = _numero_emitidos_retencion(renta_row[1][1])
        row["RENTA - valorRetenido"] = _numero_emitidos_retencion(renta_row[1][2])
        row["RENTA - codigoRetencion"] = 0

    iva_top = None
    if iva_row:
        row["IVA - baseImponible"] = _numero_emitidos_retencion(iva_row[0])
        row["IVA - porcentajeRetener"] = _numero_emitidos_retencion(iva_row[1])
        row["IVA - valorRetenido"] = _numero_emitidos_retencion(iva_row[2])
        row["IVA - codigoRetencion"] = 1
        iva_top = iva_row[3]

    total_sin_impuestos = renta_base or 0
    row["Total Sin Impuestos"] = total_sin_impuestos
    row["Base Gravada"] = 0
    row["Base No Gravada"] = 0
    row["Código Impuesto Doc. Sustento"] = "2" if total_sin_impuestos else "No Disponible"
    row["Código Porcentaje"] = "4" if total_sin_impuestos else "No Disponible"
    row["Base Imponible Impuesto"] = total_sin_impuestos if total_sin_impuestos else 0
    row["Tarifa"] = 15 if total_sin_impuestos else 0
    row["Valor Impuesto"] = round(total_sin_impuestos * 0.15, 2) if total_sin_impuestos else 0
    row["Importe Total"] = round(total_sin_impuestos + row["Valor Impuesto"], 2) if total_sin_impuestos else 0
    row["Total Pago"] = row["Importe Total"]
    row["Sustento Imp. 1 - Código"] = row["Código Impuesto Doc. Sustento"]
    row["Sustento Imp. 1 - Cod. Porcentaje"] = row["Código Porcentaje"]
    row["Sustento Imp. 1 - Base Imponible"] = row["Base Imponible Impuesto"]
    row["Sustento Imp. 1 - Tarifa"] = row["Tarifa"]
    row["Sustento Imp. 1 - Valor"] = row["Valor Impuesto"]

    conceptos = []
    if iva_top is not None:
        conceptos.append((iva_top, "2"))
    if renta_row:
        conceptos.append((renta_row[0], "1"))
    conceptos.sort(key=lambda item: item[0])
    for idx, (_, codigo) in enumerate(conceptos[:2], start=1):
        row[f"Factura Ret. {idx} - Codigo"] = codigo
        row[f"Factura Ret. {idx} - Cod. Porcentaje"] = "N/A"

    legacy = _map_retencion_legada_a_emitidos_row(_extraer_datos_pdf_retencion(pdf_path))
    return _combinar_rows_emitidos_especificos(row, legacy)


def _map_nota_credito_legada_a_emitidos_row(legacy: dict | None) -> dict:
    row = _nota_credito_emitidos_default_row()
    if not isinstance(legacy, dict):
        return row
    row["Ambiente"] = _label_ambiente_emitidos_retencion(legacy.get("ambiente"))
    row["Razón Social Emisor"] = _texto_emitidos_retencion_na(legacy.get("razonSocialEmisor"))
    row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(legacy.get("direccionSucursalEmisor"))
    row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(legacy.get("obligadoContabilidad"))
    row["Tipo Identificación Comprador"] = "No Disponible"
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(legacy.get("identificacionComprador"))
    row["Tipo Emisión"] = _label_emision_emitidos_retencion(legacy.get("emision"))
    row["Nombre Comercial"] = _texto_emitidos_retencion_na(legacy.get("nombreComercial"))
    row["Código del Documento"] = "04 - NOTA DE CRÉDITO"
    row["Establecimiento"] = _texto_emitidos_retencion(legacy.get("establecimiento"))
    row["Punto de Emisión"] = _texto_emitidos_retencion(legacy.get("puntoEmision"))
    row["Secuencial"] = _texto_emitidos_retencion(legacy.get("secuencial"))
    row["Dirección Matriz"] = _texto_emitidos_retencion(legacy.get("direccionMatrizEmisor"))
    row["Contribuyente RIMPE"] = _texto_emitidos_retencion_na(legacy.get("tipoContribuyenteRIMPE"))
    row["RUC Emisor"] = _texto_emitidos_retencion(legacy.get("rucEmisor"))
    row["Clave de Acceso"] = _texto_emitidos_retencion(legacy.get("claveAcceso"))
    row["Fecha de Emisión"] = _texto_emitidos_retencion(legacy.get("fechaEmision"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(legacy.get("fechaAutorizacion"))
    row["Razón Social Comprador"] = _texto_emitidos_retencion_na(legacy.get("razonSocialComprador"))
    row["Dirección Comprador"] = _texto_emitidos_retencion_na(legacy.get("direccionComprador"))
    row["Moneda"] = "No Disponible"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Descripciones"] = _texto_emitidos_retencion(legacy.get("descripcionesProductos"))
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Total Sin Impuestos"] = _numero_emitidos_retencion(legacy.get("subtotalSinImpuestos"))
    row["Base Gravada"] = _numero_emitidos_retencion(legacy.get("subtotal15")) or _numero_emitidos_retencion(legacy.get("subtotal12")) or _numero_emitidos_retencion(legacy.get("subtotal8")) or _numero_emitidos_retencion(legacy.get("subtotal5"))
    row["Base No Gravada"] = _numero_emitidos_retencion(legacy.get("subtotal0"))
    if _numero_emitidos_retencion(legacy.get("iva15")):
        row["Tarifas IVA"] = "15%"
    elif _numero_emitidos_retencion(legacy.get("iva12")):
        row["Tarifas IVA"] = "12%"
    elif _numero_emitidos_retencion(legacy.get("iva8")):
        row["Tarifas IVA"] = "8%"
    elif _numero_emitidos_retencion(legacy.get("iva5")):
        row["Tarifas IVA"] = "5%"
    row["Monto IVA"] = _numero_emitidos_retencion(legacy.get("iva15")) or _numero_emitidos_retencion(legacy.get("iva12")) or _numero_emitidos_retencion(legacy.get("iva8")) or _numero_emitidos_retencion(legacy.get("iva5"))
    row["Total Descuento"] = _numero_emitidos_retencion(legacy.get("totalDescuento"))
    row["Propina"] = _numero_emitidos_retencion(legacy.get("propina"))
    row["Importe Total"] = _numero_emitidos_retencion(legacy.get("valorTotal"))
    row["Total Pago"] = 0
    row["Código Documento Modificado"] = _texto_emitidos_retencion_na(legacy.get("comprobanteModificado"))
    row["Número Documento Modificado"] = _texto_emitidos_retencion_na(legacy.get("comprobanteModificado"))
    row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(legacy.get("fechaEmisionModificado"))
    row["Motivo"] = _texto_emitidos_retencion_na(legacy.get("razonModificacion"))
    row["Valor Modificación"] = _numero_emitidos_retencion(legacy.get("valorModificacion"))
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(legacy.get("informacionAdicional"))
    row["Base Gravada 15%"] = _numero_emitidos_retencion(legacy.get("subtotal15"))
    row["Monto IVA 15%"] = _numero_emitidos_retencion(legacy.get("iva15"))
    return row


def _map_nota_debito_legada_a_emitidos_row(legacy: dict | None) -> dict:
    row = _nota_debito_emitidos_default_row()
    if not isinstance(legacy, dict):
        return row
    row["Ambiente"] = _label_ambiente_emitidos_retencion(legacy.get("ambiente"))
    row["Razón Social Emisor"] = _texto_emitidos_retencion_na(legacy.get("razonSocialEmisor"))
    row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
        legacy.get("direccionSucursalEmisor") or legacy.get("direccionMatrizEmisor")
    )
    obligado = _texto_emitidos_retencion(legacy.get("obligadoContabilidad"))
    row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
    row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
        _codigo_tipo_identificacion_desde_numero(legacy.get("identificacionComprador"))
    )
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(legacy.get("identificacionComprador"))
    row["Tipo Emisión"] = _label_emision_emitidos_retencion(legacy.get("emision"))
    row["Nombre Comercial"] = _texto_emitidos_retencion_na(
        legacy.get("nombreComercial") or legacy.get("razonSocialEmisor")
    )
    row["Código del Documento"] = "05 - NOTA DE DÉBITO"
    row["Establecimiento"] = _texto_emitidos_retencion(legacy.get("establecimiento"))
    row["Punto de Emisión"] = _texto_emitidos_retencion(legacy.get("puntoEmision"))
    row["Secuencial"] = _texto_emitidos_retencion(legacy.get("secuencial"))
    row["Dirección Matriz"] = _texto_emitidos_retencion_na(legacy.get("direccionMatrizEmisor"))
    row["Contribuyente RIMPE"] = _texto_emitidos_retencion_na(legacy.get("tipoContribuyenteRIMPE"))
    row["RUC Emisor"] = _texto_emitidos_retencion(legacy.get("rucEmisor"))
    row["Clave de Acceso"] = _texto_emitidos_retencion(legacy.get("claveAcceso"))
    row["Fecha de Emisión"] = _texto_emitidos_retencion(legacy.get("fechaEmision"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(legacy.get("fechaAutorizacion"))
    row["Razón Social Comprador"] = _texto_emitidos_retencion_na(legacy.get("razonSocialComprador"))
    row["Dirección Comprador"] = _texto_emitidos_retencion_na(legacy.get("direccionComprador"))
    row["Moneda"] = "No Disponible"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Descripciones"] = _texto_emitidos_retencion(legacy.get("descripcionesProductos"))
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Total Sin Impuestos"] = _numero_emitidos_retencion(legacy.get("subtotalSinImpuestos"))
    row["Base Gravada"] = (
        _numero_emitidos_retencion(legacy.get("subtotal15"))
        or _numero_emitidos_retencion(legacy.get("subtotal12"))
        or _numero_emitidos_retencion(legacy.get("subtotal8"))
        or _numero_emitidos_retencion(legacy.get("subtotal5"))
    )
    row["Base No Gravada"] = _numero_emitidos_retencion(legacy.get("subtotal0"))
    if _numero_emitidos_retencion(legacy.get("iva15")):
        row["Tarifas IVA"] = "15%"
    elif _numero_emitidos_retencion(legacy.get("iva12")):
        row["Tarifas IVA"] = "12%"
    elif _numero_emitidos_retencion(legacy.get("iva8")):
        row["Tarifas IVA"] = "8%"
    elif _numero_emitidos_retencion(legacy.get("iva5")):
        row["Tarifas IVA"] = "5%"
    row["Monto IVA"] = (
        _numero_emitidos_retencion(legacy.get("iva15"))
        or _numero_emitidos_retencion(legacy.get("iva12"))
        or _numero_emitidos_retencion(legacy.get("iva8"))
        or _numero_emitidos_retencion(legacy.get("iva5"))
    )
    row["Total Descuento"] = _numero_emitidos_retencion(legacy.get("totalDescuento"))
    row["Propina"] = _numero_emitidos_retencion(legacy.get("propina"))
    row["Importe Total"] = _numero_emitidos_retencion(legacy.get("valorTotal"))
    row["Total Pago"] = row["Importe Total"]
    row["Código Documento Modificado"] = _codigo_documento_sri(legacy.get("comprobanteModificado"))
    row["Número Documento Modificado"] = _texto_emitidos_retencion_na(
        _extraer_regex(_texto_emitidos_retencion(legacy.get("comprobanteModificado")), [r"(\d{3}-\d{3}-\d{9})"])
        or legacy.get("comprobanteModificado")
    )
    row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(legacy.get("fechaEmisionModificado"))
    row["Motivo"] = _texto_emitidos_retencion_na(legacy.get("razonModificacion"))
    row["Valor Modificación"] = _numero_emitidos_retencion(legacy.get("valorModificacion")) or row["Importe Total"]
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(legacy.get("informacionAdicional"))
    row["Base Gravada 15%"] = _numero_emitidos_retencion(legacy.get("subtotal15"))
    row["Monto IVA 15%"] = _numero_emitidos_retencion(legacy.get("iva15"))
    return row


def _extraer_datos_pdf_nota_credito_emitido(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_layout_pdf(pdf_path)
    if not lineas:
        return _map_nota_credito_legada_a_emitidos_row(_extraer_datos_pdf_nota_credito(pdf_path))

    row = _nota_credito_emitidos_default_row()
    row["Estado"] = "AUTORIZADO"
    row["Código del Documento"] = "04 - NOTA DE CRÉDITO"
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Moneda"] = "DOLAR"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Contribuyente RIMPE"] = "No Disponible"

    texto = _leer_texto_pdf(pdf_path)
    texto_norm = _normalizar_texto_pdf(texto)
    auth = _extraer_regex(texto_norm, [r"(\d{49})"])
    numero = _extraer_regex(texto_norm, [r"No\.\s*(\d{3}-\d{3}-\d{9})"])
    ruc = _extraer_regex(texto_norm, [r"R\.U\.C\.\s*:?\s*(\d{13})"])
    fecha_hora_auth = _extraer_regex(texto_norm, [r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"])
    row["Número de Autorización"] = auth or "No Disponible"
    row["Clave de Acceso"] = auth or "No Disponible"
    row["Fecha de Autorización"] = _fecha_hora_pdf_a_iso(fecha_hora_auth)
    row["RUC Emisor"] = ruc or "No Disponible"
    if numero:
        row["Establecimiento"], row["Punto de Emisión"], row["Secuencial"] = numero.split("-")
    if auth and len(auth) == 49:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(auth[23])
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(auth[47])
    else:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(_extraer_regex(texto_norm, [r"AMBIENTE\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"]))
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(_extraer_regex(texto_norm, [r"EMISI[ÓO]N\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"]))

    idx_num_aut = _buscar_indice_linea_layout(lineas, "NUMERO DE AUTORIZACION")
    if idx_num_aut is not None:
        row["Razón Social Emisor"] = _texto_emitidos_retencion_na(
            _siguiente_linea_layout_no_vacia(lineas, idx_num_aut + 1, max_x=260)
        )
    idx_aut = _buscar_indice_linea_layout_exacta(lineas, "AUTORIZACION")
    if idx_aut is not None:
        for pos in range(idx_aut + 1, len(lineas)):
            candidato = _texto_linea_layout(lineas[pos], max_x=260)
            if candidato and candidato != row["Razón Social Emisor"]:
                row["Nombre Comercial"] = _texto_emitidos_retencion_na(candidato)
                break

    match_matriz = re.search(r"Direccion\s+(.+?)\s+EMISION\s*:\s*[A-ZÁÉÍÓÚÑ]+", texto_norm, flags=re.IGNORECASE)
    if match_matriz:
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(match_matriz.group(1))
    idx_matriz = _buscar_indice_linea_layout(lineas, "MATRIZ")
    if idx_matriz is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(_siguiente_linea_layout_no_vacia(lineas, idx_matriz))
    row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*(SI|NO)"]))

    idx_razon = _buscar_indice_linea_layout(lineas, "RAZON SOCIAL / NOMBRES Y APELLIDOS")
    if idx_razon is not None:
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(_texto_linea_layout(lineas[idx_razon], min_x=180))
    identificacion = _extraer_regex(texto_norm, [r"IDENTIFICACION\s*:?\s*(\d{10,13})"])
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(identificacion)
    row["Tipo Identificación Comprador"] = _codigo_tipo_identificacion_desde_numero(identificacion, "No Disponible")
    row["Fecha de Emisión"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"FECHA\s+EMISION\s*:?\s*(\d{2}/\d{2}/\d{4})"]))

    modificado = _extraer_regex(texto_norm, [r"COMPROBANTE\s+QUE\s+SE\s+MODIFICA\s*:?\s*([^\n]+)"])
    row["Número Documento Modificado"] = _texto_emitidos_retencion_na(_extraer_regex(modificado, [r"(\d{3}-\d{3}-\d{9})"]))
    row["Código Documento Modificado"] = _codigo_documento_sri(modificado)
    row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"FECHA\s+EMISION\s+\(COMPROBANTE\s+A\s+MODIFICAR\)\s*:?\s*(\d{2}/\d{2}/\d{4})"]))
    row["Motivo"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"RAZON\s+DE\s+MODIFICACION\s*:?\s*([^\n]+)"]))

    idx_info_adic = _buscar_indice_linea_layout(lineas, "INFORMACION ADICIONAL")
    top_info = float(lineas[idx_info_adic].get("top", 0.0)) if idx_info_adic is not None else 9999.0
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(_extraer_campos_adicionales_por_layout(lineas, top_info))

    top_start_detalle = 0.0
    idx_detalle = _buscar_indice_linea_layout(lineas, "CODIGO DESCRIPCION")
    if idx_detalle is not None:
        top_start_detalle = float(lineas[idx_detalle].get("top", 0.0))
    top_end_detalle = top_info if idx_info_adic is not None else 9999.0
    region_detalle = [ln for ln in lineas if top_start_detalle < float(ln.get("top", 0.0)) < top_end_detalle]
    items = []
    current_item = None
    for linea in region_detalle:
        left_digits = [w.get("text", "") for w in linea.get("words", []) if float(w.get("x0", 0.0)) < 70 and re.fullmatch(r"\d+", (w.get("text", "") or ""))]
        if left_digits:
            chunk = "".join(left_digits)
            if current_item is None or len(chunk) >= 8:
                if current_item:
                    items.append(current_item)
                current_item = {"code": chunk, "desc_lines": [], "qty": "", "unit": ""}
            elif current_item is not None:
                current_item["code"] += chunk
        if current_item is None:
            continue
        desc_part = _texto_linea_layout(linea, min_x=140, max_x=360)
        if desc_part:
            current_item["desc_lines"].append(desc_part)
        qty_tokens = [w.get("text", "") for w in linea.get("words", []) if 100 <= float(w.get("x0", 0.0)) < 140 and re.fullmatch(r"\d+(?:\.\d+)?", (w.get("text", "") or ""))]
        if qty_tokens and not current_item["qty"]:
            current_item["qty"] = qty_tokens[0]
        unit_tokens = [w.get("text", "") for w in linea.get("words", []) if 480 <= float(w.get("x0", 0.0)) < 530 and re.fullmatch(r"\d+(?:\.\d+)?", (w.get("text", "") or ""))]
        if unit_tokens and not current_item["unit"]:
            current_item["unit"] = unit_tokens[0]
    if current_item:
        items.append(current_item)
    detalles = []
    for item in items:
        desc_lines = [fragment for fragment in item["desc_lines"] if fragment]
        desc = "  ".join(desc_lines).strip()
        qty_num = _numero_emitidos_retencion(item.get("qty"), 0)
        unit_num = _numero_emitidos_retencion(item.get("unit"), 0)
        detalles.append(f"Código: {item.get('code')}, Desc: {desc}, Cant: {qty_num:.6f}, P.Unit: {unit_num:.6f}")
    if detalles:
        row["Descripciones"] = " | ".join(detalles)

    row["Total Sin Impuestos"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*([0-9.,]+)"]))
    row["Base Gravada 15%"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+15%\s*([0-9.,]+)"]))
    row["Base Gravada"] = row["Base Gravada 15%"]
    row["Base No Gravada"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+NO\s+OBJETO\s+DE\s+IVA\s*([0-9.,]+)"]))
    row["Monto IVA 15%"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+15%\s*([0-9.,]+)"]))
    row["Monto IVA"] = row["Monto IVA 15%"]
    if row["Monto IVA 15%"]:
        row["Tarifas IVA"] = "15%"
    row["Importe Total"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"VALOR\s+TOTAL\s*([0-9.,]+)"]))
    row["Valor Modificación"] = row["Importe Total"]

    legacy = _map_nota_credito_legada_a_emitidos_row(_extraer_datos_pdf_nota_credito(pdf_path))
    return _combinar_rows_emitidos_especificos(row, legacy)


def _extraer_datos_pdf_nota_debito_emitido(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_layout_pdf(pdf_path)
    if not lineas:
        return _map_nota_debito_legada_a_emitidos_row(_extraer_datos_pdf_nota_debito(pdf_path))

    row = _nota_debito_emitidos_default_row()
    row["Estado"] = "AUTORIZADO"
    row["Código del Documento"] = "05 - NOTA DE DÉBITO"
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Moneda"] = "DOLAR"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Contribuyente RIMPE"] = "No Disponible"

    texto = _leer_texto_pdf(pdf_path)
    texto_norm = _normalizar_texto_pdf(texto)
    lineas_raw = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    auth = _extraer_regex(texto_norm, [r"(\d{49})"])
    numero = _extraer_regex(texto_norm, [r"No\.\s*(\d{3}-\d{3}-\d{9})"])
    ruc = _extraer_regex(texto_norm, [r"R\.U\.C\.\s*:?\s*(\d{13})"])
    fecha_hora_auth = _extraer_regex(texto_norm, [r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"])
    row["Número de Autorización"] = auth or "No Disponible"
    row["Clave de Acceso"] = auth or "No Disponible"
    row["Fecha de Autorización"] = _fecha_hora_pdf_a_iso(fecha_hora_auth)
    row["RUC Emisor"] = ruc or "No Disponible"
    if numero:
        row["Establecimiento"], row["Punto de Emisión"], row["Secuencial"] = numero.split("-")
    if auth and len(auth) == 49:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(auth[23])
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(auth[47])
    else:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(
            _extraer_regex(texto_norm, [r"AMBIENTE\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"])
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(
            _extraer_regex(texto_norm, [r"EMISI[ÓO]N\s*:?\s*([A-ZÁÉÍÓÚÑ]+)"])
        )

    idx_num_aut = _buscar_indice_linea_layout(lineas, "NUMERO DE AUTORIZACION")
    if idx_num_aut is not None:
        row["Razón Social Emisor"] = _texto_emitidos_retencion_na(
            _siguiente_linea_layout_no_vacia(lineas, idx_num_aut + 1, max_x=260)
        )
    idx_aut = _buscar_indice_linea_layout_exacta(lineas, "AUTORIZACION")
    if idx_aut is not None:
        for pos in range(idx_aut + 1, len(lineas)):
            candidato = _texto_linea_layout(lineas[pos], max_x=260)
            if candidato and candidato != row["Razón Social Emisor"]:
                row["Nombre Comercial"] = _texto_emitidos_retencion_na(candidato)
                break

    match_matriz = re.search(r"Direccion\s+(.+?)\s+EMISION\s*:\s*[A-ZÁÉÍÓÚÑ]+", texto_norm, flags=re.IGNORECASE)
    if match_matriz:
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(match_matriz.group(1))
    idx_matriz = _buscar_indice_linea_layout(lineas, "MATRIZ")
    if idx_matriz is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(_siguiente_linea_layout_no_vacia(lineas, idx_matriz))
    row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(
        _extraer_regex(texto_norm, [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*(SI|NO)"])
    )

    idx_razon = _buscar_indice_linea_layout(lineas, "RAZON SOCIAL / NOMBRES Y APELLIDOS")
    if idx_razon is not None:
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(_texto_linea_layout(lineas[idx_razon], min_x=180))
    identificacion = _extraer_regex(texto_norm, [r"RUC/CI/\.?/PASAPORTE\s*:?\s*(\d{10,13})", r"IDENTIFICACION\s*:?\s*(\d{10,13})"])
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(identificacion)
    row["Tipo Identificación Comprador"] = _codigo_tipo_identificacion_desde_numero(identificacion, "No Disponible")
    row["Fecha de Emisión"] = _texto_emitidos_retencion_na(
        _extraer_regex(texto_norm, [r"FECHA\s+EMISION\s*:?\s*(\d{2}/\d{2}/\d{4})"])
    )

    modificado = _extraer_regex(texto_norm, [r"COMPROBANTE\s+QUE\s+SE\s+MODIFICA\s*:?\s*([^\n]+)"])
    row["Número Documento Modificado"] = _texto_emitidos_retencion_na(
        _extraer_regex(modificado, [r"(\d{3}-\d{3}-\d{9})"])
    )
    row["Código Documento Modificado"] = _codigo_documento_sri(modificado)
    row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(
        _extraer_regex(
            texto_norm,
            [
                r"FECHA\s+EMISION\s+\(COMPROBANTE\s+A\s+MODIFICAR\)\s*:?\s*(\d{2}/\d{2}/\d{4})",
                r"COMPROBANTE\s+QUE\s+SE\s+MODIFICA.+?FECHA\s+EMISION\s*:?\s*(\d{2}/\d{2}/\d{4})",
            ],
        )
    )
    row["Motivo"] = _texto_emitidos_retencion_na(
        _extraer_regex(texto_norm, [r"RAZON\s+DE\s+(?:LA\s+)?MODIFICACION\s*:?\s*([^\n]+)"])
    )

    valor_mod = _numero_emitidos_retencion(
        _extraer_regex(texto_norm, [r"VALOR\s+DE\s+LA\s+MODIFICACION\s*([0-9.,]+)"])
    )
    total_sin_imp = _numero_emitidos_retencion(
        _extraer_regex(texto_norm, [r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*([0-9.,]+)"])
    )
    if not total_sin_imp:
        total_sin_imp = valor_mod
    row["Total Sin Impuestos"] = total_sin_imp
    row["Base Gravada 15%"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+15%\s*([0-9.,]+)"]))
    row["Base Gravada"] = (
        row["Base Gravada 15%"]
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+12%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+8%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+5%\s*([0-9.,]+)"]))
    )
    row["Base No Gravada"] = (
        _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+0%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"SUBTOTAL\s+NO\s+OBJETO\s+DE\s+IVA\s*([0-9.,]+)"]))
    )
    row["Monto IVA 15%"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+15%\s*([0-9.,]+)"]))
    row["Monto IVA"] = (
        row["Monto IVA 15%"]
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+12%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+8%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+5%\s*([0-9.,]+)"]))
    )
    if row["Monto IVA 15%"]:
        row["Tarifas IVA"] = "15%"
    elif row["Monto IVA"]:
        row["Tarifas IVA"] = "IVA"
    row["Importe Total"] = _numero_emitidos_retencion(
        _extraer_regex(
            texto_norm,
            [r"VALOR\s+A\s+PAGAR\s*([0-9.,]+)", r"VALOR\s+TOTAL\s*([0-9.,]+)"],
        )
    )
    row["Valor Modificación"] = valor_mod or row["Importe Total"]
    row["Total Pago"] = row["Importe Total"]

    forma_pago, monto_pago = _extraer_forma_pago(lineas_raw)
    if forma_pago:
        row["Forma Pago"] = f"{forma_pago} - {forma_pago}"
    if monto_pago:
        row["Total Pago"] = _numero_emitidos_retencion(monto_pago)

    idx_info_adic = _buscar_indice_linea_layout(lineas, "INFORMACION ADICIONAL")
    top_info = float(lineas[idx_info_adic].get("top", 0.0)) if idx_info_adic is not None else 9999.0
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(_extraer_campos_adicionales_por_layout(lineas, top_info))
    if row["Motivo"] != "No Disponible":
        row["Descripciones"] = row["Motivo"]

    legacy = _map_nota_debito_legada_a_emitidos_row(_extraer_datos_pdf_nota_debito(pdf_path))
    return _combinar_rows_emitidos_especificos(row, legacy)


def _map_factura_legada_a_emitidos_row(legacy: dict | None) -> dict:
    row = _factura_emitidos_default_row()
    if not isinstance(legacy, dict):
        return row
    row["Ambiente"] = _label_ambiente_emitidos_retencion(legacy.get("ambiente"))
    row["Razón Social Emisor"] = _texto_emitidos_retencion_na(legacy.get("razonSocialEmisor"))
    row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
        legacy.get("direccionSucursalEmisor") or legacy.get("direccionMatrizEmisor")
    )
    obligado = _texto_emitidos_retencion(legacy.get("obligadoContabilidad"))
    row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
    row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
        _codigo_tipo_identificacion_desde_numero(legacy.get("identificacionComprador"))
    )
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(legacy.get("identificacionComprador"))
    row["Tipo Emisión"] = _label_emision_emitidos_retencion(legacy.get("emision"))
    row["Nombre Comercial"] = _texto_emitidos_retencion_na(
        legacy.get("nombreComercial") or legacy.get("razonSocialEmisor")
    )
    row["Código del Documento"] = "01 - FACTURA"
    row["Establecimiento"] = _texto_emitidos_retencion(legacy.get("establecimiento"))
    row["Punto de Emisión"] = _texto_emitidos_retencion(legacy.get("puntoEmision"))
    row["Secuencial"] = _texto_emitidos_retencion(legacy.get("secuencial"))
    row["Dirección Matriz"] = _texto_emitidos_retencion_na(legacy.get("direccionMatrizEmisor"))
    row["Contribuyente RIMPE"] = "No Disponible"
    row["RUC Emisor"] = _texto_emitidos_retencion(legacy.get("rucEmisor"))
    row["Clave de Acceso"] = _texto_emitidos_retencion(legacy.get("claveAcceso"))
    fecha_emision = _texto_emitidos_retencion(legacy.get("fechaEmision"))
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", fecha_emision):
        fecha_emision = datetime.strptime(fecha_emision, "%Y-%m-%d").strftime("%d/%m/%Y")
    row["Fecha de Emisión"] = fecha_emision or "No Disponible"
    row["Razón Social Comprador"] = _texto_emitidos_retencion_na(legacy.get("razonSocialComprador"))
    row["Dirección Comprador"] = _texto_emitidos_retencion_na(legacy.get("direccionComprador"))
    row["Moneda"] = _texto_emitidos_retencion_na(legacy.get("moneda") or "DOLAR")
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Descripciones"] = _texto_emitidos_retencion(legacy.get("descripcionesProductos"))
    row["Forma Pago"] = "No Disponible - No Disponible"
    total_sin_imp = _numero_emitidos_retencion(legacy.get("subtotalSinImpuestos"))
    row["Total Sin Impuestos"] = total_sin_imp
    row["Base Gravada"] = 0
    row["Base No Gravada"] = total_sin_imp
    row["Tarifas IVA"] = "0%"
    row["Monto IVA"] = (
        _numero_emitidos_retencion(legacy.get("iva15"))
        or _numero_emitidos_retencion(legacy.get("iva12"))
        or _numero_emitidos_retencion(legacy.get("iva8"))
        or _numero_emitidos_retencion(legacy.get("iva5"))
    )
    row["Total Descuento"] = _numero_emitidos_retencion(legacy.get("totalDescuento"))
    row["Propina"] = _numero_emitidos_retencion(legacy.get("propina"))
    row["Importe Total"] = _numero_emitidos_retencion(legacy.get("valorTotal"))
    row["Total Pago"] = 0
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(legacy.get("informacionAdicional"))
    row["Base No Gravada 0%"] = total_sin_imp
    return row


def _extraer_datos_xml_factura_emitido(xml_path: Path) -> dict:
    row = _factura_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_fact = root.find("infoFactura")
    detalles = root.findall(".//detalles/detalle")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(
            info_trib.findtext("nombreComercial") or info_trib.findtext("razonSocial")
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, "01 - FACTURA")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_fact is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
            info_fact.findtext("dirEstablecimiento") or row["Dirección Matriz"]
        )
        obligado = _texto_emitidos_retencion(info_fact.findtext("obligadoContabilidad"))
        row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_fact.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_fact.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("razonSocialComprador"))
        row["Dirección Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("direccionComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_fact.findtext("moneda") or "DOLAR")
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_fact.findtext("totalSinImpuestos"))
        row["Total Descuento"] = _numero_emitidos_retencion(info_fact.findtext("totalDescuento"))
        row["Propina"] = _numero_emitidos_retencion(info_fact.findtext("propina"))
        row["Importe Total"] = _numero_emitidos_retencion(info_fact.findtext("importeTotal"))

        pago = info_fact.find("./pagos/pago")
        if pago is not None:
            forma = _label_forma_pago_emitidos_retencion(pago.findtext("formaPago"))
            row["Forma Pago"] = f"{forma} - {forma}" if forma != "No Disponible" else "No Disponible - No Disponible"
            row["Total Pago"] = _numero_emitidos_retencion(pago.findtext("total"))
            row["Plazo Pago"] = _texto_emitidos_retencion_na(pago.findtext("plazo"))
            row["Unidad Tiempo Pago"] = _texto_emitidos_retencion_na(pago.findtext("unidadTiempo"))
        else:
            row["Forma Pago"] = "No Disponible - No Disponible"

    monto_iva = 0
    for imp in root.findall(".//infoFactura/totalConImpuestos/totalImpuesto"):
        valor = _numero_emitidos_retencion(imp.findtext("valor"))
        if valor:
            monto_iva += valor
    row["Base Gravada"] = 0
    row["Base No Gravada"] = row["Total Sin Impuestos"]
    row["Base No Gravada 0%"] = row["Total Sin Impuestos"]
    row["Tarifas IVA"] = "0%"
    row["Monto IVA"] = monto_iva

    detalle_textos = []
    for detalle in detalles:
        codigo = _texto_emitidos_retencion(detalle.findtext("codigoPrincipal") or detalle.findtext("codigoInterno"))
        descripcion = (detalle.findtext("descripcion") or "").strip()
        cantidad = _texto_emitidos_retencion(detalle.findtext("cantidad"))
        precio_unitario = _texto_emitidos_retencion(detalle.findtext("precioUnitario"))
        partes = [f"Código: {codigo}", f"Desc: {descripcion}"]
        if cantidad:
            partes.append(f"Cant: {cantidad}")
        if precio_unitario:
            partes.append(f"P.Unit: {precio_unitario}")
        detalle_textos.append(", ".join(partes))
    row["Descripciones"] = " ; ".join(detalle_textos)

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)
    return row


def _extraer_datos_pdf_factura_emitido(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_layout_pdf(pdf_path)
    legacy = _map_factura_legada_a_emitidos_row(_extraer_datos_pdf_por_tipo_layout_first(pdf_path))
    if not lineas:
        return legacy

    row = _factura_emitidos_default_row()
    row["Estado"] = "AUTORIZADO"
    row["Código del Documento"] = "01 - FACTURA"
    row["Contribuyente RIMPE"] = "No Disponible"
    row["Moneda"] = "DOLAR"
    row["Plazo Pago"] = "No Disponible"
    row["Unidad Tiempo Pago"] = "No Disponible"
    row["Forma Pago"] = "No Disponible - No Disponible"

    texto = _leer_texto_pdf(pdf_path)
    texto_norm = _normalizar_texto_pdf(texto)
    auth = _extraer_regex(texto_norm, [r"(\d{49})"])
    numero = _extraer_regex(texto_norm, [r"No\.\s*(\d{3}-\d{3}-\d{9})"])
    ruc = _extraer_regex(texto_norm, [r"R\.U\.C\.\s*:?\s*(\d{13})"])
    fecha_hora_auth = _extraer_regex(texto_norm, [r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"])
    row["Número de Autorización"] = auth or "No Disponible"
    row["Clave de Acceso"] = auth or "No Disponible"
    row["Fecha de Autorización"] = _fecha_hora_pdf_a_iso(fecha_hora_auth)
    row["RUC Emisor"] = ruc or "No Disponible"
    if numero:
        row["Establecimiento"], row["Punto de Emisión"], row["Secuencial"] = numero.split("-")
    if auth and len(auth) == 49:
        row["Ambiente"] = _label_ambiente_emitidos_retencion(auth[23])
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(auth[47])

    idx_num_aut = _buscar_indice_linea_layout(lineas, "NUMERO DE AUTORIZACION")
    if idx_num_aut is not None:
        row["Razón Social Emisor"] = _texto_emitidos_retencion_na(
            _siguiente_linea_layout_no_vacia(lineas, idx_num_aut + 1, max_x=260)
        )
    row["Nombre Comercial"] = row["Razón Social Emisor"]

    idx_matriz = _buscar_indice_linea_layout(lineas, "MATRIZ")
    direccion = _extraer_bloque_direccion_layout(lineas, idx_matriz)
    if direccion:
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(direccion)
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(direccion)
    obligado = _extraer_regex(texto_norm, [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*(SI|NO)"])
    row["Obligado Contabilidad"] = obligado or "No Disponible"

    idx_razon = _buscar_indice_linea_layout(lineas, "RAZON SOCIAL / NOMBRES Y APELLIDOS")
    if idx_razon is not None:
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(_texto_linea_layout(lineas[idx_razon], min_x=180))
    identificacion = _extraer_regex(texto_norm, [r"IDENTIFICACION\s*:?\s*(\d{10,13})"])
    row["Identificación Comprador"] = _texto_emitidos_retencion_na(identificacion)
    row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
        _codigo_tipo_identificacion_desde_numero(identificacion)
    )
    row["Fecha de Emisión"] = _texto_emitidos_retencion_na(_extraer_regex(texto_norm, [r"FECHA\s+(\d{2}/\d{2}/\d{4})"]))
    idx_direccion_comprador = _buscar_indice_linea_layout(lineas, "DIRECCION:", start=(idx_razon or 0))
    if idx_direccion_comprador is not None:
        row["Dirección Comprador"] = _texto_emitidos_retencion_na(
            _texto_linea_layout(lineas[idx_direccion_comprador], min_x=80)
        )

    idx_info = _buscar_indice_linea_layout(lineas, "INFORMACION ADICIONAL")
    top_info = float(lineas[idx_info].get("top", 0.0)) if idx_info is not None else 9999.0
    idx_detalle = _buscar_indice_linea_layout(lineas, "CANTIDAD DESCRIPCION")
    top_inicio = float(lineas[idx_detalle].get("top", 0.0)) if idx_detalle is not None else 450.0
    items = _extraer_items_emitidos_layout(lineas, top_inicio=top_inicio, top_fin=top_info)
    if items:
        row["Descripciones"] = _formatear_descripciones_emitidos(items)
    row["Campos Adicionales"] = _texto_emitidos_retencion_na(_extraer_campos_adicionales_emitidos_desde_texto(texto))

    row["Total Sin Impuestos"] = _numero_emitidos_retencion(
        _extraer_regex(texto_norm, [r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*([0-9.,]+)"])
    )
    row["Base Gravada"] = 0
    row["Base No Gravada"] = row["Total Sin Impuestos"]
    row["Base No Gravada 0%"] = row["Total Sin Impuestos"]
    row["Tarifas IVA"] = "0%"
    row["Monto IVA"] = (
        _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+15%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+12%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+8%\s*([0-9.,]+)"]))
        or _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"IVA\s+5%\s*([0-9.,]+)"]))
    )
    row["Total Descuento"] = _numero_emitidos_retencion(
        _extraer_regex(texto_norm, [r"TOTAL\s+DESCUENTO\s*([0-9.,]+)"])
    )
    row["Propina"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"PROPINA\s*([0-9.,]+)"]))
    row["Importe Total"] = _numero_emitidos_retencion(_extraer_regex(texto_norm, [r"VALOR\s+TOTAL\s*([0-9.,]+)"]))
    row["Total Pago"] = 0

    return _combinar_rows_emitidos_especificos(row, legacy)


def _extraer_datos_xml_liquidacion_compra_emitido(xml_path: Path) -> dict:
    datos = _extraer_datos_xml_pdf_report(xml_path)
    datos["tipoDocumento"] = "Liquidación de Compra"
    if not _valor_reporte_presente(datos.get("nombreComercial")):
        datos["nombreComercial"] = datos.get("razonSocialEmisor") or "No Disponible"
    datos["direccionSucursalEmisor"] = datos.get("direccionSucursalEmisor") or "No Disponible"
    if _texto_emitidos_retencion(datos.get("agenteRetencion")):
        datos["agenteRetencion"] = str(datos["agenteRetencion"]).replace("Resolucion", "Resolución")
    obligado = _texto_emitidos_retencion(datos.get("obligadoContabilidad"))
    datos["obligadoContabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
    datos["tipoContribuyenteRIMPE"] = datos.get("tipoContribuyenteRIMPE") or "No Disponible"
    datos["razonSocialComprador"] = "No Disponible"
    if not _valor_reporte_presente(datos.get("direccionComprador")):
        datos["direccionComprador"] = ""
    if not _valor_reporte_presente(datos.get("informacionAdicional")):
        datos["informacionAdicional"] = "No Disponible"
    return datos


def _extraer_datos_pdf_liquidacion_compra_emitido(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_layout_pdf(pdf_path)
    legacy = _extraer_datos_pdf_por_tipo_layout_first(pdf_path)
    if not lineas:
        return legacy

    datos = dict(legacy or {})
    texto = _leer_texto_pdf(pdf_path)
    texto_norm = _normalizar_texto_pdf(texto)
    datos["tipoDocumento"] = "Liquidación de Compra"
    datos["contribuyenteEspecial"] = "No Disponible"
    ruc = _extraer_regex(texto_norm, [r"R\.U\.C\.\s*:?\s*(\d{13})"])
    if ruc:
        datos["rucEmisor"] = ruc
    auth = _extraer_regex(texto_norm, [r"(\d{49})"])
    if auth:
        datos["claveAcceso"] = auth
    idx_aut = _buscar_indice_linea_layout_exacta(lineas, "AUTORIZACION")
    if idx_aut is not None:
        emisor = _siguiente_linea_layout_no_vacia(lineas, idx_aut, max_x=320)
        if emisor:
            datos["razonSocialEmisor"] = emisor
            datos["nombreComercial"] = emisor
    direccion_matriz = _extraer_regex(texto_norm, [r"DIRECCION\s+(.+?)\s+EMISION\s*:"])
    idx_matriz = _buscar_indice_linea_layout(lineas, "MATRIZ")
    if not direccion_matriz:
        direccion_matriz = _extraer_bloque_direccion_layout(
            lineas,
            idx_matriz,
            stop_tokens=("CLAVE", "SUCURSAL", "OBLIGADO", "AGENTE"),
        )
    if direccion_matriz:
        datos["direccionMatrizEmisor"] = direccion_matriz
    datos["direccionSucursalEmisor"] = "No Disponible"
    agente = _extraer_regex(texto_norm, [r"AGENTE\s+DE\s+RETENCION\s+([^\n]+)"])
    datos["agenteRetencion"] = agente.replace("Resolucion", "Resolución") if agente else "No Disponible"
    obligado = _extraer_regex(texto_norm, [r"OBLIGADO\s+A\s+LLEVAR\s+CONTABILIDAD\s*(SI|NO)\b"])
    datos["obligadoContabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
    datos["tipoContribuyenteRIMPE"] = "No Disponible"
    datos["razonSocialComprador"] = "No Disponible"
    identificacion = _extraer_regex(texto_norm, [r"IDENTIFICACION\s*:?\s*(\d{6,13})"])
    if identificacion:
        datos["identificacionComprador"] = identificacion
    datos["direccionComprador"] = ""
    datos["placa"] = "No Disponible"
    datos["guia"] = "No Disponible"
    datos["comprobanteModificado"] = "No Disponible"
    datos["fechaEmisionModificado"] = "No Disponible"
    datos["razonModificacion"] = "No Disponible"
    datos["valorModificacion"] = "No Disponible"
    fecha_auth = _extraer_regex(texto_norm, [r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})"])
    if fecha_auth:
        datos["fechaAutorizacion"] = fecha_auth
    fecha_emision = _extraer_regex(texto_norm, [r"FECHA\s+(\d{2}/\d{2}/\d{4})"])
    if fecha_emision:
        datos["fechaEmision"] = fecha_emision
    numero = _extraer_regex(texto_norm, [r"No\.\s*(\d{3}-\d{3}-\d{9})"])
    if numero:
        datos["numeroComprobante"] = numero
        estab, pto, sec = numero.split("-")
        datos["establecimiento"] = estab
        datos["puntoEmision"] = pto
        datos["secuencial"] = sec

    idx_inicio_detalle = _buscar_indice_linea_layout(lineas, "PRINCIPAL AUXILIAR")
    idx_code = idx_inicio_detalle + 1 if idx_inicio_detalle is not None and idx_inicio_detalle + 2 < len(lineas) else None
    if idx_code is not None and idx_code + 2 < len(lineas):
        codigo = (lineas[idx_code].get("text", "") or "").strip()
        valores = re.findall(r"\d+(?:\.\d+)?", lineas[idx_code + 1].get("text", ""))
        extra = (lineas[idx_code + 2].get("text", "") or "").strip()
        cantidad = valores[0] if len(valores) >= 1 else ""
        descripcion_linea = lineas[idx_code + 1].get("text", "")
        descripcion = re.sub(r"^\d+(?:\.\d+)?\s*", "", descripcion_linea).strip()
        if len(valores) > 1 and valores[1] in descripcion:
            descripcion = descripcion.split(valores[1], 1)[0].strip()
        codigo_final = f"{codigo} {extra}".strip()
        subsidio = valores[2] if len(valores) > 2 else "0.00"
        precio_sin_sub = valores[3] if len(valores) > 3 else "0.00"
        descuento = valores[4] if len(valores) > 4 else "0.00"
        total = valores[5] if len(valores) > 5 else ""
        datos["descripcionesProductos"] = (
            f"Código: {codigo_final}, Aux: , Cant: {cantidad}, Desc: {descripcion}, "
            f"P.Unit: {valores[1] if len(valores) > 1 else ''}, Subsidio: {subsidio}, "
            f"P. s/Sub: {precio_sin_sub}, Descuento: {descuento}, P.Total: {total}"
        ).strip()

    forma_match = re.search(
        r"(20\s*-\s*OTROS\s+CON\s+UTILIZACION\s+DEL\s+SISTEMA(?:\s+FINANCIERO)?)\s+([0-9.,]+)",
        texto_norm,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if forma_match:
        datos["formaPago"] = re.sub(r"\s+", " ", forma_match.group(1)).strip()
        datos["formaPagoMonto"] = _parse_decimal(forma_match.group(2))
    info_adicional = _extraer_campos_adicionales_emitidos_desde_texto(texto)
    if info_adicional:
        datos["informacionAdicional"] = info_adicional.replace("; ", "\n")

    if auth and len(auth) == 49:
        datos["ambiente"] = "PRODUCCIÓN" if auth[23] == "2" else "PRUEBAS"
        datos["emision"] = "NORMAL" if auth[47] == "1" else datos.get("emision") or "NORMAL"
    return datos


def _extraer_datos_xml_pdf_report(xml_path: Path) -> dict:
    datos = {col: "" for col in PDF_REPORT_COLUMNS}
    try:
        (
            cabecera,
            detalles,
            _impuestos,
            pagos,
            adicionales,
            _retenciones,
            error_entry,
            cabecera_tax_cols,
        ) = xml_parser._parse_recibido_xml(xml_path)
    except Exception:
        return datos
    if error_entry or cabecera is None:
        return datos

    def _safe(val):
        return "" if val is None else str(val).strip()

    def _safe_or_na(val):
        valor = _safe(val)
        return valor if valor else "No Disponible"

    def _safe_num(val):
        if val is None:
            return "0"
        try:
            return str(val).strip() if str(val).strip() else "0"
        except Exception:
            return "0"

    def _map_ambiente(val):
        v = _safe(val).upper()
        if "PRODUCCION" in v:
            return "PRODUCCION"
        if "PRUEBAS" in v:
            return "PRUEBAS"
        if v == "1":
            return "PRUEBAS"
        if v == "2":
            return "PRODUCCION"
        return v

    def _map_emision(val):
        v = _safe(val).upper()
        if "NORMAL" in v:
            return "NORMAL"
        if "INDISPONIBILIDAD" in v or "CONTINGENCIA" in v:
            return "CONTINGENCIA"
        if v == "1":
            return "NORMAL"
        if v == "2":
            return "CONTINGENCIA"
        return v

    datos["tipoDocumento"] = xml_parser.DOC_TYPES.get(cabecera.get("COD_DOC"), cabecera.get("DESCRIPCION_DOC", ""))
    datos["rucEmisor"] = _safe(cabecera.get("RUC_EMISOR"))
    datos["razonSocialEmisor"] = _safe(cabecera.get("RAZON_SOCIAL_EMISOR"))
    datos["nombreComercial"] = _safe_or_na(cabecera.get("NOMBRE_COMERCIAL_EMISOR"))
    datos["direccionMatrizEmisor"] = _safe(cabecera.get("DIR_MATRIZ"))
    datos["direccionSucursalEmisor"] = _safe_or_na(cabecera.get("DIR_ESTABLECIMIENTO"))
    datos["obligadoContabilidad"] = _safe_or_na(cabecera.get("OBLIGADO_CONTABILIDAD"))
    datos["tipoContribuyenteRIMPE"] = _safe_or_na(cabecera.get("TIPO_CONTRIBUYENTE_RIMPE"))
    datos["numeroComprobante"] = _safe(cabecera.get("SERIE_COMPROBANTE"))
    datos["establecimiento"] = _safe(cabecera.get("ESTAB"))
    datos["puntoEmision"] = _safe(cabecera.get("PTO_EMI"))
    datos["secuencial"] = _safe(cabecera.get("SECUENCIAL"))
    datos["fechaEmision"] = _safe(cabecera.get("FECHA_EMISION"))
    datos["fechaAutorizacion"] = _safe(cabecera.get("FECHA_AUTORIZACION"))
    datos["razonSocialComprador"] = _safe(cabecera.get("RAZON_SOCIAL_COMPRADOR"))
    datos["identificacionComprador"] = _safe(cabecera.get("IDENTIFICACION_COMPRADOR"))
    datos["direccionComprador"] = _safe(cabecera.get("DIRECCION_COMPRADOR"))
    datos["contribuyenteEspecial"] = _safe_or_na(cabecera.get("CONTRIBUYENTE_ESPECIAL"))
    datos["agenteRetencion"] = _safe_or_na(cabecera.get("AGENTE_RETENCION"))

    datos["comprobanteModificado"] = _safe_or_na(cabecera.get("NUM_DOC_MODIFICADO") or cabecera.get("COD_DOC_MODIFICADO"))
    datos["fechaEmisionModificado"] = _safe_or_na(cabecera.get("FECHA_EMISION_DOC_SUSTENTO"))
    datos["razonModificacion"] = _safe_or_na(cabecera.get("MOTIVO_MODIFICACION") or cabecera.get("MOTIVO") or cabecera.get("MOTIVOS_DESCRIPCION"))
    valor_mod = cabecera.get("VALOR_MODIFICACION_XML")
    if not valor_mod:
        valor_mod = cabecera.get("VALOR_MODIFICACION")
    datos["valorModificacion"] = _safe_or_na(valor_mod)

    datos["placa"] = _safe_or_na(cabecera.get("PLACA"))
    datos["guia"] = _safe_or_na(cabecera.get("GUIA"))

    detalles_texto = []
    for detalle in detalles:
        descripcion = detalle.get("DESCRIPCION") or detalle.get("descripcion") or ""
        cantidad = detalle.get("CANTIDAD") or detalle.get("cantidad") or ""
        precio_unitario = detalle.get("PRECIO_UNITARIO") or detalle.get("precioUnitario") or ""
        descuento = detalle.get("DESCUENTO") or detalle.get("descuento") or ""
        total = detalle.get("PRECIO_TOTAL_SIN_IMPUESTO") or detalle.get("precioTotal") or ""
        partes = []
        if descripcion:
            partes.append(f"Desc: {descripcion}")
        if cantidad != "":
            partes.append(f"Cant: {cantidad}")
        if precio_unitario != "":
            partes.append(f"P.Unit: {precio_unitario}")
        if descuento != "":
            partes.append(f"Desc: {descuento}")
        if total != "":
            partes.append(f"P.Total: {total}")
        if partes:
            detalles_texto.append(", ".join(partes))
    datos["descripcionesProductos"] = "\n".join(detalles_texto) if detalles_texto else "No Disponible"

    datos["subtotalSinImpuestos"] = _safe_num(cabecera.get("TOTAL_SIN_IMPUESTOS"))
    datos["totalDescuento"] = _safe_num(cabecera.get("TOTAL_DESCUENTO"))
    datos["propina"] = _safe_num(cabecera.get("PROPINA"))
    valor_total = cabecera.get("IMPORTE_TOTAL") or cabecera.get("VALOR_TOTAL") or cabecera.get("VALOR_MODIFICACION")
    datos["valorTotal"] = _safe_num(valor_total)
    datos["valorTotalSinSubsidio"] = "0"

    tax_fallback = {}
    if _impuestos:
        for row in _impuestos:
            if str(row.get("NIVEL", "")).upper() != "DOCUMENTO":
                continue
            codigo = str(row.get("CODIGO", "")).strip()
            codigo_pct = str(row.get("CODIGO_PORCENTAJE", "")).strip()
            base = row.get("BASE_IMPONIBLE")
            valor = row.get("VALOR")
            base_col, valor_col = xml_parser._tax_column_names(codigo, codigo_pct)
            tax_fallback[base_col] = (tax_fallback.get(base_col, 0.0) or 0.0) + (base or 0.0)
            tax_fallback[valor_col] = (tax_fallback.get(valor_col, 0.0) or 0.0) + (valor or 0.0)
        if not tax_fallback:
            for row in _impuestos:
                codigo = str(row.get("CODIGO", "")).strip()
                codigo_pct = str(row.get("CODIGO_PORCENTAJE", "")).strip()
                base = row.get("BASE_IMPONIBLE")
                valor = row.get("VALOR")
                base_col, valor_col = xml_parser._tax_column_names(codigo, codigo_pct)
                tax_fallback[base_col] = (tax_fallback.get(base_col, 0.0) or 0.0) + (base or 0.0)
                tax_fallback[valor_col] = (tax_fallback.get(valor_col, 0.0) or 0.0) + (valor or 0.0)

    def _tax_base(slug):
        key = f"{slug}_BASE"
        val = cabecera_tax_cols.get(key)
        if _es_cero(val):
            val = tax_fallback.get(key, val)
        return _safe_num(val)

    def _tax_valor(slug):
        key = f"{slug}_VALOR"
        val = cabecera_tax_cols.get(key)
        if _es_cero(val):
            val = tax_fallback.get(key, val)
        return _safe_num(val)

    def _es_cero(valor: str) -> bool:
        if valor is None:
            return True
        txt = str(valor).strip()
        if txt in ("", "0", "0.0", "0.00", "0,0", "0,00"):
            return True
        parsed = _parse_decimal(txt) if isinstance(txt, str) else None
        if parsed is None:
            return False
        return abs(parsed) < 1e-9

    datos["subtotal0"] = _tax_base("IVA_0")
    datos["subtotal12"] = _tax_base("IVA_12")
    datos["subtotal15"] = _tax_base("IVA_15")
    datos["subtotal5"] = _tax_base("IVA_5")
    datos["subtotalNoObjetoIVA"] = _tax_base("NO_OBJETO_IVA")
    datos["subtotalExentoIVA"] = _tax_base("EXENTO_IVA")

    base_iva8 = _tax_base("IVA_8")
    val_iva8 = _tax_valor("IVA_8")
    base_tarifa = _tax_base("IVA_14")
    val_tarifa = _tax_valor("IVA_14")
    # Si no hay IVA_8 pero si hay IVA_14, asumimos que el codigo del XML viene como tarifa especial
    # y lo enviamos a la columna de IVA 8 (para evitar mezcla en el reporte).
    if _es_cero(base_iva8) and not _es_cero(base_tarifa):
        datos["subtotal8"] = base_tarifa
        datos["iva8"] = val_tarifa
        datos["subtotalTarifaEspecial"] = "0"
        datos["ivaTarifaEspecial"] = "0"
    else:
        datos["subtotal8"] = base_iva8
        datos["iva8"] = val_iva8
        datos["subtotalTarifaEspecial"] = base_tarifa
        datos["ivaTarifaEspecial"] = val_tarifa

    datos["iva12"] = _tax_valor("IVA_12")
    datos["iva15"] = _tax_valor("IVA_15")
    datos["iva5"] = _tax_valor("IVA_5")
    datos["ice"] = _tax_valor("ICE")
    datos["irbpnr"] = _tax_valor("IRBPNR")

    if pagos:
        datos["formaPago"] = _safe_or_na(pagos[0].get("FORMA_PAGO_DESC") or pagos[0].get("FORMA_PAGO"))
        datos["formaPagoMonto"] = _safe_num(pagos[0].get("TOTAL"))
    else:
        datos["formaPago"] = "No Disponible"
        datos["formaPagoMonto"] = "0"

    datos["ambiente"] = _map_ambiente(cabecera.get("AMBIENTE"))
    datos["emision"] = _map_emision(cabecera.get("TIPO_EMISION"))
    datos["claveAcceso"] = _safe(cabecera.get("CLAVE_ACCESO"))

    if adicionales:
        datos["informacionAdicional"] = "\n".join(
            f"{item.get('NOMBRE', '').strip()}: {item.get('VALOR', '').strip()}".strip(": ")
            for item in adicionales
            if (item.get("NOMBRE") or item.get("VALOR"))
        )
    if not datos["informacionAdicional"]:
        datos["informacionAdicional"] = "No Disponible"

    return datos

def _extraer_datos_pdf_nota_credito(pdf_path: Path) -> dict:
    datos = _extraer_datos_pdf(pdf_path)
    texto = _leer_texto_pdf(pdf_path)
    if not texto:
        return datos
    texto_norm = _normalizar_texto_pdf(texto)
    lineas_raw = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    lineas_norm = [_normalizar_label_simple(ln) for ln in lineas_raw]

    datos["tipoDocumento"] = "Nota de Crédito"
    if not datos.get("nombreComercial"):
        datos["nombreComercial"] = datos.get("razonSocialEmisor", "")
    if not datos.get("direccionSucursalEmisor") or datos.get("direccionSucursalEmisor") == "No Disponible":
        if datos.get("direccionMatrizEmisor"):
            datos["direccionSucursalEmisor"] = datos["direccionMatrizEmisor"]

    datos["contribuyenteEspecial"] = "No Disponible"
    datos["agenteRetencion"] = "No Disponible"
    datos["tipoContribuyenteRIMPE"] = "No Disponible"

    datos["comprobanteModificado"] = _extraer_regex(
        texto_norm,
        [r"COMPROBANTE\s+QUE\s+SE\s+MODIFICA\s*:\s*([^\n]+)"],
    )
    datos["fechaEmisionModificado"] = _extraer_regex(
        texto_norm,
        [r"FECHA\s+EMISION\s*\(COMPROBANTE\s+A\s+MODIFICAR\)\s*:\s*([^\n]+)"],
    )
    datos["razonModificacion"] = _extraer_regex(
        texto_norm,
        [r"RAZON\s+DE\s+MODIFICACION\s*:\s*([^\n]+)"],
    )
    if not datos["comprobanteModificado"]:
        datos["comprobanteModificado"] = "No Disponible"
    if not datos["fechaEmisionModificado"]:
        datos["fechaEmisionModificado"] = "No Disponible"
    if not datos["razonModificacion"]:
        datos["razonModificacion"] = "No Disponible"
    if not datos.get("valorModificacion"):
        datos["valorModificacion"] = "No Disponible"

    datos["descripcionesProductos"] = "No se extrajeron productos"
    datos["informacionAdicional"] = "No Disponible"
    datos["formaPago"] = "No Disponible"
    datos["formaPagoMonto"] = "No Disponible"

    if not datos.get("placa"):
        datos["placa"] = "No Disponible"
    if not datos.get("guia"):
        datos["guia"] = "No Disponible"

    dir_idx = None
    for idx, raw in enumerate(lineas_raw):
        if re.match(r"(?i)^direccion\s*:", raw):
            dir_idx = idx
            break
    if dir_idx is None:
        for idx, norm in enumerate(lineas_norm):
            if norm.startswith("DIRECCION") and "MATRIZ" not in norm and "SUCURSAL" not in norm:
                dir_idx = idx
                break
    if dir_idx is not None:
        direccion_lines = list(lineas_raw[dir_idx:])
        if direccion_lines:
            direccion_lines[0] = re.sub(r"(?i)^direccion\s*:\s*", "", direccion_lines[0]).strip()
        datos["direccionComprador"] = "\n".join(direccion_lines)

    return datos

def _extraer_datos_pdf_nota_debito(pdf_path: Path) -> dict:
    datos = _extraer_datos_pdf(pdf_path)
    texto = _leer_texto_pdf(pdf_path)
    if not texto:
        return datos
    texto_norm = _normalizar_texto_pdf(texto)
    lineas_raw = [ln.strip() for ln in texto.splitlines() if ln.strip()]

    datos["tipoDocumento"] = "Nota de Débito"

    if datos.get("razonSocialEmisor"):
        razon = datos["razonSocialEmisor"]
        if "FECHA Y HORA" in razon.upper():
            datos["razonSocialEmisor"] = razon.split("FECHA Y HORA", 1)[0].strip()
    if not datos.get("nombreComercial"):
        datos["nombreComercial"] = datos.get("razonSocialEmisor", "")

    match_matriz = re.search(r"DIRECCION\s+(.+?)\s+EMISION", texto_norm, flags=re.IGNORECASE)
    if match_matriz:
        datos["direccionMatrizEmisor"] = match_matriz.group(1).strip()
    datos["direccionSucursalEmisor"] = "No Disponible"

    contribuyente = _extraer_regex(texto_norm, [r"CONTRIBUYENTE\s+ESPECIAL\s*([0-9]+)"])
    if contribuyente:
        datos["contribuyenteEspecial"] = contribuyente
    else:
        datos["contribuyenteEspecial"] = "No Disponible"

    datos["agenteRetencion"] = "No Disponible"
    datos["tipoContribuyenteRIMPE"] = ""

    datos["comprobanteModificado"] = _extraer_regex(
        texto_norm,
        [r"COMPROBANTE\s+QUE\s+SE\s+MODIFICA\s*:\s*([^\n]+)"],
    )
    if not datos["comprobanteModificado"]:
        datos["comprobanteModificado"] = "No Disponible"

    datos["fechaEmisionModificado"] = "No Disponible"
    datos["razonModificacion"] = "No Disponible"
    datos["valorModificacion"] = "No Disponible"

    match_razon_valor = re.search(
        r"RAZON\s+DE\s+LA\s+MODIFICACION\s+VALOR\s+DE\s+LA\s+MODIFICACION\s+([^\n]+)",
        texto_norm,
        flags=re.IGNORECASE,
    )
    if match_razon_valor:
        resto = match_razon_valor.group(1).strip()
        partes = resto.rsplit(" ", 1)
        if len(partes) == 2:
            razon = partes[0].strip()
            valor = partes[1].strip()
            datos["descripcionesProductos"] = f"Razón: {razon}, Valor de Modificación: {valor}"
        else:
            datos["descripcionesProductos"] = f"Razón: {resto}"
    if not datos.get("descripcionesProductos"):
        datos["descripcionesProductos"] = "No se extrajeron productos"

    datos["placa"] = "No Disponible"
    datos["guia"] = "No Disponible"
    datos["formaPago"] = "No Disponible"
    datos["formaPagoMonto"] = "No Disponible"
    datos["informacionAdicional"] = "No Disponible"

    direccion_idx = None
    for idx, raw in enumerate(lineas_raw):
        if re.match(r"(?i)^direccion\s*:", raw):
            direccion_idx = idx
            break
    if direccion_idx is not None:
        direccion_lines = list(lineas_raw[direccion_idx:])
        if direccion_lines:
            direccion_lines[0] = re.sub(r"(?i)^direccion\s*:\s*", "", direccion_lines[0]).strip()
        datos["direccionComprador"] = "\n".join(direccion_lines)

    return datos

def _guardar_reporte_pdf_retencion_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in RETENCION_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[RETENCION_REPORT_COLUMNS]

    numeric_cols = [
        "Base_Imponible_Ret_IVA",
        "Porcentaje_Ret_IVA",
        "Valor_Retenido_IVA",
        "Base_Imponible_Ret_IR",
        "Porcentaje_Ret_IR",
        "Valor_Retenido_IR",
        "Base_Imponible_Ret_IR_1",
        "Porcentaje_Ret_IR_1",
        "Valor_Retenido_IR_1",
        "Base_Imponible_Ret_IVA_1",
        "Porcentaje_Ret_IVA_1",
        "Valor_Retenido_IVA_1",
    ]

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float):
            if pd.isna(val):
                return ""
            if val.is_integer():
                return str(int(val))
        return str(val).strip()

    def _to_number(val):
        if val is None:
            return ""
        if isinstance(val, float):
            if pd.isna(val):
                return ""
            return val
        if isinstance(val, int):
            return val
        if isinstance(val, str):
            if not val.strip():
                return ""
            parsed = _parse_decimal(val)
            return parsed if parsed is not None else val
        return val

    for col in df.columns:
        if col in numeric_cols:
            df[col] = df[col].map(_to_number)
        else:
            df[col] = df[col].map(_to_text)
    try:
        df.to_excel(excel_path, index=False)
    except Exception:
        return False
    return True


def _guardar_reporte_pdf_retencion_emitidos_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in EMITIDOS_RETENCION_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = _emitidos_retencion_default_row().get(col, "")
    df = df[EMITIDOS_RETENCION_REPORT_COLUMNS].copy()

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float) and pd.isna(val):
            return ""
        return str(val).strip()

    def _to_number(val):
        if val is None:
            return 0
        if isinstance(val, float):
            if pd.isna(val):
                return 0
            return val
        if isinstance(val, int):
            return val
        parsed = _parse_decimal(str(val))
        return parsed if parsed is not None else 0

    for col in df.columns:
        if col in EMITIDOS_RETENCION_NUMERIC_COLUMNS:
            df[col] = df[col].map(_to_number)
        else:
            df[col] = df[col].map(_to_text)

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="07")
            ws = writer.sheets["07"]
            for idx, column in enumerate(EMITIDOS_RETENCION_REPORT_COLUMNS, start=1):
                if column in EMITIDOS_RETENCION_TEXT_FORCE_COLUMNS:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=idx).number_format = "@"
                else:
                    for row_idx in range(2, ws.max_row + 1):
                        cell = ws.cell(row=row_idx, column=idx)
                        if cell.value == "0":
                            cell.value = 0
                max_len = len(column)
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=idx).value
                    if value is None:
                        continue
                    max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 52)
        return True
    except Exception:
        return False


def _guardar_reporte_pdf_nota_credito_emitidos_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = _nota_credito_emitidos_default_row().get(col, "")
    df = df[EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS].copy()

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float) and pd.isna(val):
            return ""
        return str(val).strip()

    def _to_number(val):
        if val is None:
            return 0
        if isinstance(val, float):
            if pd.isna(val):
                return 0
            return val
        if isinstance(val, int):
            return val
        parsed = _parse_decimal(str(val))
        return parsed if parsed is not None else 0

    for col in df.columns:
        if col in EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS:
            df[col] = df[col].map(_to_number)
        else:
            df[col] = df[col].map(_to_text)

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="04")
            ws = writer.sheets["04"]
            for idx, column in enumerate(EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS, start=1):
                if column in EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=idx).number_format = "@"
                max_len = len(column)
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=idx).value
                    if value is None:
                        continue
                    max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 52)
        return True
    except Exception:
        return False


def _guardar_reporte_pdf_nota_debito_emitidos_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = _nota_debito_emitidos_default_row().get(col, "")
    df = df[EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS].copy()

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float) and pd.isna(val):
            return ""
        return str(val).strip()

    def _to_number(val):
        if val is None:
            return 0
        if isinstance(val, float):
            if pd.isna(val):
                return 0
            return val
        if isinstance(val, int):
            return val
        parsed = _parse_decimal(str(val))
        return parsed if parsed is not None else 0

    for col in df.columns:
        if col in EMITIDOS_NOTA_DEBITO_NUMERIC_COLUMNS:
            df[col] = df[col].map(_to_number)
        else:
            df[col] = df[col].map(_to_text)

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="05")
            ws = writer.sheets["05"]
            for idx, column in enumerate(EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS, start=1):
                if column in EMITIDOS_NOTA_DEBITO_TEXT_FORCE_COLUMNS:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=idx).number_format = "@"
                max_len = len(column)
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=idx).value
                    if value is None:
                        continue
                    max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 52)
        return True
    except Exception:
        return False


def _guardar_reporte_pdf_factura_emitidos_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in EMITIDOS_FACTURA_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = _factura_emitidos_default_row().get(col, "")
    df = df[EMITIDOS_FACTURA_REPORT_COLUMNS].copy()

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float) and pd.isna(val):
            return ""
        return str(val).strip()

    def _to_number(val):
        if val is None:
            return 0
        if isinstance(val, float):
            if pd.isna(val):
                return 0
            return val
        if isinstance(val, int):
            return val
        parsed = _parse_decimal(str(val))
        return parsed if parsed is not None else 0

    for col in df.columns:
        if col in EMITIDOS_FACTURA_NUMERIC_COLUMNS:
            df[col] = df[col].map(_to_number)
        else:
            df[col] = df[col].map(_to_text)

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="01")
            ws = writer.sheets["01"]
            for idx, column in enumerate(EMITIDOS_FACTURA_REPORT_COLUMNS, start=1):
                if column in EMITIDOS_FACTURA_TEXT_FORCE_COLUMNS:
                    for row_idx in range(2, ws.max_row + 1):
                        ws.cell(row=row_idx, column=idx).number_format = "@"
                max_len = len(column)
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=idx).value
                    if value is None:
                        continue
                    max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 52)
        return True
    except Exception:
        return False


def _guardar_reporte_pdf_excel(rows: list[dict], excel_path: Path) -> bool:
    if not rows:
        return False
    df = pd.DataFrame(rows)
    for col in PDF_REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[PDF_REPORT_COLUMNS]
    text_cols = [
        "tipoDocumento",
        "rucEmisor",
        "razonSocialEmisor",
        "nombreComercial",
        "direccionMatrizEmisor",
        "direccionSucursalEmisor",
        "contribuyenteEspecial",
        "agenteRetencion",
        "obligadoContabilidad",
        "tipoContribuyenteRIMPE",
        "numeroComprobante",
        "establecimiento",
        "puntoEmision",
        "secuencial",
        "fechaEmision",
        "fechaAutorizacion",
        "razonSocialComprador",
        "identificacionComprador",
        "direccionComprador",
        "placa",
        "guia",
        "comprobanteModificado",
        "fechaEmisionModificado",
        "razonModificacion",
        "valorModificacion",
        "descripcionesProductos",
        "formaPago",
        "ambiente",
        "emision",
        "claveAcceso",
        "informacionAdicional",
    ]

    def _to_text(val):
        if val is None:
            return ""
        if isinstance(val, float):
            if pd.isna(val):
                return ""
            if val.is_integer():
                return str(int(val))
        return str(val).strip()

    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].map(_to_text)

    numeric_cols = [
        "subtotalTarifaEspecial",
        "subtotal15",
        "subtotal12",
        "subtotal8",
        "subtotal5",
        "subtotal0",
        "subtotalNoObjetoIVA",
        "subtotalExentoIVA",
        "subtotalSinImpuestos",
        "totalDescuento",
        "ivaTarifaEspecial",
        "iva15",
        "iva12",
        "iva8",
        "iva5",
        "ice",
        "irbpnr",
        "propina",
        "valorTotal",
        "valorTotalSinSubsidio",
        "formaPagoMonto",
    ]

    def _to_number(val):
        if val is None:
            return ""
        if isinstance(val, float):
            if pd.isna(val):
                return ""
            return val
        if isinstance(val, int):
            return val
        if isinstance(val, str):
            if not val.strip():
                return ""
            parsed = _parse_decimal(val)
            return parsed if parsed is not None else val
        return val

    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].map(_to_number)
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Detalle PDF", index=False)
            ws = writer.sheets["Detalle PDF"]
            col_index = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
            for columna in ("rucEmisor", "numeroComprobante", "establecimiento", "puntoEmision", "secuencial", "claveAcceso"):
                idx = col_index.get(columna)
                if not idx:
                    continue
                col_letter = get_column_letter(idx)
                for row_idx in range(2, ws.max_row + 1):
                    celda = ws[f"{col_letter}{row_idx}"]
                    celda.number_format = "@"
                    if celda.value is not None:
                        celda.value = str(celda.value)
                    celda.alignment = Alignment(horizontal="left")
    except Exception:
        return False
    return True


def _consolidar_reportes_excel(reportes: list[str], destino: Path) -> Path | None:
    rutas = [Path(p) for p in reportes if p and Path(p).exists()]
    if not rutas:
        return None
    dataframes: list[pd.DataFrame] = []
    columnas: list[str] | None = None
    for ruta in rutas:
        try:
            df = pd.read_excel(ruta)
        except Exception as err:
            print(f"[WARN] No se pudo leer reporte para consolidar: {ruta} ({err})")
            continue
        if df is None or df.empty:
            continue
        if columnas is None:
            columnas = list(df.columns)
        else:
            for col in df.columns:
                if col not in columnas:
                    columnas.append(col)
        dataframes.append(df)
    if not dataframes or not columnas:
        return None
    for idx, df in enumerate(dataframes):
        for col in columnas:
            if col not in df.columns:
                df[col] = ""
        dataframes[idx] = df[columnas]
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        combinado = pd.concat(dataframes, ignore_index=True)
        combinado.to_excel(destino, index=False)
        return destino
    except Exception as err:
        print(f"[WARN] No se pudo escribir reporte consolidado: {destino} ({err})")
        return None


def _collect_existing_reports(base_dir: Path, prefix: str, tipo_slug: str, suffixes) -> list[str]:
    if not base_dir.exists():
        return []
    encontrados: dict[str, Path] = {}
    for suffix in suffixes or []:
        suffix_str = str(suffix or "").strip()
        if not suffix_str:
            continue
        patron = f"{prefix}_{tipo_slug}_{suffix_str}*.xlsx"
        for ruta in sorted(base_dir.glob(patron)):
            if not ruta.is_file():
                continue
            stem = ruta.stem
            esperado = f"{prefix}_{tipo_slug}_{suffix_str}"
            if stem != esperado and not re.fullmatch(rf"{re.escape(esperado)}_\d+", stem):
                continue
            encontrados[str(ruta.resolve())] = ruta
    return [str(ruta) for ruta in sorted(encontrados.values())]


def _delete_report_files(reportes: list[str]) -> None:
    for ruta in reportes or []:
        try:
            Path(ruta).unlink(missing_ok=True)
        except Exception as err:
            print(f"[WARN] No se pudo eliminar reporte intermedio '{ruta}': {err}")

EMITIDOS_FECHA_SELECTORS = [
    "input[id$='fecha_input']",
    "input[name$='fecha_input']",
    "input[id*='fechaEmision']",
    "input[name*='fechaEmision']",
    "input[id*='calFecha']",
]

EMITIDOS_ESTABLECIMIENTO_SELECTORS = [
    "input[id$='establecimiento']",
    "input[name$='establecimiento']",
    "select[id$='establecimiento']",
    "select[name$='establecimiento']",
]

EMITIDOS_PUNTO_SELECTORS = [
    "input[id$='ptoEmision']",
    "input[name$='ptoEmision']",
    "input[id$='punto']",
    "input[name$='punto']",
]



# ====== Funciones auxiliares ======
def _mes_a_texto(mes: int) -> str:
    return ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
            "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][mes-1]


def _xml_files_por_tipo(base_dir: Path, tipo_prefijo: str) -> list[Path]:
    if not base_dir.exists():
        return []
    tipo_prefijo = (tipo_prefijo or "").strip()
    tipo_slug = ""
    if tipo_prefijo:
        try:
            _, etiqueta = tipo_prefijo.split("_", 1)
            tipo_slug = _slug_tipo(etiqueta)
        except Exception:
            tipo_slug = _slug_tipo(tipo_prefijo)
    encontrados: list[Path] = []
    for ruta in base_dir.rglob("*.xml"):
        try:
            nombre_norm = ruta.name.lower()
            if (
                not tipo_prefijo
                or tipo_prefijo in ruta.parts
                or (tipo_slug and nombre_norm.startswith(f"{tipo_slug}__"))
            ):
                encontrados.append(ruta)
        except Exception:
            continue
    return sorted(encontrados)


def _xml_files_por_meses(base_dir: Path, tipo_prefijo: str, meses) -> list[Path]:
    encontrados: list[Path] = []
    for mes in meses:
        try:
            mes_int = int(mes)
        except Exception:
            continue
        try:
            mes_dir = base_dir / _mes_a_texto(mes_int)
        except Exception:
            continue
        encontrados.extend(_xml_files_por_tipo(mes_dir, tipo_prefijo))
    vistos = set()
    normalizados: list[Path] = []
    for ruta in encontrados:
        if ruta in vistos:
            continue
        vistos.add(ruta)
        normalizados.append(ruta)
    return sorted(normalizados)

def _es_clave(valor: str) -> bool:
    return bool(re.fullmatch(r"\d{49}", (valor or "").strip()))

def _detectar_delimitador(sample: str) -> str:
    counts = { ';': sample.count(';'), ',': sample.count(','), '\t': sample.count('\t') }
    return max(counts, key=counts.get) if any(counts.values()) else ';'

def _extraer_claves_desde_txt(txt_path: Path):
    claves = []
    sample = txt_path.read_text(encoding="utf-8", errors="ignore")[:4096]
    sep = _detectar_delimitador(sample)
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f, delimiter=sep)
        for row in reader:
            if not row: 
                continue
            clave = next((c.strip() for c in row if _es_clave(c)), None)
            if not clave:
                continue
            tipo = next((c.strip() for c in row if c.lower().startswith(("factura","comprobante","nota","liquidacion"))), "")
            fecha = next((c.strip() for c in row if re.fullmatch(r"\d{2}/\d{2}/\d{4}", c.strip())), "")
            claves.append({"clave": clave, "tipo": tipo, "fecha": fecha})
    return claves

def _sanear_nombre_archivo(texto: str, sufijo: str = "") -> str:
    base = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_")
    if not base:
        base = "documento"
    if sufijo:
        base = f"{base}_{sufijo}"
    return base


def _nombre_carpeta_tipo(tipo: str) -> str:
    base = unicodedata.normalize("NFKD", (tipo or "")).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_")
    return base or "Otros"


def _slug_tipo(tipo: str) -> str:
    return _nombre_carpeta_tipo(tipo).lower()

def _normalizar_label_simple(texto_label: str) -> str:
    base = unicodedata.normalize("NFKD", texto_label or "").encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9 ]+", " ", base).upper()
    base = re.sub(r"\s+", " ", base).strip()
    return base

def _es_tipo_retencion(tipo: str) -> bool:
    slug = _slug_tipo(tipo or "")
    return slug in {
        "retencion",
        "retenciones",
        "comprobante_de_retencion",
        "comprobantes_de_retencion",
        "comprobante_de_retencion_venta",
    }

def _es_tipo_nota_credito(tipo: str) -> bool:
    slug = _slug_tipo(tipo or "")
    return slug in {
        "nota_de_credito",
        "notas_de_credito",
        "nota_credito",
        "notas_credito",
    }

def _es_tipo_nota_debito(tipo: str) -> bool:
    slug = _slug_tipo(tipo or "")
    return slug in {
        "nota_de_debito",
        "notas_de_debito",
        "nota_debito",
        "notas_debito",
    }


def _es_tipo_factura(tipo: str) -> bool:
    slug = _slug_tipo(tipo or "")
    return slug in {"factura", "facturas"}


def _es_tipo_liquidacion_compra(tipo: str) -> bool:
    slug = _slug_tipo(tipo or "")
    return slug in {
        "liquidacion_de_compra",
        "liquidacion_de_compra_de_bienes_y_prestacion_de_servicios",
    }


TIPO_LABEL_MAP = {
    "factura": (1, "Factura"),
    "facturas": (1, "Factura"),
    "liquidacion_de_compra": (2, "Liquidacion_de_Compra"),
    "liquidacion_de_compra_de_bienes_y_prestacion_de_servicios": (2, "Liquidacion_de_Compra"),
    "comprobante_de_retencion": (6, "Retencion"),
    "comprobantes_de_retencion": (6, "Retencion"),
    "retencion": (6, "Retencion"),
    "retenciones": (6, "Retencion"),
    "nota_de_credito": (3, "NotaCredito"),
    "notas_de_credito": (3, "NotaCredito"),
    "nota_credito": (3, "NotaCredito"),
    "notas_credito": (3, "NotaCredito"),
    "nota_de_debito": (4, "NotaDebito"),
    "notas_de_debito": (4, "NotaDebito"),
    "nota_debito": (4, "NotaDebito"),
    "notas_debito": (4, "NotaDebito"),
    "guia_de_remision": (5, "GuiaRemision"),
    "guias_de_remision": (5, "GuiaRemision"),
    "guia_remision": (5, "GuiaRemision"),
    "comprobante_de_retencion_venta": (6, "Retencion"),
}


def _normalizar_tipo_clave(texto: str) -> str:
    base = unicodedata.normalize("NFKD", (texto or "")).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_").lower()
    return base


def _formatear_label(texto: str) -> str:
    if not texto:
        return "Documentos"
    partes = [fragment.capitalize() for fragment in texto.split("_") if fragment]
    return "_".join(partes) if partes else (texto or "Documentos")


def _nombre_carpeta_tipo_visible(tipo_texto: str) -> str:
    clave = _normalizar_tipo_clave(tipo_texto)
    if clave in {"retencion", "retenciones", "comprobante_de_retencion", "comprobantes_de_retencion", "comprobante_de_retencion_venta"}:
        return "Comprobante de Retencion"
    if clave in {"factura", "facturas"}:
        return "Factura"
    if clave in {"liquidacion_de_compra", "liquidacion_de_compra_de_bienes_y_prestacion_de_servicios"}:
        return "Liquidacion de Compra"
    if clave in {"nota_de_credito", "notas_de_credito", "nota_credito", "notas_credito"}:
        return "Nota de Credito"
    if clave in {"nota_de_debito", "notas_de_debito", "nota_debito", "notas_debito"}:
        return "Nota de Debito"
    if clave in {"guia_de_remision", "guias_de_remision", "guia_remision"}:
        return "Guia de Remision"
    return _formatear_label(_nombre_carpeta_tipo(tipo_texto)).replace("_", " ")


def _resolver_tipo_label(tipo_texto: str) -> tuple[int, str]:
    clave = _normalizar_tipo_clave(tipo_texto)
    if clave in TIPO_LABEL_MAP:
        return TIPO_LABEL_MAP[clave]
    if clave.endswith("s"):
        clave_singular = clave.rstrip("s")
        if clave_singular in TIPO_LABEL_MAP:
            return TIPO_LABEL_MAP[clave_singular]
    label_sanitizado = _nombre_carpeta_tipo(tipo_texto)
    label = _formatear_label(label_sanitizado)
    return 99, label


def _coincide_tipo_documental(tipo_esperado: str, tipo_detectado: str) -> bool:
    if not tipo_esperado or not tipo_detectado:
        return True
    orden_esperado, _ = _resolver_tipo_label(tipo_esperado)
    orden_detectado, _ = _resolver_tipo_label(tipo_detectado)
    if orden_esperado == 99 or orden_detectado == 99:
        return _slug_tipo(tipo_esperado) == _slug_tipo(tipo_detectado)
    return orden_esperado == orden_detectado


def _prefijo_tipo(tipo_texto: str) -> tuple[int, str, str]:
    orden, etiqueta = _resolver_tipo_label(tipo_texto)
    prefijo = f"{orden:02d}_{etiqueta}"
    return orden, etiqueta, prefijo


def _nombre_documento_mes(tipo_slug: str, fecha_token: str, nombre_base: str) -> str:
    tipo_parte = _slug_tipo(tipo_slug or "") or "documento"
    fecha_parte = re.sub(r"[^0-9]", "", fecha_token or "") or "00000000"
    base = _sanear_nombre_archivo(nombre_base or "archivo")
    return f"{tipo_parte}__{fecha_parte}__{base}"

def _portal_indisponible(page) -> bool:
    try:
        contenido = page.content()
    except Exception:
        return False
    if not contenido:
        return False
    texto = unicodedata.normalize("NFKD", contenido).lower()
    return "ha ocurrido un error" in texto and "indisponibil" in texto


def _asegurar_portal_disponible(page):
    if _portal_indisponible(page):
        raise RuntimeError(PORTAL_INDISPONIBLE_MENSAJE)


def _obtener_view_state(page) -> str:
    try:
        view_input = page.locator("input[name='javax.faces.ViewState']")
        if view_input.count():
            return (view_input.first.input_value() or "").strip()
    except Exception:
        pass
    return ""


def _actualizar_view_state_input(page, nuevo_view_state: str):
    if not nuevo_view_state:
        return
    try:
        page.locator("input[name='javax.faces.ViewState']").evaluate(
            "(el, val) => { el.value = val; }",
            nuevo_view_state,
        )
    except Exception:
        pass


def _notificar_usuario_captcha(tipo: str, contexto: str):
    mensaje = (
        f"[ACCION] Se detecto un {tipo} en '{contexto}'. "
        'Resuelvelo manualmente en la ventana del navegador y luego continua.'
    )
    print(mensaje)
    if USER_NOTIFICATION_CALLBACK:
        try:
            USER_NOTIFICATION_CALLBACK(mensaje)
        except Exception as err:
            print(f"[WARN] No se pudo enviar notificacion al UI: {err}")


def _notificar_usuario_accion(mensaje: str):
    mensaje = (mensaje or "").strip()
    if not mensaje:
        return
    print(mensaje)
    if USER_NOTIFICATION_CALLBACK:
        try:
            USER_NOTIFICATION_CALLBACK(mensaje)
        except Exception as err:
            print(f"[WARN] No se pudo enviar notificacion al UI: {err}")


def _obtener_form_base_emitidos(page):
    try:
        datos = page.evaluate(
            """() => {
                const form = document.querySelector('form#frmPrincipal');
                if (!form) { return {}; }
                const payload = {};
                const elementos = form.querySelectorAll('input, select, textarea');
                for (const el of elementos) {
                    if (!el.name) { continue; }
                    if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) {
                        continue;
                    }
                    payload[el.name] = el.value ?? '';
                }
                return payload;
            }"""
        ) or {}
    except Exception:
        datos = {}
    datos.pop("javax.faces.ViewState", None)
    return datos


def _extraer_autorizacion_desde_partial(respuesta: str):
    if not respuesta:
        return None, None
    autorizacion_bruta = None
    cdata_matches = re.findall(r"<!\[CDATA\[(.*?)\]\]>", respuesta, flags=re.DOTALL)
    for chunk in cdata_matches:
        texto = chunk.strip()
        if not texto:
            continue
        if "&lt;autorizacion" in texto and "<autorizacion" not in texto:
            texto = html.unescape(texto)
        if "<autorizacion" in texto:
            autorizacion_bruta = texto
            break
    if autorizacion_bruta is None:
        desen = html.unescape(respuesta)
        match = re.search(r"(<autorizacion.*?</autorizacion>)", desen, flags=re.DOTALL | re.IGNORECASE)
        if match:
            autorizacion_bruta = match.group(1)
    view_state_match = re.search(
        r'<update id="javax\.faces\.ViewState"><!\[CDATA\[(.*?)\]\]></update>',
        respuesta,
        flags=re.DOTALL,
    )
    nuevo_view_state = view_state_match.group(1).strip() if view_state_match else None
    return autorizacion_bruta, nuevo_view_state


def _strip_xml_namespaces(element: ET.Element):
    if element is None:
        return
    for node in element.iter():
        if '}' in node.tag:
            node.tag = node.tag.split('}', 1)[1]
        if ':' in node.tag:
            node.tag = node.tag.split(':', 1)[1]
        if node.attrib:
            node.attrib = {
                key.split('}', 1)[-1].split(':', 1)[-1]: val
                for key, val in node.attrib.items()
            }


def _limpiar_cdata(texto: str) -> str:
    if not texto:
        return ""
    contenido = texto.strip()
    if contenido.startswith("<![CDATA[") and contenido.endswith("]]>"):
        contenido = contenido[9:-3]
    return contenido.strip()


def _buscar_autorizacion_en_json(payload):
    if isinstance(payload, dict):
        for val in payload.values():
            encontrado = _buscar_autorizacion_en_json(val)
            if encontrado:
                return encontrado
    elif isinstance(payload, list):
        for item in payload:
            encontrado = _buscar_autorizacion_en_json(item)
            if encontrado:
                return encontrado
    elif isinstance(payload, str) and "<autorizacion" in payload:
        return payload
    return None


def _es_url_autorizacion(url: str) -> bool:
    if not url:
        return False
    url_lower = url.lower()
    if "sri.gob.ec" not in url_lower:
        return False
    return "autoriz" in url_lower


def _extraer_comprobante_desde_autorizacion(payload: str):
    if not payload:
        raise ValueError("Respuesta de autorizacion vacia.")
    bruto = payload.strip()
    if bruto.startswith("{"):
        try:
            data = json.loads(bruto)
            posible = _buscar_autorizacion_en_json(data)
            if isinstance(posible, str) and "<autorizacion" in posible:
                bruto = posible
        except json.JSONDecodeError:
            pass
    bruto = html.unescape(bruto)
    try:
        root = ET.fromstring(bruto)
    except ET.ParseError as err:
        match = re.search(r"(<autorizaciones?.*?</autorizaciones?>)", bruto, re.DOTALL | re.IGNORECASE)
        if match:
            try:
                root = ET.fromstring(match.group(1))
            except ET.ParseError as inner_err:
                raise ValueError("No se pudo interpretar la respuesta de autorizacion.") from inner_err
        else:
            contenido_match = re.search(r"<comprobante[^>]*>(.*?)</comprobante>", bruto, re.DOTALL | re.IGNORECASE)
            if contenido_match:
                return _limpiar_cdata(contenido_match.group(1)), {}
            raise ValueError("No se pudo interpretar la respuesta de autorizacion.") from err
    _strip_xml_namespaces(root)
    candidatos = []
    if root.tag.lower().endswith("autorizacion"):
        candidatos = [root]
    else:
        candidatos = list(root.findall(".//autorizacion"))
    for autorizacion in candidatos:
        comprobante_texto = autorizacion.findtext("comprobante", "")
        if not comprobante_texto:
            continue
        meta = {
            "estado": (autorizacion.findtext("estado") or "").strip(),
            "numero_autorizacion": (autorizacion.findtext("numeroAutorizacion") or "").strip(),
            "fecha_autorizacion": (autorizacion.findtext("fechaAutorizacion") or "").strip(),
            "ambiente": (autorizacion.findtext("ambiente") or "").strip(),
        }
        return _limpiar_cdata(comprobante_texto), meta
    raise ValueError("La respuesta de autorizacion no contiene un comprobante valido.")


def _primer_texto(node: ET.Element, nombres):
    if node is None:
        return ""
    for nombre in nombres:
        valor = node.findtext(nombre)
        if valor and valor.strip():
            return valor.strip()
    return ""


def _fecha_slug(fecha: str) -> str:
    if not fecha:
        return ""
    texto = fecha.strip()
    if not texto:
        return ""
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", texto)
    if m:
        return f"{m.group(3)}{m.group(2)}{m.group(1)}"
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"
    m = re.search(r"(\d{4})(\d{2})(\d{2})", texto)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"
    solo_digitos = re.sub(r"[^\d]", "", texto)
    return solo_digitos[:8]


def _total_slug(valor) -> str:
    if isinstance(valor, (int, float)):
        numero = float(valor)
    else:
        numero = _parse_decimal(valor) if isinstance(valor, str) else None
        if numero is None:
            return ""
    return f"{numero:.2f}".replace(".", "_")


def _parse_emitido_comprobante(xml_texto: str, meta_autorizacion: Optional[dict] = None):
    contenido = _limpiar_cdata(xml_texto)
    if not contenido:
        raise ValueError("El comprobante emitido llego vacio.")
    try:
        root = ET.fromstring(contenido)
    except ET.ParseError as err:
        raise ValueError(f"XML de comprobante emitido invalido: {err}") from err
    _strip_xml_namespaces(root)
    info_trib = root.find("infoTributaria")
    if info_trib is None:
        raise ValueError("El comprobante emitido no contiene infoTributaria.")
    cod_doc = (info_trib.findtext("codDoc") or "").strip()
    estab = (info_trib.findtext("estab") or "").strip()
    pto_emi = (info_trib.findtext("ptoEmi") or "").strip()
    secuencial = (info_trib.findtext("secuencial") or "").strip()
    clave = (info_trib.findtext("claveAcceso") or "").strip()

    info_node = None
    for candidato in (
        "infoFactura",
        "infoLiquidacionCompra",
        "infoNotaCredito",
        "infoNotaDebito",
        "infoCompRetencion",
        "infoGuiaRemision",
    ):
        info_node = root.find(candidato)
        if info_node is not None:
            break

    fecha_emision = ""
    identificacion = ""
    razon_receptor = ""
    importe_total = ""
    if info_node is not None:
        fecha_emision = (info_node.findtext("fechaEmision") or "").strip()
        identificacion = _primer_texto(
            info_node,
            [
                "identificacionComprador",
                "identificacionSujetoRetenido",
                "identificacionBeneficiario",
                "identificacionCliente",
                "identificacionDestinatario",
            ],
        )
        razon_receptor = _primer_texto(
            info_node,
            [
                "razonSocialComprador",
                "razonSocialSujetoRetenido",
                "razonSocialBeneficiario",
                "razonSocialCliente",
                "razonSocialDestinatario",
                "nombreComercialComprador",
                "nombreComprador",
            ],
        )
        importe_total = _primer_texto(
            info_node,
            [
                "importeTotal",
                "valorTotal",
                "total",
                "totalPagar",
                "valorModificacion",
                "valorRetenido",
            ],
        )
        if not importe_total:
            importe_total = _primer_texto(info_node, ["totalSinImpuestos"])

    meta = {
        "cod_doc": cod_doc,
        "doc_label": DOC_LABELS.get(cod_doc, cod_doc or "Documento"),
        "estab": estab,
        "pto_emi": pto_emi,
        "secuencial": secuencial,
        "clave_acceso": clave,
        "identificacion_receptor": identificacion,
        "razon_social_receptor": razon_receptor,
        "fecha_emision": fecha_emision,
        "importe_total": importe_total,
        "xml_contenido": contenido,
    }
    if meta_autorizacion:
        meta["estado_autorizacion"] = meta_autorizacion.get("estado", "")
        meta["numero_autorizacion"] = meta_autorizacion.get("numero_autorizacion", "")
        meta["fecha_autorizacion"] = meta_autorizacion.get("fecha_autorizacion", "")
        meta["ambiente"] = meta_autorizacion.get("ambiente", "")
    return meta


def _construir_nombre_xml_emitido(meta: dict, fallback: str) -> str:
    serie = "-".join(
        parte for parte in (meta.get("estab"), meta.get("pto_emi"), meta.get("secuencial")) if parte
    )
    partes = []
    doc_label = meta.get("doc_label")
    if doc_label:
        partes.append(doc_label)
    if serie:
        partes.append(serie)
    identificacion = meta.get("identificacion_receptor")
    if identificacion:
        partes.append(identificacion)
    fecha_token = _fecha_slug(meta.get("fecha_emision", ""))
    if fecha_token:
        partes.append(fecha_token)
    total_token = _total_slug(meta.get("importe_total", ""))
    if total_token:
        partes.append(f"total-{total_token}")
    if not partes:
        fallback_nombre = fallback or meta.get("clave_acceso") or meta.get("secuencial") or "emitido"
        partes.append(fallback_nombre)
    nombre = _sanear_nombre_archivo("_".join(partes))
    if len(nombre) > 180:
        nombre = nombre[:180].rstrip("._-")
    return nombre or "emitido"


def _capturar_xml_emitido(
    page,
    request_context,
    source_id: str,
    xml_dir: Path,
    fallback_nombre: str,
    claves_guardadas: set[str],
    payload_base: dict,
    view_state: str,
    headers: dict,
):
    if not source_id:
        return None, view_state
    dialog_error = None
    try:
        resultado_dialogo, view_state_dialogo = _capturar_xml_emitido_por_dialogo(
            page,
            request_context,
            source_id,
            xml_dir,
            fallback_nombre,
            claves_guardadas,
            headers,
        )
        return resultado_dialogo, view_state_dialogo
    except Exception as err:
        dialog_error = err
    payload = dict(payload_base or {})
    payload.update(
        {
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": source_id,
            "javax.faces.partial.execute": source_id,
            "javax.faces.partial.render": "form-detalle-factura:panel-detalle-factura",
            source_id: source_id,
            "javax.faces.ViewState": view_state or "",
        }
    )
    payload.setdefault("frmPrincipal", "frmPrincipal")

    respuesta = request_context.post(
        RECUPERAR_COMPROBANTES_URL,
        data=payload,
        headers=headers,
    )
    if respuesta.status != 200:
        raise RuntimeError(f"Error HTTP {respuesta.status} al recuperar comprobante.")
    cuerpo = respuesta.text()
    autorizacion_bruta, nuevo_view_state = _extraer_autorizacion_desde_partial(cuerpo)
    if not autorizacion_bruta:
        if dialog_error:
            raise RuntimeError("No se pudo extraer la autorizacion del comprobante.") from dialog_error
        raise RuntimeError("No se pudo extraer la autorizacion del comprobante.")

    comprobante_xml, meta_aut = _extraer_comprobante_desde_autorizacion(autorizacion_bruta)
    meta = _parse_emitido_comprobante(comprobante_xml, meta_aut)
    clave = meta.get("clave_acceso")
    if clave and clave in claves_guardadas:
        return None, nuevo_view_state or view_state

    nombre_archivo = _construir_nombre_xml_emitido(meta, fallback_nombre)
    destino = xml_dir / f"{nombre_archivo}.xml"
    sufijo = 1
    while destino.exists():
        destino = xml_dir / f"{nombre_archivo}_{sufijo}.xml"
        sufijo += 1

    contenido = meta.get("xml_contenido", "").strip()
    if not contenido:
        raise ValueError("El comprobante retornado esta vacio.")
    if not contenido.startswith("<?xml"):
        contenido = '<?xml version="1.0" encoding="UTF-8"?>\n' + contenido
    destino.write_text(contenido, encoding="utf-8")
    if clave:
        claves_guardadas.add(clave)

    if nuevo_view_state:
        _actualizar_view_state_input(page, nuevo_view_state)
    return destino, nuevo_view_state or view_state




def _capturar_xml_emitido_por_dialogo(
    page,
    request_context,
    source_id: str,
    xml_dir: Path,
    fallback_nombre: str,
    claves_guardadas: set[str],
    base_headers: Optional[dict] = None,
) -> tuple[Optional[Path], str]:
    if not source_id:
        return None, _obtener_view_state(page)

    dialog_actual = None
    resultado_path: Optional[Path] = None

    try:
        disparador = None
        candidatos = [
            f"[id='{source_id}']",
            f"a[id='{source_id}']",
            f"button[id='{source_id}']",
        ]
        sufijo = source_id.split(":")[-1] if ":" in source_id else source_id
        candidatos.extend(
            [
                f"a[onclick*='{sufijo}']",
                f"button[onclick*='{sufijo}']",
            ]
        )
        for selector in candidatos:
            try:
                loc = page.locator(selector)
            except Exception:
                continue
            if loc.count():
                disparador = loc.first
                break
        if disparador is None:
            return None, _obtener_view_state(page)

        try:
            disparador.scroll_into_view_if_needed(timeout=500)
        except Exception:
            pass
        try:
            disparador.click(timeout=1000)
        except Exception as err:
            raise RuntimeError(f"No se pudo abrir el detalle emitido ({source_id}).") from err

        dialog_locator = page.locator("div.ui-dialog:has(span.ui-dialog-title:has-text('XML'))")
        if not dialog_locator.count():
            dialog_locator = page.locator("div.ui-dialog:has(span.ui-dialog-title)")
        try:
            dialog_locator.evaluate_all(
                "(els) => els.forEach(el => { el.classList.remove('ui-overlay-hidden'); el.style.visibility = 'visible'; })"
            )
        except Exception:
            pass
        try:
            dialog_locator.first.wait_for(state="visible", timeout=DOWNLOAD_TIMEOUT)
        except PlaywrightTimeoutError as err:
            print(f"[WARN] No aparecio el dialogo XML para {source_id}: {err}")
            raise RuntimeError("No aparecio el dialogo de descarga de XML.") from err
        dialog_actual = dialog_locator.first

        form_payload = page.evaluate(
            """() => {
                const form = document.querySelector('form#j_idt913');
                if (!form) { return null; }
                const data = new FormData(form);
                const payload = {};
                for (const [key, value] of data.entries()) {
                    payload[key] = value;
                }
                return payload;
            }"""
        )
        if not form_payload:
            print(f"[WARN] No se encontro el formulario j_idt913 en el dialogo para {source_id}.")
            raise RuntimeError("No se pudo obtener el formulario de descarga de XML.")

        encoded_payload = urlencode(form_payload, doseq=True)
        request_headers = dict(base_headers or {})
        request_headers.update(
            {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Faces-Request": "partial/ajax",
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/xml, text/xml, */*; q=0.01",
                "Origin": "https://srienlinea.sri.gob.ec",
                "Referer": page.url,
            }
        )

        try:
            respuesta = request_context.post(
            RECUPERAR_COMPROBANTES_URL,
            data=encoded_payload,
            headers=request_headers,
        )
        except Exception as err:
            print(f"[WARN] Fallo la solicitud POST de XML para {source_id}: {err}")
            raise
        if respuesta.status != 200:
            print(f"[WARN] HTTP {respuesta.status} al solicitar XML de emitidos para {source_id}.")
            raise RuntimeError(f"Error HTTP {respuesta.status} al descargar XML de emitidos.")
        try:
            cuerpo_bytes = respuesta.body()
        except Exception as err:
            print(f"[WARN] No se pudo leer cuerpo de respuesta XML para {source_id}: {err}")
            raise RuntimeError("No se pudo leer la respuesta del XML de emitidos.") from err
        if not cuerpo_bytes:
            print(f"[WARN] Respuesta vacia al descargar XML para {source_id}.")
            raise RuntimeError("La respuesta del XML de emitidos llego vacia.")
        try:
            contenido = cuerpo_bytes.decode("utf-8")
        except UnicodeDecodeError:
            contenido = cuerpo_bytes.decode("utf-8", "ignore")
        contenido = contenido.strip()
        if not contenido:
            raise ValueError("El XML descargado esta vacio.")


        try:
            meta = _parse_emitido_comprobante(contenido, None)
        except Exception as err:
            print(f"[WARN] No se pudo interpretar el XML de emitidos descargado: {err}")
            meta = {"xml_contenido": contenido}
        if not isinstance(meta, dict):
            meta = {"xml_contenido": contenido}
        meta.setdefault("xml_contenido", contenido)

        clave = meta.get("clave_acceso")
        if clave and clave in claves_guardadas:
            resultado_path = None
        else:
            contenido_xml = meta.get("xml_contenido") or contenido
            if not contenido_xml.lstrip().startswith("<?xml"):
                contenido_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' + contenido_xml.lstrip()

            nombre_archivo = _construir_nombre_xml_emitido(meta, fallback_nombre)
            destino_base = xml_dir / nombre_archivo
            destino_final = _resolver_destino_unico(destino_base, ".xml")
            destino_final.write_text(contenido_xml, encoding="utf-8")
            resultado_path = destino_final

            if clave:
                claves_guardadas.add(clave)

    finally:
        if dialog_actual:
            try:
                cierre = dialog_actual.locator("a.ui-dialog-titlebar-close")
                if cierre.count():
                    cierre.first.click()
                else:
                    cierre_btn = dialog_actual.locator("button[title*='Cerrar' i], button:has-text('Cerrar')")
                    if cierre_btn.count():
                        cierre_btn.first.click()
            except Exception:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass

    nuevo_view_state = _obtener_view_state(page)
    return resultado_path, nuevo_view_state


SOAP_ENVELOPE_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:ec="http://ec.gob.sri.ws.autorizacion">
  <soapenv:Header/>
  <soapenv:Body>
    <ec:autorizacionComprobante>
      <claveAccesoComprobante>{clave}</claveAccesoComprobante>
    </ec:autorizacionComprobante>
  </soapenv:Body>
</soapenv:Envelope>
"""


def _descargar_xml_emitido_por_clave(
    request_context,
    clave_acceso: str,
    xml_dir: Path,
    fallback_nombre: str,
    claves_guardadas: set[str],
) -> Optional[Path]:
    clave = (clave_acceso or "").strip()
    if not clave:
        raise ValueError("Clave de acceso vacia para la descarga SOAP.")

    envelope = SOAP_ENVELOPE_TEMPLATE.format(clave=clave)
    headers = {
        "Content-Type": "text/xml; charset=UTF-8",
        "SOAPAction": "",
        "Accept": "text/xml",
    }
    respuesta = request_context.post(
        AUTORIZACION_COMPROBANTES_SOAP_URL,
        data=envelope.encode("utf-8"),
        headers=headers,
    )
    if respuesta.status != 200:
        raise RuntimeError(f"Servicio SOAP respondio HTTP {respuesta.status}.")
    cuerpo = respuesta.text()
    if not cuerpo:
        raise RuntimeError("El servicio SOAP devolvio respuesta vacia.")

    match = re.search(r"(<autorizacion[\s\S]*?</autorizacion>)", cuerpo, flags=re.IGNORECASE)
    if not match:
        raise RuntimeError("El servicio SOAP no devolvio un bloque <autorizacion>.")
    autorizacion_xml = html.unescape(match.group(1))

    comprobante_xml, meta_aut = _extraer_comprobante_desde_autorizacion(autorizacion_xml)
    if not comprobante_xml or not comprobante_xml.strip():
        estado = (meta_aut or {}).get("estado", "desconocido") if meta_aut else "desconocido"
        raise RuntimeError(f"El servicio SOAP retorno estado '{estado}' sin comprobante.")

    meta = _parse_emitido_comprobante(comprobante_xml, meta_aut)
    meta.setdefault("xml_contenido", comprobante_xml)

    clave_meta = (meta.get("clave_acceso") or clave).strip()
    if clave_meta and clave_meta in claves_guardadas:
        return None

    contenido_xml = meta.get("xml_contenido") or comprobante_xml
    if not contenido_xml.lstrip().startswith("<?xml"):
        contenido_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' + contenido_xml.lstrip()

    nombre_archivo = _construir_nombre_xml_emitido(meta, fallback_nombre)
    destino_base = xml_dir / nombre_archivo
    destino_final = _resolver_destino_unico(destino_base, ".xml")
    destino_final.write_text(contenido_xml, encoding="utf-8")

    if clave_meta:
        claves_guardadas.add(clave_meta)
    return destino_final


def _parse_decimal(texto: str) -> Optional[float]:
    bruto = (texto or "").strip().replace("\xa0", "").replace(" ", "")
    if not bruto:
        return None
    candidatos = [
        bruto,
        bruto.replace(".", "").replace(",", "."),
        bruto.replace(",", ""),
    ]
    for candidato in candidatos:
        try:
            return float(candidato)
        except ValueError:
            continue
    return None


def _parse_datetime_local(texto: str) -> Optional[datetime]:
    bruto = (texto or "").strip()
    if not bruto:
        return None
    bruto = re.sub(r"\s+", " ", bruto)
    formatos = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y")
    for fmt in formatos:
        try:
            return datetime.strptime(bruto, fmt)
        except ValueError:
            continue
    return None


def _debe_omitir_soap_xml(fecha_emision: str, descargar_xml: bool, dias_limite: int = 30) -> bool:
    if descargar_xml:
        return False
    if not fecha_emision:
        return False
    fecha_dt = _parse_datetime_local(fecha_emision)
    if not fecha_dt:
        return False
    try:
        limite = int(dias_limite)
    except Exception:
        limite = 30
    dias = (datetime.now().date() - fecha_dt.date()).days
    return dias > limite


def _inferir_iva_columna(iva_val: float | None, base_val: float | None) -> str | None:
    if not iva_val or not base_val:
        return None
    if base_val <= 0:
        return None
    rate = iva_val / base_val
    candidatos = [
        (0.15, "iva15"),
        (0.12, "iva12"),
        (0.08, "iva8"),
        (0.05, "iva5"),
    ]
    for esperado, col in candidatos:
        if abs(rate - esperado) <= 0.01:
            return col
    return None


def _valor_reporte_presente(valor) -> bool:
    if valor is None:
        return False
    if isinstance(valor, (int, float)):
        return True
    texto = str(valor).strip()
    if not texto:
        return False
    token = _normalizar_token(texto)
    return token not in {"nodisponible", "na", "n/a", "none", "null", "sindato", "nohaydato"}


def _texto_probable_comprador_desde_fila(texto: str) -> str:
    valor = (texto or "").strip()
    if not valor:
        return ""
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2}:\d{2})?", valor):
        return ""
    if re.fullmatch(r"[\d.,/-]+", valor):
        return ""
    if not re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", valor):
        return ""
    return valor


def _combinar_datos_reporte_emitidos(*fuentes: dict | None) -> dict:
    datos = {col: "" for col in PDF_REPORT_COLUMNS}
    for fuente in fuentes:
        if not isinstance(fuente, dict):
            continue
        for col in PDF_REPORT_COLUMNS:
            nuevo = fuente.get(col)
            if not _valor_reporte_presente(nuevo):
                continue
            if not _valor_reporte_presente(datos.get(col)):
                datos[col] = nuevo
    return datos


def _extraer_datos_pdf_por_tipo_layout_first(
    pdf_path: Path,
    *,
    es_retencion: bool = False,
    es_nota_credito: bool = False,
    es_nota_debito: bool = False,
) -> dict:
    layout_data = {}
    if _extract_pdf_layout_fields is not None and not es_retencion:
        try:
            layout_data = _extract_pdf_layout_fields(pdf_path) or {}
        except Exception as err:
            print(f"[WARN] No se pudo extraer por layout visual el PDF '{pdf_path.name}': {err}")

    if es_retencion:
        legacy_data = _extraer_datos_pdf_retencion(pdf_path)
    elif es_nota_credito:
        legacy_data = _extraer_datos_pdf_nota_credito(pdf_path)
    elif es_nota_debito:
        legacy_data = _extraer_datos_pdf_nota_debito(pdf_path)
    else:
        legacy_data = _extraer_datos_pdf(pdf_path)
    if not layout_data:
        return legacy_data
    return _combinar_datos_reporte_emitidos(layout_data, legacy_data)


def _extraer_datos_emitidos_dom(
    tipo_visible: str,
    tipo_serie_texto: str,
    clave_texto: str,
    fecha_emision: str,
    fecha_autorizacion: str,
    razon_texto: str,
    valor_sin_impuestos: str,
    iva_valor: str,
    importe_total: str,
    ruc_emisor: str | None = None,
) -> dict:
    datos = {col: "" for col in PDF_REPORT_COLUMNS}
    tipo_val = (tipo_visible or "").strip()
    if not tipo_val and tipo_serie_texto:
        tipo_val = tipo_serie_texto.split()[0].strip()
    datos["tipoDocumento"] = tipo_val
    datos["razonSocialComprador"] = _texto_probable_comprador_desde_fila(razon_texto)
    datos["fechaEmision"] = (fecha_emision or "").strip()
    datos["fechaAutorizacion"] = (fecha_autorizacion or "").strip()
    if ruc_emisor:
        datos["rucEmisor"] = ruc_emisor

    numero = ""
    if tipo_serie_texto:
        match = re.search(r"\d{3}-\d{3}-\d{9}", tipo_serie_texto)
        if match:
            numero = match.group(0)
    if numero:
        datos["numeroComprobante"] = numero
        partes = numero.split("-")
        if len(partes) == 3:
            datos["establecimiento"] = partes[0]
            datos["puntoEmision"] = partes[1]
            datos["secuencial"] = partes[2]

    if clave_texto:
        datos["claveAcceso"] = clave_texto

    base_val = _parse_decimal(valor_sin_impuestos) if isinstance(valor_sin_impuestos, str) else None
    iva_val = _parse_decimal(iva_valor) if isinstance(iva_valor, str) else None
    total_val = _parse_decimal(importe_total) if isinstance(importe_total, str) else None

    if base_val is not None:
        datos["subtotalSinImpuestos"] = base_val
    if total_val is not None:
        datos["valorTotal"] = total_val

    iva_col = _inferir_iva_columna(iva_val, base_val)
    if iva_col and iva_val is not None:
        datos[iva_col] = iva_val

    return datos


def _strip_html(texto: str) -> str:
    if not texto:
        return ""
    sin_tags = re.sub(r"<[^>]+>", " ", texto)
    return re.sub(r"\s+", " ", sin_tags).strip()


def _extraer_detalle_emitido_desde_partial(respuesta: str) -> dict:
    if not respuesta:
        return {}
    chunks = re.findall(r"<!\[CDATA\[(.*?)\]\]>", respuesta, flags=re.DOTALL)
    if not chunks:
        chunks = [respuesta]
    html_total = html.unescape(" ".join(chunks))
    patron = re.compile(
        r"<td[^>]*formulario-label[^>]*>.*?<label[^>]*>(.*?)</label>.*?</td>\s*"
        r"<td[^>]*middle[^>]*>(.*?)</td>",
        re.IGNORECASE | re.DOTALL,
    )
    kv: dict[str, str] = {}
    for etiqueta_html, valor_html in patron.findall(html_total):
        etiqueta = _strip_html(etiqueta_html)
        valor = _strip_html(valor_html)
        if not etiqueta:
            continue
        if etiqueta not in kv or not kv.get(etiqueta):
            kv[etiqueta] = valor
    return kv


def _mapear_detalle_emitido_a_pdf(
    detalle: dict,
    tipo_visible: str,
    tipo_serie_texto: str,
    clave_texto: str,
    ruc_emisor: str | None = None,
) -> dict:
    datos = {col: "" for col in PDF_REPORT_COLUMNS}
    datos["tipoDocumento"] = (tipo_visible or "").strip()
    if ruc_emisor:
        datos["rucEmisor"] = ruc_emisor
    if clave_texto:
        datos["claveAcceso"] = clave_texto

    def _set_text(key: str, value: str):
        if value is None:
            return
        val = str(value).strip()
        if val:
            datos[key] = val

    def _set_num(key: str, value: str):
        if value is None:
            return
        val = _parse_decimal(str(value))
        if val is not None:
            datos[key] = val

    mapa_directo = {
        "ambiente": "ambiente",
        "tipoemision": "emision",
        "razonsocial": "razonSocialEmisor",
        "nombrecomercial": "nombreComercial",
        "numeroruc": "rucEmisor",
        "claveacceso": "claveAcceso",
        "establecimiento": "establecimiento",
        "puntoemision": "puntoEmision",
        "secuencial": "secuencial",
        "direccionmatriz": "direccionMatrizEmisor",
        "direccionestablecimiento": "direccionSucursalEmisor",
        "fechaemision": "fechaEmision",
        "contribuyenteespecial": "contribuyenteEspecial",
        "obligadocontabilidad": "obligadoContabilidad",
        "contribuyenteregimenrimpe": "tipoContribuyenteRIMPE",
        "contribuyenterimpe": "tipoContribuyenteRIMPE",
        "agenteretencionnroresolucion": "agenteRetencion",
        "razonsocialcomprador": "razonSocialComprador",
        "identificacioncomprador": "identificacionComprador",
        "placa": "placa",
        "placamatricula": "placa",
        "guiaremision": "guia",
        "comprobantemodificado": "comprobanteModificado",
        "fechaemisionmodificado": "fechaEmisionModificado",
        "razonmodificacion": "razonModificacion",
        "valormodificacion": "valorModificacion",
        "informacionadicional": "informacionAdicional",
        "formadepago": "formaPago",
    }

    for etiqueta, valor in (detalle or {}).items():
        token = _normalizar_token(etiqueta)
        if not token:
            continue
        if token in mapa_directo:
            _set_text(mapa_directo[token], valor)
            continue
        if "subtotal" in token:
            if "tarifaespecial" in token:
                _set_num("subtotalTarifaEspecial", valor)
            elif "15" in token:
                _set_num("subtotal15", valor)
            elif "12" in token:
                _set_num("subtotal12", valor)
            elif "8" in token:
                _set_num("subtotal8", valor)
            elif "5" in token:
                _set_num("subtotal5", valor)
            elif "0" in token:
                _set_num("subtotal0", valor)
            elif "noobjetoiva" in token:
                _set_num("subtotalNoObjetoIVA", valor)
            elif "exentoiva" in token:
                _set_num("subtotalExentoIVA", valor)
            elif "sinimpuestos" in token or "sinimpuesto" in token:
                _set_num("subtotalSinImpuestos", valor)
            continue
        if token.startswith("iva"):
            if "tarifaespecial" in token:
                _set_num("ivaTarifaEspecial", valor)
            elif "15" in token:
                _set_num("iva15", valor)
            elif "12" in token:
                _set_num("iva12", valor)
            elif "8" in token:
                _set_num("iva8", valor)
            elif "5" in token:
                _set_num("iva5", valor)
            continue
        if "ice" in token:
            _set_num("ice", valor)
            continue
        if "irbpnr" in token:
            _set_num("irbpnr", valor)
            continue
        if "propina" in token:
            _set_num("propina", valor)
            continue
        if "totalsinimpuestos" in token:
            _set_num("subtotalSinImpuestos", valor)
            continue
        if "totaldescuento" in token:
            _set_num("totalDescuento", valor)
            continue
        if "importetotal" in token or "valortotal" in token:
            _set_num("valorTotal", valor)
            continue
        if "valortotalsinsubsidio" in token:
            _set_num("valorTotalSinSubsidio", valor)
            continue

    if not datos.get("numeroComprobante"):
        est = datos.get("establecimiento") or ""
        pto = datos.get("puntoEmision") or ""
        sec = datos.get("secuencial") or ""
        if est and pto and sec:
            datos["numeroComprobante"] = f"{est}-{pto}-{sec}"
        elif tipo_serie_texto:
            match = re.search(r"\d{3}-\d{3}-\d{9}", tipo_serie_texto)
            if match:
                datos["numeroComprobante"] = match.group(0)

    return datos


def _obtener_source_detalle_emitido(page, row_index: int) -> str:
    try:
        return page.evaluate(
            """(idx) => {
                const rows = document.querySelectorAll("#frmPrincipal\\\\:tablaCompEmitidos_data tr");
                const row = rows[idx];
                if (!row) return "";
                const candidatos = Array.from(row.querySelectorAll("a, button"));
                const score = (el) => {
                    const id = el.id || "";
                    const onclick = (el.getAttribute("onclick") || "").toLowerCase();
                    const title = (el.getAttribute("title") || "").toLowerCase();
                    const aria = (el.getAttribute("aria-label") || "").toLowerCase();
                    const txt = (el.textContent || "").toLowerCase();
                    if (onclick.includes("panel-detalle-factura") || onclick.includes("detalle")) return 3;
                    if (title.includes("detalle") || aria.includes("detalle") || txt.includes("detalle") || txt.includes("ver")) return 2;
                    if (id && !id.includes("lnkPdf") && !id.includes("lnkXml")) return 1;
                    return 0;
                };
                let best = "";
                let bestScore = 0;
                for (const el of candidatos) {
                    if (!el.id) continue;
                    const s = score(el);
                    if (s > bestScore) {
                        bestScore = s;
                        best = el.id;
                    }
                }
                if (best) return best;
                const conId = candidatos.find(el => el.id);
                return conId ? conId.id : "";
            }""",
            row_index,
        )
    except Exception:
        return ""


def _snapshot_filas_emitidos(page, es_rechazado: bool = False) -> list[dict]:
    selector = "#frmPrincipal\\\\:tablaCompRechazados_data tr" if es_rechazado else "#frmPrincipal\\\\:tablaCompEmitidos_data tr"
    try:
        rows = page.evaluate(
            """(selector) => {
                const allRows = Array.from(document.querySelectorAll(selector));
                const textOf = (cells, idx) => ((cells[idx]?.innerText) || "").trim();
                const bestDetailId = (row) => {
                    const candidatos = Array.from(row.querySelectorAll("a, button"));
                    const score = (el) => {
                        const id = el.id || "";
                        const onclick = (el.getAttribute("onclick") || "").toLowerCase();
                        const title = (el.getAttribute("title") || "").toLowerCase();
                        const aria = (el.getAttribute("aria-label") || "").toLowerCase();
                        const txt = (el.textContent || "").toLowerCase();
                        if (onclick.includes("panel-detalle-factura") || onclick.includes("detalle")) return 3;
                        if (title.includes("detalle") || aria.includes("detalle") || txt.includes("detalle") || txt.includes("ver")) return 2;
                        if (id && !id.includes("lnkPdf") && !id.includes("lnkXml")) return 1;
                        return 0;
                    };
                    let best = "";
                    let bestScore = 0;
                    for (const el of candidatos) {
                        if (!el.id) continue;
                        const s = score(el);
                        if (s > bestScore) {
                            bestScore = s;
                            best = el.id;
                        }
                    }
                    return best;
                };
                return allRows.map((row, idx) => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    const pdfNode = row.querySelector("a[id$=':lnkPdf'], a[title*='pdf' i], button[title*='pdf' i], img[alt*='pdf' i], img[title*='pdf' i]");
                    const pdfAnchor = pdfNode ? (pdfNode.closest("a") || pdfNode) : null;
                    return {
                        row_index: idx,
                        total_celdas: cells.length,
                        tipo_serie_texto: textOf(cells, 1),
                        clave_texto: textOf(cells, 2),
                        fecha_aut_texto: textOf(cells, 3),
                        razon_texto: textOf(cells, 4),
                        valor_sin_imp_texto: textOf(cells, 5),
                        iva_texto: textOf(cells, 6),
                        importe_total_texto: textOf(cells, 7),
                        pdf_link_id: pdfAnchor && pdfAnchor.id ? pdfAnchor.id : "",
                        source_id_detalle: bestDetailId(row),
                    };
                });
            }""",
            selector,
        )
        return rows if isinstance(rows, list) else []
    except Exception:
        return []


def _extraer_lineas_pdf_layout(pdf_path: Path) -> list[dict]:
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTTextContainer, LTTextLine
    except Exception:
        return []
    lineas = []
    try:
        for page_layout in extract_pages(str(pdf_path)):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    for linea in element:
                        if isinstance(linea, LTTextLine):
                            texto = (linea.get_text() or "").strip()
                            if not texto:
                                continue
                            x0, y0, x1, y1 = linea.bbox
                            lineas.append(
                                {
                                    "text": texto,
                                    "x0": float(x0),
                                    "x1": float(x1),
                                    "y0": float(y0),
                                    "y1": float(y1),
                                }
                            )
    except Exception:
        return []
    for linea in lineas:
        linea["norm"] = _normalizar_label_simple(linea.get("text") or "")
    return lineas


def _extraer_numero_desde_texto(texto: str) -> str:
    if not texto:
        return ""
    match = re.search(r"([0-9][0-9.,]*)", texto)
    if not match:
        return ""
    return match.group(1)


def _buscar_valor_layout(lineas: list[dict], etiquetas: list[str], y_tol: float = 2.5) -> str:
    if not lineas:
        return ""
    etiquetas_norm = [_normalizar_label_simple(e) for e in etiquetas if e]
    for linea in lineas:
        norm = linea.get("norm") or ""
        if not any(e in norm for e in etiquetas_norm):
            continue
        y_obj = linea["y0"]
        x_ref = linea["x1"] + 2
        candidatos = [
            l
            for l in lineas
            if abs(l["y0"] - y_obj) <= y_tol and l["x0"] >= x_ref
        ]
        if candidatos:
            candidatos.sort(key=lambda l: l["x0"])
            return candidatos[0].get("text", "").strip()
        candidatos = [
            l
            for l in lineas
            if 0 < (y_obj - l["y0"]) <= 12 and l["x0"] >= linea["x0"]
        ]
        if candidatos:
            candidatos.sort(key=lambda l: (-l["y0"], l["x0"]))
            return candidatos[0].get("text", "").strip()
    return ""


def _extraer_datos_pdf_layout(pdf_path: Path) -> dict:
    lineas = _extraer_lineas_pdf_layout(pdf_path)
    if not lineas:
        return {}
    datos = {}
    # Campos de texto
    texto_map = {
        "ambiente": ["AMBIENTE"],
        "emision": ["EMISION", "TIPO EMISION"],
        "direccionMatrizEmisor": ["DIRECCION MATRIZ", "DIRECCION MATRIS", "DIRECCION MATRIZ:"],
        "direccionSucursalEmisor": ["DIRECCION SUCURSAL", "DIRECCION ESTABLECIMIENTO"],
        "obligadoContabilidad": ["OBLIGADO A LLEVAR CONTABILIDAD", "OBLIGADO CONTABILIDAD"],
        "agenteRetencion": ["AGENTE DE RETENCION RESOLUCION", "AGENTE RETENCION RESOLUCION"],
        "contribuyenteEspecial": ["CONTRIBUYENTE ESPECIAL"],
        "tipoContribuyenteRIMPE": ["CONTRIBUYENTE REGIMEN RIMPE", "CONTRIBUYENTE RIMPE"],
        "razonSocialComprador": ["RAZON SOCIAL / NOMBRES Y APELLIDOS", "RAZON SOCIAL COMPRADOR"],
        "identificacionComprador": ["IDENTIFICACION COMPRADOR", "IDENTIFICACION"],
        "direccionComprador": ["DIRECCION COMPRADOR", "DIRECCION ADQUIRENTE", "DIRECCION:"],
        "placa": ["PLACA", "MATRICULA"],
        "guia": ["GUIA REMISION", "GUIA DE REMISION", "GUIA"],
        "formaPago": ["FORMA PAGO", "FORMA DE PAGO"],
        "informacionAdicional": ["INFORMACION ADICIONAL"],
        "fechaEmision": ["FECHA EMISION", "FECHA"],
        "fechaAutorizacion": ["FECHA Y HORA DE AUTORIZACION", "FECHA Y HORA DE"],
    }
    for campo, etiquetas in texto_map.items():
        valor = _buscar_valor_layout(lineas, etiquetas)
        if valor:
            datos[campo] = valor

    # Clave de acceso
    clave = ""
    for linea in lineas:
        match = re.search(r"(\d{49})", linea.get("text", ""))
        if match:
            clave = match.group(1)
            break
    if clave:
        datos["claveAcceso"] = clave

    # Numero comprobante
    for linea in lineas:
        match = re.search(r"(\d{3}-\d{3}-\d{9})", linea.get("text", ""))
        if match:
            datos["numeroComprobante"] = match.group(1)
            partes = match.group(1).split("-")
            if len(partes) == 3:
                datos["establecimiento"] = partes[0]
                datos["puntoEmision"] = partes[1]
                datos["secuencial"] = partes[2]
            break

    # Totales
    numeric_map = {
        "subtotalTarifaEspecial": ["SUBTOTAL TARIFA ESPECIAL"],
        "subtotal15": ["SUBTOTAL 15%"],
        "subtotal12": ["SUBTOTAL 12%"],
        "subtotal8": ["SUBTOTAL 8%"],
        "subtotal5": ["SUBTOTAL 5%"],
        "subtotal0": ["SUBTOTAL 0%"],
        "subtotalNoObjetoIVA": ["SUBTOTAL NO OBJETO DE IVA"],
        "subtotalExentoIVA": ["SUBTOTAL EXENTO DE IVA"],
        "subtotalSinImpuestos": ["SUBTOTAL SIN IMPUESTOS", "TOTAL SIN IMPUESTOS"],
        "totalDescuento": ["TOTAL DESCUENTO"],
        "ivaTarifaEspecial": ["IVA TARIFA ESPECIAL"],
        "iva15": ["IVA 15%"],
        "iva12": ["IVA 12%"],
        "iva8": ["IVA 8%"],
        "iva5": ["IVA 5%"],
        "ice": ["ICE"],
        "irbpnr": ["IRBPNR"],
        "propina": ["PROPINA"],
        "valorTotal": ["IMPORTE TOTAL", "VALOR TOTAL"],
        "valorTotalSinSubsidio": ["VALOR TOTAL SIN SUBSIDIO"],
    }
    for campo, etiquetas in numeric_map.items():
        valor_raw = _buscar_valor_layout(lineas, etiquetas)
        valor_num = _extraer_numero_desde_texto(valor_raw)
        if valor_num:
            datos[campo] = valor_num

    # Forma de pago con monto
    if datos.get("formaPago"):
        texto = datos["formaPago"]
        match = re.search(r"([0-9][0-9.,]+)$", texto)
        if match:
            datos["formaPagoMonto"] = match.group(1)
            datos["formaPago"] = texto[:match.start()].strip(" -")

    return datos


def _obtener_detalle_emitido_xhr(
    page,
    request_context,
    source_id: str,
    payload_base: dict,
    view_state: str,
    tipo_visible: str,
    tipo_serie_texto: str,
    clave_texto: str,
    ruc_emisor: str | None = None,
) -> Optional[dict]:
    if not source_id or not view_state:
        return None
    payload = dict(payload_base or {})
    payload.update(
        {
            "javax.faces.partial.ajax": "true",
            "javax.faces.source": source_id,
            "javax.faces.partial.execute": source_id,
            "javax.faces.partial.render": "form-detalle-factura:panel-detalle-factura",
            source_id: source_id,
            "javax.faces.ViewState": view_state,
        }
    )
    payload.setdefault("frmPrincipal", "frmPrincipal")
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Faces-Request": "partial/ajax",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/xml, text/xml, */*; q=0.01",
        "Origin": "https://srienlinea.sri.gob.ec",
        "Referer": page.url,
    }
    try:
        respuesta = request_context.post(RECUPERAR_COMPROBANTES_URL, data=payload, headers=headers)
    except Exception:
        return None
    if respuesta.status != 200:
        return None
    try:
        cuerpo = respuesta.text()
    except Exception:
        return None
    detalle = _extraer_detalle_emitido_desde_partial(cuerpo)
    if not detalle:
        return None
    return _mapear_detalle_emitido_a_pdf(
        detalle, tipo_visible, tipo_serie_texto, clave_texto, ruc_emisor=ruc_emisor
    )


def _click_texto(page, texto: str) -> bool:
    for metodo in [
        lambda: page.get_by_role("button", name=texto, exact=False),
        lambda: page.get_by_text(texto, exact=False),
        lambda: page.locator(f"//button[contains(., '{texto}') or @title[contains(.,'{texto}')]]")
    ]:
        try:
            metodo().first.click(timeout=1000)
            return True
        except Exception:
            continue
    return False


def _click_consultar_emitidos(page) -> bool:
    selectores = [
        "button#frmPrincipal\\:btnBuscar",
        "button[id$='consultar']",
        "input[id$='consultar']",
        "button:has-text('Consultar')",
        "input[value='Consultar']",
        "a:has-text('Consultar')",
    ]
    for selector in selectores:
        try:
            locator = page.locator(selector)
            if not locator.count():
                continue
            try:
                locator.first.scroll_into_view_if_needed(timeout=500)
            except Exception:
                pass
            try:
                locator.first.click(timeout=800, force=True)
                return True
            except Exception:
                continue
        except Exception:
            continue
    try:
        return bool(
            page.evaluate(
                """() => {
                    const candidatos = Array.from(document.querySelectorAll('button, input, a'));
                    const normalizar = (t) => (t || '').trim().toLowerCase();
                    const btn = candidatos.find(el => {
                        const txt = normalizar(el.textContent);
                        const val = normalizar(el.value);
                        return txt === 'consultar' || val === 'consultar';
                    });
                    if (btn && typeof btn.click === 'function') {
                        btn.click();
                        return true;
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False

def _resolver_destino_unico(base_path: Path, extension: str) -> Path:
    """
    Ajusta el nombre final asegurando que exista la extension indicada
    y evitando colisiones con archivos existentes.
    """
    extension = extension if extension.startswith(".") else f".{extension}"
    destino = base_path.with_suffix(extension)
    contador = 1
    while destino.exists():
        destino = base_path.with_name(f"{base_path.stem}_{contador}{extension}")
        contador += 1
    return destino

def _es_respuesta_pdf(response) -> bool:
    try:
        headers = response.headers or {}
    except Exception:
        headers = {}
    content_type = headers.get("content-type", "").lower()
    content_disposition = headers.get("content-disposition", "").lower()
    if "application/pdf" in content_type:
        return True
    if "pdf" in content_disposition:
        return True
    return False

def _guardar_pdf_desde_enlace(page, link_locator, base_destino: Path) -> Optional[Path]:
    """
    Intenta descargar un PDF a partir de un locator. Primero espera un evento de download.
    Si no llega, intenta capturar la respuesta HTTP del PDF y guardarla manualmente.
    Devuelve la ruta final del archivo o None si salio mal.
    """
    def _click_objetivo():
        try:
            link_locator.scroll_into_view_if_needed(timeout=500)
        except Exception:
            pass
        link_locator.click(no_wait_after=True)

    errores = []

    # Primer intento: evento de descarga tradicional
    try:
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as descarga_info:
            _click_objetivo()
        descarga = descarga_info.value
        sugerido = Path(descarga.suggested_filename or base_destino.name)
        extension = sugerido.suffix or ".pdf"
        destino_final = _resolver_destino_unico(base_destino, extension)
        descarga.save_as(str(destino_final))
        return destino_final
    except PlaywrightTimeoutError as err:
        errores.append(f"Falla al esperar descarga directa de PDF (timeout): {err}")
    except Exception as err:
        errores.append(f"Falla al esperar descarga directa de PDF: {err}")

    # Segundo intento: capturar la respuesta HTTP del PDF
    try:
        with page.expect_response(
            lambda response: _es_respuesta_pdf(response), timeout=DOWNLOAD_TIMEOUT
        ) as respuesta_info:
            _click_objetivo()
        respuesta = respuesta_info.value
        if not _es_respuesta_pdf(respuesta):
            return None
        try:
            cuerpo = respuesta.body()
        except Exception:
            return None
        if not cuerpo:
            return None
        headers = respuesta.headers or {}
        sugerido = headers.get("content-disposition", "")
        extension = ".pdf"
        if "filename=" in sugerido:
            nombre = sugerido.split("filename=")[-1].strip().strip('"').strip("'")
            if "." in nombre:
                extension = Path(nombre).suffix or extension
        destino_final = _resolver_destino_unico(base_destino, extension)
        Path(destino_final).write_bytes(cuerpo)
        return destino_final
    except PlaywrightTimeoutError as err:
        errores.append(f"Falla al capturar respuesta PDF (timeout): {err}")
    except Exception as err:
        errores.append(f"Falla al capturar respuesta PDF: {err}")

    for mensaje in errores:
        print(f"[WARN] {mensaje}")
    return None


def _guardar_pdf_desde_jsf(page, link_locator, base_destino: Path) -> Optional[Path]:
    """
    Ejecuta directamente mojarra.jsfcljs (JSF) para disparar la descarga del PDF.
    Evita depender del click visual del boton/imagen.
    """
    try:
        link_id = link_locator.get_attribute("id")
    except Exception:
        link_id = None
    if not link_id:
        return None
    try:
        form_id = link_locator.evaluate("el => el.closest('form')?.id || 'frmPrincipal'")
    except Exception:
        form_id = "frmPrincipal"

    def _ejecutar_jsf():
        try:
            return page.evaluate(
                """({formId, linkId}) => {
                    const form = document.getElementById(formId);
                    if (!form) { return false; }
                    if (window.mojarra && typeof window.mojarra.jsfcljs === 'function') {
                        window.mojarra.jsfcljs(form, { [linkId]: linkId }, '');
                        return true;
                    }
                    if (typeof window.jsfcljs === 'function') {
                        window.jsfcljs(form, { [linkId]: linkId }, '');
                        return true;
                    }
                    const el = document.getElementById(linkId);
                    if (el) { el.click(); return true; }
                    return false;
                }""",
                arg={"formId": form_id, "linkId": link_id},
            )
        except Exception:
            return False

    try:
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as descarga_info:
            if not _ejecutar_jsf():
                return None
        descarga = descarga_info.value
        sugerido = Path(descarga.suggested_filename or base_destino.name)
        extension = sugerido.suffix or ".pdf"
        destino_final = _resolver_destino_unico(base_destino, extension)
        descarga.save_as(str(destino_final))
        return destino_final
    except PlaywrightTimeoutError:
        pass
    except Exception:
        pass

    try:
        with page.expect_response(
            lambda response: _es_respuesta_pdf(response), timeout=DOWNLOAD_TIMEOUT
        ) as respuesta_info:
            if not _ejecutar_jsf():
                return None
        respuesta = respuesta_info.value
        if not _es_respuesta_pdf(respuesta):
            return None
        try:
            cuerpo = respuesta.body()
        except Exception:
            return None
        if not cuerpo:
            return None
        headers = respuesta.headers or {}
        sugerido = headers.get("content-disposition", "")
        extension = ".pdf"
        if "filename=" in sugerido:
            nombre = sugerido.split("filename=")[-1].strip().strip('"').strip("'")
            if "." in nombre:
                extension = Path(nombre).suffix or extension
        destino_final = _resolver_destino_unico(base_destino, extension)
        Path(destino_final).write_bytes(cuerpo)
        return destino_final
    except Exception:
        return None







def _descargar_pdf_recibidos_post(page, link_locator, base_destino: Path) -> Optional[Path]:
    try:
        link_id = link_locator.get_attribute("id")
    except Exception:
        link_id = None
    if not link_id:
        return None
    try:
        view_state = page.locator("input[name='javax.faces.ViewState']").first.get_attribute("value")
    except Exception:
        view_state = None
    if not view_state:
        return None

    def _input_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return loc.first.input_value()
        except Exception:
            return ""
        return ""

    def _checked_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return (loc.first.get_attribute("value") or "")
        except Exception:
            return ""
        return ""

    form_data = {"frmPrincipal": "frmPrincipal"}
    opciones = _checked_value("input[name='frmPrincipal:opciones']:checked")
    if opciones:
        form_data["frmPrincipal:opciones"] = opciones

    campos = {
        "frmPrincipal:ano": "select#frmPrincipal\:ano",
        "frmPrincipal:mes": "select#frmPrincipal\:mes",
        "frmPrincipal:dia": "select#frmPrincipal\:dia",
        "frmPrincipal:cmbTipoComprobante": "select#frmPrincipal\:cmbTipoComprobante",
    }
    for clave, selector in campos.items():
        valor = _input_value(selector)
        if valor:
            form_data[clave] = valor

    recaptcha_val = _input_value("textarea[name='g-recaptcha-response']")
    form_data["g-recaptcha-response"] = recaptcha_val
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id

    url = page.url.split("#")[0]
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": url,
    }
    try:
        respuesta = page.context.request.post(url, data=form_data, headers=headers)
    except Exception:
        return None
    try:
        status = respuesta.status
        if status and status >= 400:
            return None
    except Exception:
        pass
    if not _es_respuesta_pdf(respuesta):
        return None
    try:
        cuerpo = respuesta.body()
    except Exception:
        return None
    if not cuerpo:
        return None
    headers_resp = respuesta.headers or {}
    disposition = headers_resp.get("content-disposition", "")
    extension = ".pdf"
    if "filename=" in disposition:
        nombre = disposition.split("filename=")[-1].strip().strip('"').strip("'")
        if "." in nombre:
            extension = Path(nombre).suffix or extension
    destino_final = _resolver_destino_unico(base_destino, extension)
    try:
        Path(destino_final).write_bytes(cuerpo)
    except Exception:
        return None
    return destino_final




def _descargar_pdf_emitidos_post(page, link_locator, base_destino: Path) -> Optional[Path]:
    try:
        link_id = link_locator.get_attribute("id")
    except Exception:
        link_id = None
    if not link_id:
        return None
    try:
        view_state = page.locator("input[name='javax.faces.ViewState']").first.get_attribute("value")
    except Exception:
        view_state = None
    if not view_state:
        return None

    def _input_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return loc.first.input_value()
        except Exception:
            return ""
        return ""

    def _checked_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return (loc.first.get_attribute("value") or "")
        except Exception:
            return ""
        return ""

    form_data = {"frmPrincipal": "frmPrincipal"}
    opciones = _checked_value("input[name='frmPrincipal:opciones']:checked")
    if opciones:
        form_data["frmPrincipal:opciones"] = opciones

    campos = {
        "frmPrincipal:calendarFechaDesde_input": "input#frmPrincipal\:calendarFechaDesde_input",
        "frmPrincipal:cmbEstadoAutorizacion": "select#frmPrincipal\:cmbEstadoAutorizacion",
        "frmPrincipal:cmbTipoComprobante": "select#frmPrincipal\:cmbTipoComprobante",
        "frmPrincipal:cmbEstablecimiento": "select#frmPrincipal\:cmbEstablecimiento",
        "frmPrincipal:txtPuntoEmision": "input#frmPrincipal\:txtPuntoEmision",
    }
    for clave, selector in campos.items():
        valor = _input_value(selector)
        if valor:
            form_data[clave] = valor

    recaptcha_val = _input_value("textarea[name='g-recaptcha-response']")
    if recaptcha_val:
        form_data["g-recaptcha-response"] = recaptcha_val
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id

    url = page.url.split("#")[0]
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": url,
    }
    try:
        respuesta = page.context.request.post(url, data=form_data, headers=headers)
    except Exception:
        return None
    try:
        status = respuesta.status
        if status and status >= 400:
            return None
    except Exception:
        pass
    if not _es_respuesta_pdf(respuesta):
        return None
    try:
        cuerpo = respuesta.body()
    except Exception:
        return None
    if not cuerpo:
        return None
    headers_resp = respuesta.headers or {}
    disposition = headers_resp.get("content-disposition", "")
    extension = ".pdf"
    if "filename=" in disposition:
        nombre = disposition.split("filename=")[-1].strip().strip('"').strip("'")
        if "." in nombre:
            extension = Path(nombre).suffix or extension
    destino_final = _resolver_destino_unico(base_destino, extension)
    try:
        Path(destino_final).write_bytes(cuerpo)
    except Exception:
        return None
    return destino_final


def _descargar_pdf_recibidos_post_con_viewstate(
    page,
    link_id: str,
    view_state: str,
    base_destino: Path,
) -> Optional[Path]:
    if not link_id or not view_state:
        return None

    def _input_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return loc.first.input_value()
        except Exception:
            return ""
        return ""

    def _checked_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return (loc.first.get_attribute("value") or "")
        except Exception:
            return ""
        return ""

    form_data = {"frmPrincipal": "frmPrincipal"}
    opciones = _checked_value("input[name='frmPrincipal:opciones']:checked")
    if opciones:
        form_data["frmPrincipal:opciones"] = opciones

    campos = {
        "frmPrincipal:ano": "select#frmPrincipal\\:ano",
        "frmPrincipal:mes": "select#frmPrincipal\\:mes",
        "frmPrincipal:dia": "select#frmPrincipal\\:dia",
        "frmPrincipal:cmbTipoComprobante": "select#frmPrincipal\\:cmbTipoComprobante",
    }
    for clave, selector in campos.items():
        valor = _input_value(selector)
        if valor:
            form_data[clave] = valor

    recaptcha_val = _input_value("textarea[name='g-recaptcha-response']")
    form_data["g-recaptcha-response"] = recaptcha_val
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id

    url = page.url.split("#")[0]
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": url,
    }
    try:
        respuesta = page.context.request.post(url, data=form_data, headers=headers)
    except Exception:
        return None
    try:
        status = respuesta.status
        if status and status >= 400:
            return None
    except Exception:
        pass
    if not _es_respuesta_pdf(respuesta):
        return None
    try:
        cuerpo = respuesta.body()
    except Exception:
        return None
    if not cuerpo:
        return None
    headers_resp = respuesta.headers or {}
    disposition = headers_resp.get("content-disposition", "")
    extension = ".pdf"
    if "filename=" in disposition:
        nombre = disposition.split("filename=")[-1].strip().strip('"').strip("'")
        if "." in nombre:
            extension = Path(nombre).suffix or extension
    destino_final = _resolver_destino_unico(base_destino, extension)
    try:
        Path(destino_final).write_bytes(cuerpo)
    except Exception:
        return None
    return destino_final

def _descargar_pdf_emitidos_post_con_viewstate(
    page,
    link_id: str,
    view_state: str,
    base_destino: Path,
) -> Optional[Path]:
    if not link_id or not view_state:
        return None

    def _input_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return loc.first.input_value()
        except Exception:
            return ""
        return ""

    def _checked_value(selector: str) -> str:
        try:
            loc = page.locator(selector)
            if loc.count():
                return (loc.first.get_attribute("value") or "")
        except Exception:
            return ""
        return ""

    form_data = {"frmPrincipal": "frmPrincipal"}
    opciones = _checked_value("input[name='frmPrincipal:opciones']:checked")
    if opciones:
        form_data["frmPrincipal:opciones"] = opciones

    campos = {
        "frmPrincipal:calendarFechaDesde_input": "input#frmPrincipal\\:calendarFechaDesde_input",
        "frmPrincipal:cmbEstadoAutorizacion": "select#frmPrincipal\\:cmbEstadoAutorizacion",
        "frmPrincipal:cmbTipoComprobante": "select#frmPrincipal\\:cmbTipoComprobante",
        "frmPrincipal:cmbEstablecimiento": "select#frmPrincipal\\:cmbEstablecimiento",
        "frmPrincipal:txtPuntoEmision": "input#frmPrincipal\\:txtPuntoEmision",
    }
    for clave, selector in campos.items():
        valor = _input_value(selector)
        if valor:
            form_data[clave] = valor

    recaptcha_val = _input_value("textarea[name='g-recaptcha-response']")
    if recaptcha_val:
        form_data["g-recaptcha-response"] = recaptcha_val
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id

    url = page.url.split("#")[0]
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": url,
    }
    try:
        respuesta = page.context.request.post(url, data=form_data, headers=headers)
    except Exception:
        return None
    try:
        status = respuesta.status
        if status and status >= 400:
            return None
    except Exception:
        pass
    if not _es_respuesta_pdf(respuesta):
        return None
    try:
        cuerpo = respuesta.body()
    except Exception:
        return None
    if not cuerpo:
        return None
    headers_resp = respuesta.headers or {}
    disposition = headers_resp.get("content-disposition", "")
    extension = ".pdf"
    if "filename=" in disposition:
        nombre = disposition.split("filename=")[-1].strip().strip('"').strip("'")
        if "." in nombre:
            extension = Path(nombre).suffix or extension
    destino_final = _resolver_destino_unico(base_destino, extension)
    try:
        Path(destino_final).write_bytes(cuerpo)
    except Exception:
        return None
    return destino_final

def _es_respuesta_xml(response) -> bool:
    try:
        headers = response.headers or {}
    except Exception:
        headers = {}
    content_type = headers.get("content-type", "").lower()
    content_disposition = headers.get("content-disposition", "").lower()
    if "xml" in content_type:
        return True
    if "xml" in content_disposition:
        return True
    return False

def _guardar_xml_desde_enlace(page, link_locator, base_destino: Path) -> Optional[Path]:
    """
    Descarga un XML ya sea por evento de download o capturando la respuesta HTTP asociada.
    """
    def _click_objetivo():
        try:
            link_locator.scroll_into_view_if_needed(timeout=500)
        except Exception:
            pass
        link_locator.click(no_wait_after=True)

    errores = []

    try:
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as descarga_info:
            _click_objetivo()
        descarga = descarga_info.value
        sugerido = Path(descarga.suggested_filename or f"{base_destino.name}.xml")
        extension = sugerido.suffix or ".xml"
        destino_final = _resolver_destino_unico(base_destino, extension)
        descarga.save_as(str(destino_final))
        return destino_final
    except PlaywrightTimeoutError as err:
        errores.append(f"Falla al esperar descarga directa de XML (timeout): {err}")
    except Exception as err:
        errores.append(f"Falla al esperar descarga directa de XML: {err}")

    try:
        with page.expect_response(
            lambda response: _es_respuesta_xml(response), timeout=DOWNLOAD_TIMEOUT
        ) as respuesta_info:
            _click_objetivo()
        respuesta = respuesta_info.value
        if not _es_respuesta_xml(respuesta):
            return None
        try:
            cuerpo = respuesta.body()
        except Exception:
            cuerpo = b""
        if not cuerpo:
            return None
        headers = respuesta.headers or {}
        sugerido = headers.get("content-disposition", "")
        extension = ".xml"
        if "filename=" in sugerido:
            nombre = sugerido.split("filename=")[-1].strip().strip('"').strip("'")
            if "." in nombre:
                extension = Path(nombre).suffix or extension
        destino_final = _resolver_destino_unico(base_destino, extension)
        destino_final.write_bytes(cuerpo)
        return destino_final
    except PlaywrightTimeoutError as err:
        errores.append(f"Falla al capturar respuesta XML (timeout): {err}")
    except Exception as err:
        errores.append(f"Falla al capturar respuesta XML: {err}")

    for mensaje in errores:
        print(f"[WARN] {mensaje}")
    return None

def _seleccionar(page, etiqueta: str, valor_visible: str):
    try:
        sel = page.get_by_label(etiqueta, exact=False).locator("select")
        sel.select_option(label=valor_visible)
        return True
    except Exception:
        pass
    try:
        page.locator(f"text={etiqueta}").locator("xpath=..").locator("select").select_option(label=valor_visible)
        return True
    except Exception:
        pass
    return False


def _normalizar_token(texto: str) -> str:
    base = unicodedata.normalize("NFKD", (texto or "")).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", "", base)


def _rellenar_input_por_label(page, etiquetas, valor: str, selectores_extra=None) -> bool:
    if valor is None:
        return False
    texto = (valor or "").strip()
    if not texto:
        return False
    if isinstance(etiquetas, str):
        etiquetas = [etiquetas]
    if selectores_extra is None:
        selectores_extra = []
    try:
        selectores_extra = list(selectores_extra)
    except TypeError:
        selectores_extra = [selectores_extra]

    for etiqueta in etiquetas:
        try:
            loc = page.get_by_label(etiqueta, exact=False)
            if loc.count():
                loc.first.fill("")
                loc.first.fill(texto)
                return True
        except Exception:
            continue

    tokens = [_normalizar_token(e) for e in etiquetas if e]
    tokens = [t for t in tokens if t]
    candidatos = list(selectores_extra)
    for token in tokens:
        candidatos.extend([
            f"input[name*='{token}']",
            f"input[id*='{token}']",
            f"input[data-testid*='{token}']",
        ])

    for selector in candidatos:
        try:
            loc = page.locator(selector)
            if loc.count():
                loc.first.fill("")
                loc.first.fill(texto)
                return True
        except Exception:
            continue

    return False


def _seleccionar_en_select(page, selector: str, *valores) -> bool:
    try:
        locator = page.locator(selector)
    except Exception:
        return False
    if not locator or not locator.count():
        return False
    valores_limpios = [
        v for v in (valores or []) if v is not None and str(v).strip()
    ]

    def _seleccionar_js() -> bool:
        try:
            return bool(
                page.evaluate(
                    """({selector, valores}) => {
                        const el = document.querySelector(selector);
                        if (!el) { return false; }
                        const candidatos = valores
                            .map(v => (v == null ? "" : String(v)))
                            .filter(v => v.trim().length > 0);
                        if (!candidatos.length) { return false; }
                        const comparar = (texto) => {
                            const base = (texto || "").toString().normalize("NFD");
                            return base.replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
                        };
                        const stopwords = new Set(["de", "del", "la", "el", "los", "las", "y"]);
                        const singular = (token) => {
                            if (!token) return token;
                            if (token.length > 4 && token.endsWith("es")) {
                                return token.slice(0, -2);
                            }
                            if (token.length > 3 && token.endsWith("s")) {
                                return token.slice(0, -1);
                            }
                            return token;
                        };
                        const normalizarTokens = (texto) => {
                            const norm = comparar(texto);
                            const tokens = norm
                                .split(/\\s+/)
                                .filter(t => t && !stopwords.has(t))
                                .map(singular);
                            return tokens.join(" ");
                        };
                        const compactar = (texto) => {
                            return (texto || "").replace(/[^a-z0-9]+/g, "");
                        };
                        for (const opcion of Array.from(el.options)) {
                            const labelRaw = opcion.label || opcion.textContent || "";
                            const valueRaw = opcion.value || "";
                            const label = comparar(labelRaw);
                            const value = comparar(valueRaw);
                            const labelTokens = normalizarTokens(labelRaw);
                            const valueTokens = normalizarTokens(valueRaw);
                            const labelCompact = compactar(label);
                            const valueCompact = compactar(value);
                            const labelTokensCompact = compactar(labelTokens);
                            const valueTokensCompact = compactar(valueTokens);
                            for (const objetivo of candidatos) {
                                const norm = comparar(objetivo);
                                const normTokens = normalizarTokens(objetivo);
                                const normCompact = compactar(norm);
                                const normTokensCompact = compactar(normTokens);
                                const matchExact = norm && (norm === label || norm === value);
                                const matchTokens = normTokens && (normTokens === labelTokens || normTokens === valueTokens);
                                const matchCompact = normCompact && (normCompact === labelCompact || normCompact === valueCompact);
                                const matchTokensCompact = normTokensCompact && (normTokensCompact === labelTokensCompact || normTokensCompact === valueTokensCompact);
                                const matchContains = normTokens && (labelTokens.includes(normTokens) || valueTokens.includes(normTokens));
                                const matchContainsRev = labelTokens && (normTokens.includes(labelTokens) || normTokens.includes(valueTokens));
                                if (matchExact || matchTokens || matchCompact || matchTokensCompact || matchContains || matchContainsRev) {
                                    if (el.value !== opcion.value) {
                                        el.value = opcion.value;
                                        el.dispatchEvent(new Event('change', { bubbles: true }));
                                    }
                                    return true;
                                }
                            }
                        }
                        return false;
                    }""",
                    {"selector": selector, "valores": valores_limpios},
                )
            )
        except Exception:
            return False

    if valores_limpios and _seleccionar_js():
        return True
    if _seleccionar_por_label(locator, *valores):
        return True
    for valor in valores:
        if valor is None:
            continue
        try:
            locator.select_option(value=valor, timeout=1500)
            return True
        except Exception:
            continue
    if valores_limpios:
        try:
            page.wait_for_timeout(200)
        except Exception:
            pass
        if _seleccionar_js():
            return True
    return False

def _espera_captcha(page, timeout: int = 1000):
    try:
        loc = page.locator("img[alt='captcha']")
        if loc.is_visible(timeout=1000):
            page.wait_for_selector("img[alt='captcha']", state="detached", timeout=timeout)
    except Exception:
        pass


def _captcha_visible(page, timeout: int = 0) -> bool:
    try:
        loc = page.locator("img[alt='captcha']")
        if timeout:
            return loc.is_visible(timeout=timeout)
        return loc.is_visible()
    except Exception:
        return False


def _recaptcha_presente(page) -> bool:
    try:
        if page.locator("iframe[src*='recaptcha']").count():
            return True
    except Exception:
        pass
    try:
        if page.locator("[data-sitekey]").count():
            return True
    except Exception:
        pass
    return False


def _esperar_recaptcha_resuelto(page, timeout: int = 300000) -> bool:
    """Espera a que el desafío de reCAPTCHA desaparezca u obtenga respuesta."""
    fin = time.time() + timeout / 1000
    while time.time() < fin:
        challenge_activo = False
        try:
            challenge_activo = _recaptcha_challenge_activo(page)
        except Exception:
            challenge_activo = False
        if not challenge_activo:
            return True
        try:
            page.wait_for_timeout(500)
        except Exception:
            pass
    return False


def _recaptcha_challenge_activo(page) -> bool:
    try:
        frame = page.locator("iframe[src*='recaptcha/api2/bframe']")
        if frame.count():
            try:
                return frame.first.is_visible()
            except Exception:
                pass
    except Exception:
        pass
    return False


CAPTCHA_INPUT_SELECTORS = [
    "input[name*='captcha' i]",
    "input[id*='captcha' i]",
    "input[name='captcha']",
    "input[id='captcha']",
    "input#captchaIngresar",
    "input#captchaTxt",
]
CAPTCHA_INPUT_QUERY = ",".join(CAPTCHA_INPUT_SELECTORS)


def _esperar_captcha_manual_input(page, timeout: int = 300000) -> bool:
    """
    Espera hasta que el usuario ingrese un valor de captcha de forma manual.
    Se considera resuelto cuando cualquiera de los inputs registrados tiene texto.
    """
    try:
        page.wait_for_function(
            """(selectorCadena) => {
                const inputs = document.querySelectorAll(selectorCadena);
                for (const input of inputs) {
                    const valor = (input.value || "").trim();
                    if (valor.length >= 4) {
                        return true;
                    }
                }
                return false;
            }""",
            arg=CAPTCHA_INPUT_QUERY,
            timeout=timeout,
        )
        return True
    except Exception:
        return False


def _localizar_input_captcha(page):
    for selector in CAPTCHA_INPUT_SELECTORS:
        try:
            locator = page.locator(selector)
            if locator.count():
                return locator.first
        except Exception:
            continue
    return None


def _resolver_captcha(page, contexto: str) -> bool:
    recaptcha_detectado = False
    try:
        recaptcha_detectado = _recaptcha_presente(page)
    except Exception:
        recaptcha_detectado = False

    if recaptcha_detectado:
        _notificar_usuario_captcha("reCAPTCHA", contexto)
        _esperar_recaptcha_resuelto(page, timeout=300000)
        return True

    try:
        if not _captcha_visible(page, timeout=1000):
            return False
    except Exception:
        return False

    if not captcha_solver_enabled():
        _notificar_usuario_captcha("captcha de imagen", contexto)
        _espera_captcha(page)
        return False

    for intento in range(1, CAPTCHA_MAX_ATTEMPTS + 1):
        try:
            if not _captcha_visible(page, timeout=1000):
                return False
        except Exception:
            return False

        input_captcha = _localizar_input_captcha(page)
        if input_captcha is None:
            print(f"[WARN] Campo de texto para captcha no encontrado ({contexto}); esperando resolucion manual.")
            _notificar_usuario_captcha("captcha de imagen", contexto)
            _espera_captcha(page)
            return False

        try:
            imagen = page.locator("img[alt='captcha']").screenshot(type="png")
        except Exception as err:
            print(f"[WARN] No se pudo capturar la imagen del captcha (intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}")
            break

        try:
            codigo = solve_captcha_image(imagen)
        except CaptchaSolverError as err:
            print(f"[WARN] Fallo al resolver captcha con 2Captcha (intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}")
            continue

        try:
            input_captcha.fill("")
            input_captcha.fill(codigo)
            return True
        except Exception as err:
            print(f"[WARN] No se pudo escribir el captcha resuelto (intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}")

    print("[WARN] Se agotaron los intentos automaticos de captcha; esperando resolucion manual.")
    _notificar_usuario_captcha("captcha de imagen", contexto)
    _espera_captcha(page)
    return False

def _resolver_autenticacion_persistente(page) -> bool:
    """Gestiona la pantalla opcional de 'autenticacion persistente' mostrada por Keycloak."""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=1000)
    except Exception:
        pass

    objetivos = [page]
    try:
        objetivos.extend([frame for frame in page.frames if frame not in objetivos])
    except Exception:
        pass

    palabras_clave = [
        "autenticacion persistente",
        "autenticacion persistente",
        "mantener la sesion",
        "mantener la sesion",
        "permanecer conectado",
        "persistente",
        "recordar este dispositivo",
    ]

    selectores_prioridad = [
        "button#cancel",
        "input#cancel",
        "button[name='cancel']",
        "input[name='cancel']",
        "button[value='cancel']",
        "input[value='cancel']",
        "button[data-kc-button='cancel']",
        "button#kc-cancel",
        "input#kc-cancel",
        "a#cancel",
        "a[name='cancel']",
        "button:has-text('No')",
        "button:has-text('NO')",
        "button:has-text('Cancelar')",
        "button:has-text('Cancelar autenticacion')",
        "input[value='No']",
        "input[value='NO']",
        "input[value='Cancelar']",
        "button:has-text('Continuar sin recordar')",
        "button:has-text('No mantener')",
    ]

    secundarios = [
        "button:has-text('Si')",
        "button:has-text('Si')",
        "input[value='Si']",
        "button[type='submit']",
        "input[type='submit']",
    ]

    def _clic_en_objetivo(objetivo, selector):
        try:
            loc = objetivo.locator(selector)
            if not loc.count():
                return False
            loc = loc.first
            loc.wait_for(state="visible", timeout=1000)
            loc.click()
            try:
                page.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
            time.sleep(0.3)
            return True
        except Exception:
            return False

    for objetivo in objetivos:
        try:
            cuerpo = objetivo.locator("body")
            texto_cuerpo = (
                cuerpo.inner_text(timeout=1000).lower()
                if cuerpo.count()
                else ""
            )
        except Exception:
            texto_cuerpo = ""

        if texto_cuerpo and any(palabra in texto_cuerpo for palabra in palabras_clave):
            for selector in selectores_prioridad + secundarios:
                if _clic_en_objetivo(objetivo, selector):
                    return True

        # Si no hubo coincidencia de texto, igual intentar con selectores
        for selector in selectores_prioridad:
            if _clic_en_objetivo(objetivo, selector):
                return True

    # Como ultimo recurso, intentar con coincidencia textual global.
    for etiqueta in ["No", "No, gracias", "No mantener", "Cancelar", "Continuar"]:
        if _click_texto(page, etiqueta):
            try:
                page.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
            time.sleep(0.3)
            return True

    try:
        Path("debug_login_persistente.html").write_text(
            page.content(),
            encoding="utf-8",
            errors="ignore",
        )
    except Exception:
        pass
    return False

def _cerrar_modal_encuesta(page) -> bool:
    """Intenta cerrar la alerta de encuesta mostrada tras el inicio de sesion."""
    try:
        page.wait_for_timeout(200)
    except Exception:
        pass

    selectores_prioridad = [
        "[role='dialog'] button[aria-label*='Cerrar']",
        "[role='dialog'] button[title*='Cerrar']",
        "[role='dialog'] button[aria-label*='Close']",
        "[role='dialog'] button[title*='Close']",
        "[role='dialog'] button.close",
        "[role='dialog'] button.mat-icon-button",
        "[role='dialog'] button:has-text('×')",
        "[role='dialog'] button:has-text('X')",
        "[role='dialog'] button:has-text('?')",
        "button[aria-label='Cerrar']",
        "button[aria-label='Close']",
        "button[title='Cerrar']",
        "button[title='Close']",
        "button[aria-label*='Cerrar encuesta']",
    ]

    for selector in selectores_prioridad:
        try:
            boton = page.locator(selector)
        except Exception:
            continue
        if not boton.count():
            continue
        boton = boton.first
        try:
            boton.wait_for(state="visible", timeout=500)
        except Exception:
            pass
        try:
            if not boton.is_enabled():
                continue
        except Exception:
            pass
        try:
            boton.click()
            try:
                page.wait_for_timeout(300)
            except Exception:
                pass
            return True
        except Exception:
            continue

    try:
        cerrado = page.evaluate(
            """
            () => {
                const normalize = (texto) => {
                    if (!texto) {
                        return "";
                    }
                    if (typeof texto.normalize === "function") {
                        return texto.normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").toLowerCase();
                    }
                    return String(texto).toLowerCase();
                };
                const keywords = [
                    "tu opinion nos permite mejorar",
                    "encuesta de satisfaccion de servicios",
                    "quiero responder"
                ];
                const candidatos = Array.from(document.querySelectorAll("[role='dialog'], .mat-dialog-container, .modal, .cdk-overlay-pane"))
                    .map(el => el.closest("[role='dialog']") || el)
                    .filter(Boolean);
                const dialog = candidatos.find(el => {
                    const texto = normalize(el.innerText || el.textContent || "");
                    return keywords.some(keyword => texto.includes(keyword));
                });
                if (!dialog) {
                    return false;
                }
                const posibles = Array.from(dialog.querySelectorAll("button, a, span, mat-icon, i, div"));
                for (const elemento of posibles) {
                    const raw = (elemento.textContent || elemento.innerText || "").trim();
                    const texto = normalize(raw);
                    const aria = normalize(elemento.getAttribute && elemento.getAttribute("aria-label"));
                    const title = normalize(elemento.getAttribute && elemento.getAttribute("title"));
                    const clases = normalize(elemento.className || "");
                    if (["x", "×", "?"].includes(raw) ||
                        texto === "cerrar" ||
                        texto === "close" ||
                        aria.includes("cerrar") ||
                        title.includes("cerrar") ||
                        clases.includes("close") ||
                        clases.includes("cerrar")) {
                        if (typeof elemento.click === "function") {
                            elemento.click();
                            return true;
                        }
                    }
                }
                return false;
            }
            """
        )
        if cerrado:
            try:
                page.wait_for_timeout(300)
            except Exception:
                pass
            return True
    except Exception:
        pass

    for texto in ["×", "X", "Cerrar", "CERRAR", "Close", "CLOSE"]:
        if _click_texto(page, texto):
            try:
                page.wait_for_timeout(300)
            except Exception:
                pass
            return True
    return False

def _abrir_modulo_consultas(page, origen: str):
    """Cierra popups, abre el panel izquierdo y navega al formulario correspondiente."""
    _cerrar_modal_encuesta(page)
    destino_url = RECUPERAR_COMPROBANTES_URL if origen == "Emitidos" else RECIBIDOS_DIRECT_URL
    goto_timeout = 8000 if origen == "Emitidos" else 15000
    idle_timeout = 2000 if origen == "Emitidos" else 5000

    def _click_locator(selector: str, descripcion: str, timeout: int = 1200):
        try:
            locator = page.locator(selector)
            if not locator.count():
                return False
            locator.first.wait_for(state="visible", timeout=timeout)
            locator.first.click(timeout=timeout)
            return True
        except Exception as err:
            print(f"[WARN] No se pudo interactuar con {descripcion}: {err}")
            return False

    def _wait_overlays():
        for overlay in OVERLAY_SELECTORS:
            try:
                page.wait_for_selector(overlay, state="hidden", timeout=3000)
            except Exception:
                pass

    def _goto_form():
        ultimo_error = None
        for intento in range(3):
            try:
                page.goto(destino_url, wait_until="domcontentloaded", timeout=goto_timeout)
                page.wait_for_load_state("domcontentloaded", timeout=idle_timeout)
                page.wait_for_load_state("networkidle", timeout=idle_timeout)
                return
            except Exception as err:
                ultimo_error = err
                print(f"[WARN] Reintentando acceso directo al formulario ({intento + 1}/3): {err}")
        raise RuntimeError(f"No se pudo abrir el formulario de {origen.lower()}: {ultimo_error}")

    def _ensure_menu_visible():
        try:
            if page.locator(FACTURACION_MENU_SELECTOR).first.is_visible(timeout=500):
                return True
        except Exception:
            pass
        return _click_locator(MENU_TOGGLE_SELECTOR, "el botón de menú", timeout=1500)

    def _expand_panel(selector: str, descripcion: str):
        header = page.locator(selector)
        if not header.count():
            print(f"[WARN] No se encontró {descripcion}.")
            return False
        try:
            expanded = (header.first.get_attribute("aria-expanded") or "").lower()
        except Exception:
            expanded = ""
        if expanded == "true":
            return True
        try:
            header.first.click(timeout=1500)
            page.wait_for_timeout(200)
            return True
        except Exception as err:
            print(f"[WARN] No se pudo expandir {descripcion}: {err}")
            return False

    _wait_overlays()
    current_url = ""
    try:
        current_url = page.url or ""
    except Exception:
        current_url = ""
    if "consultas/menu.jsf" in current_url:
        _goto_form()
        return page

    _ensure_menu_visible()
    _expand_panel(FACTURACION_MENU_SELECTOR, "el panel de Facturación Electrónica")
    _expand_panel(MODULO_PRODUCCION_SELECTOR, "el panel Producción")

    def _en_formulario():
        try:
            current = page.url or ""
        except Exception:
            return False
        return (
            "comprobantesRecibidos.jsf" in current
            or "recibidos/comprobantesRecibidos.jsf" in current
            or "recuperarComprobantes.jsf" in current
        )

    consultas_locator = page.locator(CONSULTAS_SELECTOR)
    if consultas_locator.count():
        try:
            consultas_locator.first.click(timeout=1500)
        except Exception as err:
            print(f"[WARN] No se pudo hacer clic en 'Consultas': {err}")
    else:
        print("[WARN] No se encontró el botón de Consultas; intentando acceso directo.")

    try:
        page.wait_for_load_state("domcontentloaded", timeout=2000)
    except Exception:
        pass
    if not _en_formulario():
        _goto_form()

    return page


def _esperar_ajax(page, timeout: int = 1000):
    """Espera a que la cola AJAX de PrimeFaces quede vacia para evitar sobrescrituras."""
    try:
        page.wait_for_function(
            "() => window.PrimeFaces && PrimeFaces.ajax && PrimeFaces.ajax.Queue && PrimeFaces.ajax.Queue.isEmpty()",
            timeout=timeout,
        )
    except PlaywrightTimeoutError:
        try:
            page.wait_for_timeout(300)
        except Exception:
            pass


def _cerrar_sesion(pagina) -> bool:
    """Intenta cerrar la sesion activa del SRI en la pagina dada."""
    if not pagina:
        return False

    selectores = [
        "a[href*='salir.jspa']",
        "a[title*='Cerrar sesion']",
        "a[tooltip*='Cerrar sesion']",
        "a:has-text('Cerrar sesion')",
        "button:has-text('Cerrar sesion')",
    ]
    for selector in selectores:
        try:
            objetivo = pagina.locator(selector)
            if not objetivo.count():
                continue
            try:
                objetivo.first.scroll_into_view_if_needed(timeout=1000)
            except Exception:
                pass
            with pagina.expect_navigation(wait_until="load", timeout=1000):
                objetivo.first.click()
            try:
                pagina.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
            return True
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue

    salir_urls = [
        "https://srienlinea.sri.gob.ec/tuportal-internet/salir.jspa",
        "/tuportal-internet/salir.jspa",
    ]
    for url in salir_urls:
        try:
            pagina.goto(url, wait_until="load", timeout=1000)
            try:
                pagina.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
            return True
        except Exception:
            continue

    try:
        pagina.evaluate(
            """() => {
                if (window.location && typeof window.location.replace === "function") {
                    window.location.replace("https://srienlinea.sri.gob.ec/tuportal-internet/salir.jspa");
                    return true;
                }
                return false;
            }"""
        )
    except Exception:
        return False

    try:
        pagina.wait_for_load_state("load", timeout=1000)
    except Exception:
        pass
    return True

def _login(
    context,
    page,
    ruc,
    clave,
    cookies_path: Path,
    destino_url: str,
    ci_adicional: Optional[str] = None,
):
    """Realiza la autenticacion y gestiona la reutilizacion de cookies almacenadas."""

    def _navegar():
        try:
            page.goto(destino_url, timeout=1000)
        except Exception:
            pass
        try:
            page.wait_for_load_state("domcontentloaded", timeout=1000)
        except Exception:
            pass
        _asegurar_portal_disponible(page)

    if cookies_path.exists():
        try:
            context.add_cookies(json.loads(cookies_path.read_text()))
            _navegar()
            if "auth/realms" not in page.url:
                return
        except Exception:
            pass

    _navegar()

    if "auth/realms" in page.url:
        page.wait_for_selector(
            "input[name='usuario']:visible, input[name='username']:visible",
            timeout=1000,
        )

        usuario_selectores = [
            "input[name='usuario']:visible",
            "input[name='username']:visible",
            "input#usuario:visible",
            "input#username:visible",
        ]

        captcha_retry = 0
        while True:
            for selector in usuario_selectores:
                locator = page.locator(selector)
                if locator.count() and locator.first.is_enabled():
                    try:
                        locator.first.fill(ruc)
                        break
                    except Exception:
                        continue
            else:
                raise RuntimeError("No se encontro el campo de usuario en el formulario de login del SRI.")

            ci_locator = page.locator("input[name='ciAdicional']:visible")
            if ci_locator.count() and ci_locator.first.is_enabled():
                valor_ci = (ci_adicional or "").strip()
                try:
                    ci_locator.first.fill(valor_ci)
                except Exception:
                    pass

            pass_locator = page.locator("input[name='password']:visible, input[type='password']:visible")
            if pass_locator.count():
                pass_locator.first.fill(clave)
            else:
                raise RuntimeError("No se encontro el campo de contrasena en el formulario de login del SRI.")

            captcha_resuelto = _resolver_captcha(page, "login")

            if PAUSE_AT_LOGIN:
                print(PAUSE_PROMPT, flush=True)
                try:
                    input(">>> Presiona Enter para continuar con el clic en 'Ingresar'...")
                except EOFError:
                    print("Entrada estandar no disponible; continuando automaticamente en 5 segundos.", flush=True)
                    time.sleep(0.2)

            if not _click_texto(page, "Ingresar"):
                try:
                    page.locator("input[name='login'], button[type='submit']").first.click()
                except Exception:
                    keyboard = getattr(page, "keyboard", None)
                    if keyboard:
                        try:
                            keyboard.press("Enter")
                        except Exception:
                            pass
                    else:
                        print("[WARN] No se pudo accionar el boton 'Ingresar'; el objeto page no expone teclado.")

            try:
                page.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass

            autenticado = False
            for intento in range(2):
                try:
                    page.wait_for_url(lambda url: "auth/realms" not in url, timeout=1000)
                    autenticado = True
                    break
                except PlaywrightTimeoutError:
                    if _resolver_autenticacion_persistente(page):
                        continue
                    if intento == 1:
                        raise RuntimeError("No fue posible completar el login del SRI (pantalla de autenticacion persistente).")

            if autenticado and "auth/realms" not in page.url:
                break

            recaptcha_activo = False
            if captcha_resuelto:
                try:
                    recaptcha_activo = _recaptcha_challenge_activo(page)
                except Exception:
                    recaptcha_activo = False

            if captcha_resuelto and (_captcha_visible(page, timeout=1000) or recaptcha_activo):
                captcha_retry += 1
                if captcha_retry >= CAPTCHA_MAX_ATTEMPTS:
                    raise RuntimeError("No fue posible completar el login del SRI (captcha).")
                print(f"[INFO] Reintentando login por captcha adicional ({captcha_retry}/{CAPTCHA_MAX_ATTEMPTS}).")
                continue

            raise RuntimeError("No fue posible completar el login del SRI (credenciales o captcha).")
        if "auth/realms" in page.url:
            if not _resolver_autenticacion_persistente(page):
                raise RuntimeError("No fue posible completar el login del SRI (pantalla de autenticacion persistente).")
            try:
                page.wait_for_url(lambda url: "auth/realms" not in url, timeout=1000)
            except PlaywrightTimeoutError:
                raise RuntimeError("No fue posible completar el login del SRI (pantalla de autenticacion persistente).")

    if "auth/realms" not in page.url:
        _cerrar_modal_encuesta(page)
        try:
            cookies_path.write_text(json.dumps(context.cookies()))
        except Exception:
            pass

# ============================================================
# ?? DESCARGA DE COMPROBANTES RECIBIDOS (TXT + XML + PDF)
# ============================================================
def _seleccionar_por_label(locator, *labels) -> bool:
    for etiqueta in labels:
        if etiqueta is None:
            continue
        try:
            locator.select_option(label=etiqueta, timeout=1500)
            return True
        except Exception:
            try:
                locator.select_option(value=etiqueta, timeout=1500)
                return True
            except Exception:
                continue
    return False

def _flujo_recibidos(page, destino: Path, anio: int, mes: int, dia: int, tipo: str, formatos: list):
    _check_cancel("inicio_recibidos")
    if RECIBIDOS_DIRECT_URL not in (page.url or ""):
        try:
            page.goto(RECIBIDOS_DIRECT_URL, wait_until="domcontentloaded", timeout=5000)
            page.wait_for_load_state("domcontentloaded", timeout=1000)
        except Exception:
            pass
    page.wait_for_selector("select#frmPrincipal\\:ano", state="visible", timeout=10000)

    selector_ano_css = "select#frmPrincipal\\:ano"
    selector_mes_css = "select#frmPrincipal\\:mes"
    selector_dia_css = "select#frmPrincipal\\:dia"
    selector_tipo_css = "select#frmPrincipal\\:cmbTipoComprobante"
    mes_texto = _mes_a_texto(mes)
    dia_labels = ("Todos", "0") if dia in (None, 0) else (str(dia), f"{dia:02d}")
    tipo_visible = TIPOS_MAP.get(tipo, tipo)

    def _aplicar_filtros_recibidos(estricto: bool = True) -> bool:
        try:
            selector_ano = page.locator(selector_ano_css)
            if not _seleccionar_por_label(selector_ano, str(anio)):
                if estricto:
                    raise RuntimeError("No fue posible seleccionar el año solicitado en el SRI.")
                return False

            selector_mes = page.locator(selector_mes_css)
            if not _seleccionar_por_label(selector_mes, mes_texto, f"{mes:02d}", str(mes)):
                if estricto:
                    raise RuntimeError("No fue posible seleccionar el mes solicitado en el SRI.")
                return False

            selector_dia = page.locator(selector_dia_css)
            if not _seleccionar_por_label(selector_dia, *dia_labels):
                objetivo = "Todos" if dia in (None, 0) else str(dia)
                if estricto:
                    raise RuntimeError(f"No fue posible seleccionar el dia '{objetivo}' en el SRI.")
                return False

            selector_tipo = page.locator(selector_tipo_css)
            if not _seleccionar_en_select(page, selector_tipo_css, tipo_visible, tipo):
                if not _seleccionar(page, "Tipo de comprobante", tipo_visible):
                    if estricto:
                        raise RuntimeError(
                            f"No se pudo seleccionar el tipo de comprobante '{tipo_visible}' en Recibidos."
                        )
                    return False

            _esperar_ajax(page)
            return True
        except Exception:
            if estricto:
                raise
            return False

    _aplicar_filtros_recibidos(estricto=True)

    def _valor_actual_dia():
        try:
            return page.evaluate(
                """({selector}) => {
                    const el = document.querySelector(selector);
                    if (!el) { return null; }
                    const selected = el.options[el.selectedIndex];
                    const texto = selected ? (selected.textContent || selected.label || "").trim() : "";
                    return {
                        valor: (el.value || "").trim(),
                        texto
                    };
                }""",
                arg={"selector": "select#frmPrincipal\\:dia"},
            )
        except Exception:
            return None

    def _normalizar_dia(token: str) -> str:
        token = (token or "").strip()
        if not token:
            return ""
        if token.lower() == "todos":
            return "todos"
        if token.isdigit():
            numero = int(token)
            return "0" if numero == 0 else str(numero)
        return token.lower()

    esperados_norm = {_normalizar_dia(opcion) for opcion in dia_labels}

    def _coincide(info) -> bool:
        if not info:
            return False
        candidatos = [
            _normalizar_dia(info.get("valor")),
            _normalizar_dia(info.get("texto")),
        ]
        return any(c in esperados_norm for c in candidatos if c)

    valor_actual = _valor_actual_dia()
    if not _coincide(valor_actual):
        try:
            page.evaluate(
                """({selector, opciones}) => {
                    const normalizar = (token) => {
                        if (!token) { return ""; }
                        const trimmed = token.trim();
                        if (trimmed.toLowerCase() === "todos") { return "todos"; }
                        if (/^\\d+$/.test(trimmed)) {
                            const numero = parseInt(trimmed, 10);
                            return numero === 0 ? "0" : String(numero);
                        }
                        return trimmed.toLowerCase();
                    };
                    const el = document.querySelector(selector);
                    if (!el) { return false; }
                    const objetivos = opciones.map(normalizar);
                    for (const option of Array.from(el.options)) {
                        const lbl = normalizar(option.label || option.textContent || "");
                        const val = normalizar(option.value || "");
                        if (objetivos.includes(lbl) || objetivos.includes(val)) {
                            if (el.value !== option.value) {
                                el.value = option.value;
                                el.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                            return true;
                        }
                    }
                    return false;
                }""",
                arg={"selector": "select#frmPrincipal\\:dia", "opciones": list(dia_labels)},
            )
        except Exception:
            pass
        _esperar_ajax(page)
        valor_actual = _valor_actual_dia()

    if not _coincide(valor_actual):
        objetivo = "Todos" if dia in (None, 0) else str(dia)
        raise RuntimeError(f"No se logro confirmar el dia '{objetivo}' en el SRI.")

    _orden_tipo, _label_tipo, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
    tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
    tipo_slug = _slug_tipo(tipo_visible or tipo)
    es_retencion = _es_tipo_retencion(tipo_visible or tipo)
    es_nota_credito = _es_tipo_nota_credito(tipo_visible or tipo)
    es_nota_debito = _es_tipo_nota_debito(tipo_visible or tipo)
    es_retencion = _es_tipo_retencion(tipo_visible or tipo)
    es_nota_credito = _es_tipo_nota_credito(tipo_visible or tipo)
    es_nota_debito = _es_tipo_nota_debito(tipo_visible or tipo)
    es_nota_credito = _es_tipo_nota_credito(tipo_visible or tipo)
    es_nota_debito = _es_tipo_nota_debito(tipo_visible or tipo)
    es_retencion = _es_tipo_retencion(tipo_visible or tipo)
    es_nota_credito = _es_tipo_nota_credito(tipo_visible or tipo)
    anio_dir = f"{anio:04d}"
    mes_dir = _mes_a_texto(mes)
    try:
        dia_int = int(dia)
    except (TypeError, ValueError):
        dia_int = None
    dia_dir = "Todos" if dia_int in (None, 0) else f"{dia_int:02d}"
    fecha_token_doc = f"{anio:04d}{mes:02d}{(dia_int or 0):02d}"

    carpeta_mes = destino / tipo_dir_nombre / anio_dir / mes_dir
    carpeta_mes.mkdir(parents=True, exist_ok=True)
    carpeta_tipo = carpeta_mes
    txt_dir = carpeta_mes / "TXT"
    xml_dir = carpeta_mes / "XML"
    pdf_dir = carpeta_mes / "PDF"

    boton_consultar = None
    selectores_consultar = [
        "button#frmPrincipal\\:btnBuscar",
        "button[name='frmPrincipal:btnBuscar']",
        "#btnRecaptcha",
        "button:has-text('Consultar')",
        "input[value='Consultar']",
        "a:has-text('Consultar')",
    ]
    for selector in selectores_consultar:
        locator = page.locator(selector)
        try:
            locator.first.wait_for(state="visible", timeout=3000)
            boton_consultar = locator.first
            break
        except Exception:
            continue
    if boton_consultar is None:
        raise RuntimeError("No se encontró el botón 'Consultar' en Recibidos.")
    tabla_datos = page.locator("#frmPrincipal\\:tablaCompRecibidos_data")

    mensaje_vacio = page.locator(".ui-datatable-empty-message")
    alerta_parametros = page.locator(".ui-messages-info, .ui-messages-warn, .ui-messages-error")

    def _resultado_sin_datos(texto: str = "") -> dict:
        texto = (texto or "No se encontraron comprobantes para los filtros seleccionados.").strip()
        return {
            "estado": "sin_resultados",
            "mensaje": texto,
            "n_xml": 0,
            "n_pdf": 0,
        }

    def _texto_alerta() -> str:
        if alerta_parametros.count():
            try:
                if alerta_parametros.first.is_visible():
                    return alerta_parametros.first.inner_text().strip()
            except Exception:
                pass
        if mensaje_vacio.count():
            try:
                if mensaje_vacio.first.is_visible():
                    return mensaje_vacio.first.inner_text().strip()
            except Exception:
                pass
        return ""

    def _es_alerta_captcha(texto: str) -> bool:
        texto_norm = (texto or "").strip().lower()
        return "captcha incorrect" in texto_norm or "captcha inval" in texto_norm

    def _leer_token_recaptcha() -> str:
        try:
            return page.evaluate(
                "() => {"
                " const el = document.querySelector('textarea[name=\"g-recaptcha-response\"]');"
                " return el ? (el.value || '').trim() : '';"
                " }"
            ) or ""
        except Exception:
            return ""

    def _esperar_token_recaptcha(
        timeout: int = RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS,
        token_previo: str = "",
    ) -> tuple[bool, str]:
        min_len = 20
        try:
            page.wait_for_function(
                """({previo, minLen}) => {
                    const el = document.querySelector('textarea[name="g-recaptcha-response"]');
                    if (!el) { return true; }
                    const valor = (el.value || '').trim();
                    if (valor.length < minLen) { return false; }
                    if (previo && previo.trim().length > 0) {
                        return valor !== previo.trim();
                    }
                    return true;
                }""",
                arg={"previo": token_previo, "minLen": min_len},
                timeout=timeout,
            )
            token = _leer_token_recaptcha()
            return True, token
        except Exception:
            token = _leer_token_recaptcha()
            return False, token

    def _limpiar_estado_consulta() -> None:
        try:
            close_btn = alerta_parametros.locator(
                ".ui-messages-close, .ui-messages-close-icon, .ui-icon-close"
            ).first
            if close_btn.count():
                close_btn.click(timeout=500)
        except Exception:
            pass
        try:
            page.evaluate(
                "() => {"
                "  const campos = document.querySelectorAll("
                "    'textarea[name=\"g-recaptcha-response\"], input[name=\"g-recaptcha-response\"],"
                " textarea[id*=\"g-recaptcha-response\"], input[id*=\"g-recaptcha-response\"]'"
                "  );"
                "  campos.forEach((el) => {"
                "    el.value = '';"
                "    try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}"
                "    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) {}"
                "  });"
                "  try {"
                "    if (window.grecaptcha && grecaptcha.enterprise && typeof grecaptcha.enterprise.reset === 'function') {"
                "      grecaptcha.enterprise.reset();"
                "    }"
                "  } catch (e) {}"
                "  try { window.__auditRecibidosSubmitting = false; } catch (e) {}"
                "}"
            )
        except Exception:
            pass
        try:
            _esperar_ajax(page, timeout=1500)
        except Exception:
            pass
        try:
            page.wait_for_timeout(300)
        except Exception:
            pass

    def _forzar_submit_unico_recibidos() -> None:
        """
        Evita doble envio al consultar Recibidos.
        El portal dispara un submit inmediato y otro tras reCAPTCHA.
        Interceptamos el click para ejecutar solo el flujo con token.
        """
        try:
            ok = page.evaluate(
                """() => {
                    const btn = document.getElementById('frmPrincipal:btnBuscar');
                    if (!btn) return false;
                    if (btn.dataset.auditSingleSubmit === '1') return true;

                    const actionName = 'consulta_recibidos';
                    btn.dataset.auditSingleSubmit = '1';
                    try {
                        btn.setAttribute('onclick', 'return false;');
                    } catch (e) {}

                    const handler = function(ev) {
                        try {
                            if (ev) {
                                ev.preventDefault();
                                ev.stopPropagation();
                                if (typeof ev.stopImmediatePropagation === 'function') {
                                    ev.stopImmediatePropagation();
                                }
                            }
                        } catch (e) {}

                        try {
                            if (window.__auditRecibidosSubmitting) {
                                return false;
                            }
                            window.__auditRecibidosSubmitting = true;
                            window.setTimeout(() => { window.__auditRecibidosSubmitting = false; }, 2500);
                        } catch (e) {}

                        try { if (typeof deshabilitarBoton === 'function') deshabilitarBoton(btn); } catch (e) {}

                        try {
                            if (typeof executeRecaptcha === 'function') {
                                executeRecaptcha(actionName);
                                return false;
                            }
                        } catch (e) {}

                        try {
                            if (typeof rcBuscar === 'function') {
                                rcBuscar();
                            }
                        } catch (e) {}
                        return false;
                    };
                    btn.addEventListener('click', handler, true);
                    try {
                        btn.onclick = handler;
                    } catch (e) {}
                    return true;
                }"""
            )
            if ok:
                print("[INFO] Recibidos configurado en modo submit unico (evita doble XHR por clic).")
        except Exception as err:
            print(f"[WARN] No se pudo configurar submit unico en Recibidos: {err}")

    def _esperar_api_recaptcha_lista(timeout: int = 8000) -> bool:
        try:
            page.wait_for_function(
                """() => {
                    return !!(
                        window.grecaptcha &&
                        grecaptcha.enterprise &&
                        typeof grecaptcha.enterprise.execute === 'function' &&
                        typeof window.executeRecaptcha === 'function'
                    );
                }""",
                timeout=timeout,
            )
            return True
        except Exception:
            return False

    def _disparar_consulta_recibidos_automatica() -> str:
        try:
            modo = page.evaluate(
                """() => {
                    const actionName = 'consulta_recibidos';
                    try {
                        const campos = document.querySelectorAll(
                            'textarea[name="g-recaptcha-response"], input[name="g-recaptcha-response"], textarea[id*="g-recaptcha-response"], input[id*="g-recaptcha-response"]'
                        );
                        campos.forEach((el) => {
                            el.value = '';
                            try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
                            try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) {}
                        });
                    } catch (e) {}
                    try { window.__auditRecibidosSubmitting = false; } catch (e) {}
                    try {
                        if (window.grecaptcha && grecaptcha.enterprise && typeof grecaptcha.enterprise.reset === 'function') {
                            grecaptcha.enterprise.reset();
                        }
                    } catch (e) {}
                    try {
                        if (typeof window.executeRecaptcha === 'function') {
                            window.executeRecaptcha(actionName);
                            return 'executeRecaptcha';
                        }
                    } catch (e) {}
                    try {
                        if (typeof window.rcBuscar === 'function') {
                            window.rcBuscar();
                            return 'rcBuscar';
                        }
                    } catch (e) {}
                    try {
                        const btn = document.getElementById('frmPrincipal:btnBuscar');
                        if (btn) {
                            btn.click();
                            return 'btn.click';
                        }
                    } catch (e) {}
                    return '';
                }"""
            )
            return str(modo or "").strip()
        except Exception:
            return ""

    def _esperar_resultado_consulta(timeout: int = 300000) -> bool:
        limite = time.time() + (timeout / 1000)
        while time.time() < limite:
            try:
                if tabla_datos.is_visible():
                    return True
            except Exception:
                pass
            texto = _texto_alerta()
            if texto:
                return True
            time.sleep(0.2)
        return False

    def _rehidratar_consulta_recibidos() -> bool:
        try:
            page.goto(RECIBIDOS_DIRECT_URL, wait_until="domcontentloaded", timeout=5000)
            page.wait_for_selector(selector_ano_css, state="visible", timeout=10000)
        except Exception as err:
            print(f"[WARN] No se pudo recargar Recibidos para reintentar captcha: {err}")
            return False
        ok = _aplicar_filtros_recibidos(estricto=False)
        if not ok:
            print("[WARN] No se pudieron reaplicar filtros de Recibidos al reintentar captcha.")
            return False
        _forzar_submit_unico_recibidos()
        return ok

    _forzar_submit_unico_recibidos()

    def _intentar_consulta_recibidos(intentos: int = RECIBIDOS_CONSULTA_INTENTOS) -> bool:
        if PAUSE_BEFORE_CONSULTAR_SECONDS > 0:
            print(
                f"[INFO] Pausa {PAUSE_BEFORE_CONSULTAR_SECONDS}s antes de 'Consultar' para abrir DevTools."
            )
            time.sleep(PAUSE_BEFORE_CONSULTAR_SECONDS)
        if PAUSE_BEFORE_CONSULTAR:
            try:
                input("[INFO] Abre DevTools (F12) y presiona Enter para continuar...")
            except Exception:
                pass
        if MANUAL_CONSULTA_RECIBIDOS:
            for intento in range(1, intentos + 1):
                _notificar_usuario_accion(
                    f"[ACCION] Da clic manual en 'Consultar' (Recibidos). "
                    f"Intento {intento}/{intentos}."
                )
                ok_manual = _esperar_resultado_consulta(timeout=300000)
                alerta_manual = _texto_alerta()
                if ok_manual:
                    return True
                if alerta_manual and _es_alerta_captcha(alerta_manual):
                    print(
                        f"[WARN] Captcha incorrecto tras clic manual "
                        f"({intento}/{intentos})."
                    )
                    if intento >= intentos:
                        return ok_manual
                    _limpiar_estado_consulta()
                    if RECIBIDOS_REHIDRATAR_ON_CAPTCHA and intento + 1 >= RECIBIDOS_REHIDRATAR_DESDE_INTENTO:
                        _rehidratar_consulta_recibidos()
                    if RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                        espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                        print(f"[INFO] Esperando {espera:.1f}s antes de reintento manual.")
                        time.sleep(espera)
                    continue
                if alerta_manual:
                    # Hay mensaje del portal (ej. sin resultados), salir para que lo procese la capa superior.
                    return True
                print(
                    f"[WARN] Recibidos sin tabla ni alerta tras clic manual "
                    f"({intento}/{intentos})."
                )
                if intento >= intentos:
                    return False
                _limpiar_estado_consulta()
                if RECIBIDOS_REHIDRATAR_ON_CAPTCHA and intento + 1 >= RECIBIDOS_REHIDRATAR_DESDE_INTENTO:
                    _rehidratar_consulta_recibidos()
                if RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                    espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                    print(f"[INFO] Esperando {espera:.1f}s antes de reintento manual.")
                    time.sleep(espera)
            return False
        token_previo = _leer_token_recaptcha()
        for intento in range(1, intentos + 1):
            inicio_intento = time.perf_counter()
            if intento > 1:
                _limpiar_estado_consulta()
            if RECIBIDOS_REHIDRATAR_ON_CAPTCHA and intento >= RECIBIDOS_REHIDRATAR_DESDE_INTENTO:
                _rehidratar_consulta_recibidos()
            _esperar_api_recaptcha_lista(timeout=7000)
            if RECIBIDOS_AUTO_PRE_EXECUTE_MS > 0:
                try:
                    page.wait_for_timeout(RECIBIDOS_AUTO_PRE_EXECUTE_MS)
                except Exception:
                    pass
            modo_disparo = _disparar_consulta_recibidos_automatica()
            if RECIBIDOS_AUTO_POST_EXECUTE_MS > 0:
                try:
                    page.wait_for_timeout(RECIBIDOS_AUTO_POST_EXECUTE_MS)
                except Exception:
                    pass
            token_ok, token_actual = _esperar_token_recaptcha(
                timeout=RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS,
                token_previo=token_previo,
            )
            print(
                f"[INFO] Recibidos intento {intento}/{intentos}: "
                f"modo={modo_disparo or 'desconocido'}, token_ok={token_ok}, "
                f"token_len={len(token_actual or '')}"
            )
            if not modo_disparo:
                try:
                    boton_consultar.first.scroll_into_view_if_needed()
                    boton_consultar.first.click(delay=80)
                    modo_disparo = "fallback-click"
                except Exception:
                    pass
            try:
                page.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
            _esperar_resultado_consulta(timeout=RECIBIDOS_AUTO_RESULT_TIMEOUT_MS)
            alerta_post = _texto_alerta()
            if alerta_post and _es_alerta_captcha(alerta_post):
                dur = time.perf_counter() - inicio_intento
                print(
                    f"[WARN] Captcha incorrecto en intento {intento}/{intentos} "
                    f"({dur:.2f}s)."
                )
                token_previo = token_actual or token_previo
                if intento < intentos and RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                    espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                    print(f"[INFO] Esperando {espera:.1f}s antes de reintentar Recibidos.")
                    time.sleep(espera)
                continue
            try:
                if tabla_datos.is_visible():
                    dur = time.perf_counter() - inicio_intento
                    print(
                        f"[INFO] Recibidos intento {intento}/{intentos} exitoso "
                        f"({dur:.2f}s)."
                    )
                    return True
            except Exception:
                pass
            if alerta_post:
                return True
            token_previo = token_actual or token_previo
            if intento < intentos and RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                print(f"[INFO] Sin tabla tras intento {intento}/{intentos}. Espera {espera:.1f}s.")
                time.sleep(espera)
        return False

    resultado_listo = _intentar_consulta_recibidos(intentos=RECIBIDOS_CONSULTA_INTENTOS)
    if not resultado_listo:
        try:
            tabla_datos.wait_for(state="visible", timeout=2000)
        except Exception:
            pass

    alerta_texto = _texto_alerta()
    if alerta_texto:
        if _es_alerta_captcha(alerta_texto):
            _limpiar_estado_consulta()
            resultado_listo = _intentar_consulta_recibidos(intentos=RECIBIDOS_CONSULTA_INTENTOS)
            try:
                if not resultado_listo:
                    tabla_datos.wait_for(state="visible", timeout=2000)
                alerta_texto = ""
            except Exception:
                return _resultado_sin_datos(alerta_texto)
        else:
            return _resultado_sin_datos(alerta_texto)
    try:
        tabla_datos.wait_for(state="visible", timeout=180000)
    except PlaywrightTimeoutError:
        alerta_texto = _texto_alerta()
        if alerta_texto:
            return _resultado_sin_datos(alerta_texto)
        try:
            if _es_alerta_captcha(_texto_alerta()):
                return _resultado_sin_datos("Captcha incorrecta al consultar comprobantes recibidos.")
        except Exception:
            pass
        raise
    except Exception:
        alerta_texto = _texto_alerta()
        if alerta_texto:
            return _resultado_sin_datos(alerta_texto)
        raise

    alerta_texto = _texto_alerta()
    if alerta_texto:
        if _es_alerta_captcha(alerta_texto):
            return _resultado_sin_datos(alerta_texto)
        else:
            return _resultado_sin_datos(alerta_texto)
    n_xml = 0
    n_pdf = 0
    txt_path = None
    pdf_report_rows = []

    if "XML" in formatos:
        link_reporte = page.locator("a#frmPrincipal\\:lnkTxtlistado")
        if link_reporte.count():
            try:
                with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as descarga_info:
                    link_reporte.first.click(no_wait_after=True)
                descarga = descarga_info.value
                txt_dir.mkdir(parents=True, exist_ok=True)
                txt_nombre = descarga.suggested_filename or f"recibidos_{anio}_{mes:02d}.txt"
                txt_path = txt_dir / txt_nombre
                descarga.save_as(str(txt_path))
            except Exception as err:
                print(f"[WARN] No se pudo descargar el reporte TXT/XML: {err}")
        else:
            print("[WARN] No se encontro el enlace 'Descargar reporte' para XML.")

    descargar_xml = "XML" in formatos
    descargar_pdf = "PDF" in formatos
    descargar_xml_para_reporte = descargar_xml or descargar_pdf

    if descargar_xml_para_reporte:
        xml_dir.mkdir(parents=True, exist_ok=True)
    if descargar_pdf:
        pdf_dir.mkdir(parents=True, exist_ok=True)

    xml_temp_paths = []
    if descargar_xml_para_reporte or descargar_pdf:
        pagina = 1
        lote_size = 10
        def _obtener_viewstate_actual() -> str:
            try:
                return page.locator("input[name='javax.faces.ViewState']").first.get_attribute("value") or ""
            except Exception:
                return ""

        def _extraer_clave_fila(celdas) -> str:
            def _buscar_clave(texto: str) -> str:
                texto = (texto or "").strip()
                if not texto:
                    return ""
                match = re.search(r"\d{49}", texto)
                if match:
                    return match.group(0)
                solo_digitos = re.sub(r"\D", "", texto)
                return solo_digitos if len(solo_digitos) == 49 else ""

            try:
                total = celdas.count()
            except Exception:
                total = 0

            for idx_celda in range(min(total, 6)):
                try:
                    texto = celdas.nth(idx_celda).inner_text().strip()
                except Exception:
                    continue
                clave = _buscar_clave(texto)
                if clave:
                    return clave

            try:
                texto_fila = " ".join(
                    celdas.nth(i).inner_text().strip() for i in range(min(total, 8))
                )
            except Exception:
                return ""
            return _buscar_clave(texto_fila)
        while True:
            _check_cancel("recibidos_pagina")
            page_inicio = time.perf_counter()
            view_state = _obtener_viewstate_actual()
            filas = tabla_datos.locator("tr")
            total_filas = filas.count()
            lote_inicio = time.perf_counter()
            lote_contador = 0
            lote_xml_ok = 0
            lote_pdf_ok = 0
            for idx in range(total_filas):
                _check_cancel("recibidos_fila")
                fila = filas.nth(idx)
                celdas = fila.locator("td")
                if not celdas.count():
                    continue
                razon_texto = celdas.nth(1).inner_text().strip()
                bloques = [segmento.strip() for segmento in razon_texto.splitlines() if segmento.strip()]
                razon_social = bloques[-1] if bloques else f"documento_{pagina}_{idx+1}"
                nombre_base = _nombre_documento_mes(tipo_slug, fecha_token_doc, razon_social)

                xml_guardado = False
                xml_path_report = None
                if descargar_xml_para_reporte and not xml_guardado:
                    xml_link_directo = fila.locator("a[id$=':lnkXml']")
                    if not xml_link_directo.count():
                        xml_link_directo = fila.locator("a[title*='xml' i], button[title*='xml' i]")
                    if not xml_link_directo.count():
                        icono_xml = fila.locator("img[title*='xml' i], img[alt*='xml' i]")
                        if icono_xml.count():
                            contenedor = icono_xml.first.locator("xpath=ancestor::a[1] | xpath=ancestor::button[1]")
                            if not contenedor.count():
                                contenedor = icono_xml.first.locator("xpath=ancestor::span[1]")
                            if contenedor.count():
                                xml_link_directo = contenedor.first
                    if xml_link_directo and xml_link_directo.count():
                        destino_base_directo = xml_dir / nombre_base
                        destino_xml_directo = _resolver_destino_unico(destino_base_directo, ".xml")
                        resultado_directo = _guardar_xml_desde_enlace(page, xml_link_directo, destino_xml_directo)
                        if resultado_directo:
                            xml_path_report = resultado_directo
                            if descargar_xml:
                                n_xml += 1
                            else:
                                xml_temp_paths.append(xml_path_report)
                            xml_guardado = True
                            lote_xml_ok += 1

                if descargar_xml_para_reporte and not xml_guardado:
                    xml_link = None
                    selectores_xml = [
                        "a[id$=':lnkXml']",
                        "a[id*='lnkXml']",
                        "a[title*='xml' i]",
                        "button[id*='lnkXml']",
                        "button[title*='xml' i]",
                    ]
                    for selector in selectores_xml:
                        posible = fila.locator(selector)
                        if posible.count():
                            xml_link = posible.first
                            break
                    if not xml_link:
                        icono_xml = fila.locator("img[title*='xml' i], img[alt*='xml' i]")
                        if icono_xml.count():
                            contenedor = icono_xml.first.locator("xpath=ancestor::a[1]")
                            if contenedor.count():
                                xml_link = contenedor.first
                    if xml_link:
                        destino_xml = xml_dir / f"{nombre_base}.xml"
                        sufijo_xml = 1
                        while destino_xml.exists():
                            destino_xml = xml_dir / f"{nombre_base}_{sufijo_xml}.xml"
                            sufijo_xml += 1
                        try:
                            with page.expect_download(timeout=DOWNLOAD_TIMEOUT) as descarga_info:
                                xml_link.click(no_wait_after=True)
                            descarga_xml = descarga_info.value
                            sugerido = descarga_xml.suggested_filename or destino_xml.name
                            extension = Path(sugerido).suffix or ".xml"
                            destino_final = destino_xml.with_suffix(extension)
                            sufijo_xml = 1
                            while destino_final.exists():
                                destino_final = destino_xml.with_name(f"{destino_xml.stem}_{sufijo_xml}{extension}")
                                sufijo_xml += 1
                            descarga_xml.save_as(str(destino_final))
                            xml_path_report = destino_final
                            if descargar_xml:
                                n_xml += 1
                            else:
                                xml_temp_paths.append(xml_path_report)
                            xml_guardado = True
                            lote_xml_ok += 1
                        except Exception as err:
                            print(f"[WARN] No se pudo descargar XML para '{razon_social}': {err}")

                usar_xml_reporte = False
                if descargar_pdf and xml_path_report:
                    try:
                        if es_retencion:
                            datos_xml = _extraer_datos_xml_retencion(xml_path_report)
                        else:
                            datos_xml = _extraer_datos_xml_pdf_report(xml_path_report)
                        pdf_report_rows.append(datos_xml)
                        usar_xml_reporte = True
                    except Exception as err:
                        print(f"[WARN] No se pudo procesar XML para reporte: {err}")

                if descargar_pdf:
                    destino_pdf = pdf_dir / f"{nombre_base}.pdf"
                    link_id = f"frmPrincipal:tablaCompRecibidos:{idx}:lnkPdf"
                    resultado_pdf = None
                    if view_state:
                        resultado_pdf = _descargar_pdf_recibidos_post_con_viewstate(
                            page, link_id, view_state, destino_pdf
                        )
                    if not resultado_pdf:
                        view_state = _obtener_viewstate_actual() or view_state
                        if view_state:
                            resultado_pdf = _descargar_pdf_recibidos_post_con_viewstate(
                                page, link_id, view_state, destino_pdf
                            )
                    if not resultado_pdf:
                        link_pdf = fila.locator("a[id$=':lnkPdf']")
                        if not link_pdf.count():
                            link_pdf = fila.locator("a[title*='pdf' i], button[title*='pdf' i]")
                        if link_pdf.count():
                            resultado_pdf = _guardar_pdf_desde_jsf(page, link_pdf.first, destino_pdf)
                            if not resultado_pdf:
                                resultado_pdf = _guardar_pdf_desde_enlace(page, link_pdf.first, destino_pdf)
                    if resultado_pdf:
                        n_pdf += 1
                        lote_pdf_ok += 1
                        if resultado_pdf.suffix.lower() == ".pdf" and _es_archivo_pdf(resultado_pdf):
                            if not usar_xml_reporte:
                                datos_pdf = _extraer_datos_pdf_por_tipo_layout_first(
                                    resultado_pdf,
                                    es_retencion=es_retencion,
                                    es_nota_credito=es_nota_credito,
                                    es_nota_debito=es_nota_debito,
                                )
                                pdf_report_rows.append(datos_pdf)
                    else:
                        print(f"[WARN] No se pudo descargar PDF para '{razon_social}': no se obtuvo archivo.")

                lote_contador += 1
                if lote_contador >= lote_size or idx == total_filas - 1:
                    duracion_lote = time.perf_counter() - lote_inicio
                    print(
                        f"[INFO] Pag {pagina} lote {((idx // lote_size) + 1)}: "
                        f"{lote_contador} filas, XML {lote_xml_ok}, PDF {lote_pdf_ok}, "
                        f"{duracion_lote:.2f}s"
                    )
                    lote_inicio = time.perf_counter()
                    lote_contador = 0
                    lote_xml_ok = 0
                    lote_pdf_ok = 0

            duracion_pagina = time.perf_counter() - page_inicio
            print(f"[INFO] Pag {pagina} completa: {total_filas} filas en {duracion_pagina:.2f}s")

            boton_siguiente = page.locator("span.ui-paginator-next:not(.ui-state-disabled)")
            if boton_siguiente.count():
                boton_siguiente.first.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=1000)
                except Exception:
                    pass
                time.sleep(0.2)
                pagina += 1
                continue
            break

    reporte_pdf_path = None
    if descargar_pdf and n_pdf > 0 and pdf_report_rows:
        fecha_slug = f"{anio:04d}{mes:02d}"
        if dia_dir != "Todos":
            fecha_slug = f"{fecha_slug}{dia_dir}"
        pdf_dir.mkdir(parents=True, exist_ok=True)
        pdf_report_path = pdf_dir / f"recibidos_reporte_pdf_{tipo_slug}_{fecha_slug}.xlsx"
        if pdf_report_path.exists():
            try:
                pdf_report_path.unlink()
            except PermissionError:
                sufijo_pdf = 1
                while True:
                    candidato = pdf_dir / f"recibidos_reporte_pdf_{tipo_slug}_{fecha_slug}_{sufijo_pdf}.xlsx"
                    if not candidato.exists():
                        pdf_report_path = candidato
                        break
                    sufijo_pdf += 1
        if es_retencion:
            if _guardar_reporte_pdf_retencion_excel(pdf_report_rows, pdf_report_path):
                reporte_pdf_path = pdf_report_path
            else:
                print("[WARN] No se pudo construir el reporte PDF de retenciones (recibidos).")
        elif _guardar_reporte_pdf_excel(pdf_report_rows, pdf_report_path):
            reporte_pdf_path = pdf_report_path
        else:
            print("[WARN] No se pudo construir el reporte PDF de recibidos.")
    resultado = {
        "estado": "ok",
        "n_xml": n_xml,
        "n_pdf": n_pdf,
        "carpeta_tipo": str(carpeta_mes),
        "tipo_slug": tipo_slug,
        "tipo_visible": tipo_visible,
        "txt_dir": str(txt_dir),
        "xml_dir": str(xml_dir),
        "pdf_dir": str(pdf_dir),
    }
    if reporte_pdf_path:
        resultado["reporte_pdf"] = str(reporte_pdf_path)
    if txt_path:
        resultado["txt"] = str(txt_path)

    if es_retencion and not descargar_xml and xml_temp_paths:
        for xml_tmp in xml_temp_paths:
            try:
                Path(xml_tmp).unlink(missing_ok=True)
            except Exception:
                pass
    return resultado

# ============================================================
# ?? LECTURA DE TABLA PARA COMPROBANTES EMITIDOS (sin TXT)
# ============================================================


def _guardar_reporte_emitidos_excel(df_emitidos: pd.DataFrame, excel_path: Path, titulo: str = "EMITIDOS") -> bool:
    if df_emitidos.empty:
        return False

    columns_order = [
        "COMPROBANTE",
        "SERIE_COMPROBANTE",
        "CLAVE_ACCESO",
        "FECHA_AUTORIZACION",
        "FECHA_EMISION",
        "VALOR_SIN_IMPUESTOS",
        "IVA",
        "IMPORTE_TOTAL",
    ]
    for col in columns_order:
        if col not in df_emitidos.columns:
            df_emitidos[col] = ""
    df_emitidos = df_emitidos[columns_order].copy()

    def _coerce_decimal_value(val):
        if isinstance(val, (int, float)):
            return float(val)
        parsed = _parse_decimal(val) if isinstance(val, str) else None
        return parsed if parsed is not None else val

    def _coerce_datetime_value(val):
        if isinstance(val, datetime):
            return val
        parsed = _parse_datetime_local(val) if isinstance(val, str) else None
        return parsed if parsed is not None else val

    for columna in ["VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"]:
        df_emitidos[columna] = df_emitidos[columna].apply(_coerce_decimal_value)
    for columna in ["FECHA_AUTORIZACION"]:
        df_emitidos[columna] = df_emitidos[columna].apply(_coerce_datetime_value)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        sheet_name = "Emitidos"
        df_emitidos.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1)
        ws = writer.sheets[sheet_name]
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns_order))
        titulo_cell = ws.cell(row=1, column=1, value=titulo)
        titulo_cell.font = Font(bold=True, size=14)
        titulo_cell.alignment = Alignment(horizontal="center", vertical="center")

        header_fill = PatternFill("solid", fgColor="305496")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[2]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.freeze_panes = "A3"

        text_columns = {"COMPROBANTE", "SERIE_COMPROBANTE", "CLAVE_ACCESO", "FECHA_EMISION"}
        numeric_columns = {"VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"}
        date_columns = {"FECHA_AUTORIZACION"}

        for idx, column in enumerate(columns_order, start=1):
            max_len = len(column)
            for cell_tuple in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=idx, max_col=idx):
                cell = cell_tuple[0]
                valor = cell.value
                if valor is None:
                    continue
                if isinstance(valor, datetime):
                    texto_len = len(valor.strftime("%d/%m/%Y %H:%M"))
                else:
                    texto_len = len(str(valor))
                if texto_len > max_len:
                    max_len = texto_len
            ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 45)

        for columna in text_columns:
            if columna not in columns_order:
                continue
            col_idx = columns_order.index(columna) + 1
            for cell_tuple in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                celda = cell_tuple[0]
                celda.number_format = "@"
                celda.alignment = Alignment(horizontal="left", vertical="center")

        for columna in numeric_columns:
            if columna not in columns_order:
                continue
            col_idx = columns_order.index(columna) + 1
            for cell_tuple in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                celda = cell_tuple[0]
                if isinstance(celda.value, (int, float)):
                    celda.number_format = "#,##0.00"
                    celda.alignment = Alignment(horizontal="right", vertical="center")

        for columna in date_columns:
            if columna not in columns_order:
                continue
            col_idx = columns_order.index(columna) + 1
            for cell_tuple in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                celda = cell_tuple[0]
                if isinstance(celda.value, datetime):
                    celda.number_format = "dd/mm/yyyy hh:mm"
                    celda.alignment = Alignment(horizontal="center", vertical="center")

    return True

def _flujo_emitidos(
    page,
    destino: Path,
    fecha_emision: Optional[str],
    tipo: str,
    estado_autorizacion: Optional[str],
    establecimiento: Optional[str],
    punto_emision: Optional[str],
    formatos: list,
    ruc_emisor: Optional[str] = None,
):
    _check_cancel("inicio_emitidos")
    es_retencion = False
    es_nota_credito = False
    es_nota_debito = False
    fecha_emision = (fecha_emision or "").strip()
    estado_autorizacion = (estado_autorizacion or "").strip()
    establecimiento = (establecimiento or "").strip()
    punto_emision = (punto_emision or "").strip()
    if punto_emision:
        punto_emision = "".join(ch for ch in punto_emision if ch.isdigit())[:3]

    formatos_norm = [
        (fmt or "").strip().upper() for fmt in (formatos or []) if isinstance(fmt, str)
    ]
    descargar_pdf = "PDF" in formatos_norm
    descargar_xml = "XML" in formatos_norm
    selecciono_pdf = descargar_pdf
    selecciono_xml = descargar_xml

    tipo_visible = TIPOS_MAP.get(tipo, tipo)
    es_retencion = _es_tipo_retencion(tipo_visible or tipo)
    es_nota_credito = _es_tipo_nota_credito(tipo_visible or tipo)
    es_nota_debito = _es_tipo_nota_debito(tipo_visible or tipo)
    es_factura_emitida = _es_tipo_factura(tipo_visible or tipo)
    es_liquidacion_compra = _es_tipo_liquidacion_compra(tipo_visible or tipo)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=1000)
    except Exception:
        pass

    if not _seleccionar_en_select(
        page,
        "select#frmPrincipal\:cmbTipoComprobante",
        tipo_visible,
        tipo,
    ):
        if not _seleccionar(page, "Tipo de comprobante", tipo_visible):
            print(f"[WARN] No se pudo seleccionar el tipo de comprobante '{tipo_visible}' en Emitidos.")

    estado_visible = ESTADOS_EMITIDOS_MAP.get(estado_autorizacion, estado_autorizacion)
    if estado_visible:
        if not _seleccionar_en_select(
            page,
            "select#frmPrincipal\:cmbEstadoAutorizacion",
            estado_visible,
        ):
            etiquetas_estado = [
                "Estado autorizacion",
                "Estado autorizacion",
                "Estado",
            ]
            seleccionado = any(_seleccionar(page, etiqueta, estado_visible) for etiqueta in etiquetas_estado)
            if not seleccionado:
                print(
                    f"[WARN] No se pudo seleccionar el estado de autorizacion '{estado_visible}' en Emitidos."
                )
    estado_norm = (
        unicodedata.normalize("NFKD", estado_visible).encode("ascii", "ignore").decode("ascii").lower()
        if estado_visible
        else ""
    )
    modo_no_autorizados = "no autoriz" in estado_norm
    if modo_no_autorizados:
        if selecciono_pdf and not selecciono_xml:
            _notificar_usuario_accion(
                "[INFO] En 'No Autorizados' no hay PDF disponible. Marca 'XML' para descargar."
            )
        descargar_pdf = False

    omitir_soap_xml = _debe_omitir_soap_xml(fecha_emision, descargar_xml)

    if fecha_emision:
        fecha_selector = "input#frmPrincipal\:calendarFechaDesde_input"
        fecha_ok = False
        try:
            fecha_loc = page.locator(fecha_selector)
            if fecha_loc.count():
                fecha_loc.first.fill("")
                fecha_loc.first.fill(fecha_emision)
                try:
                    fecha_loc.first.dispatch_event("input")
                    fecha_loc.first.dispatch_event("change")
                except Exception:
                    pass
                fecha_ok = True
        except Exception:
            fecha_ok = False
        if not fecha_ok:
            if not _rellenar_input_por_label(
                page,
                ["Fecha de emision", "Fecha emision", "Fecha"],
                fecha_emision,
                EMITIDOS_FECHA_SELECTORS,
            ):
                print(f"[WARN] No se pudo completar la fecha de emision con '{fecha_emision}' en Emitidos.")

    est_valor = establecimiento or ""
    if not est_valor or est_valor.lower() == "todos":
        est_valor = "Todos"
        if not _seleccionar_en_select(
            page,
            "select#frmPrincipal\:cmbEstablecimiento",
            est_valor,
        ):
            _rellenar_input_por_label(
                page,
                ["Establecimiento"],
                est_valor,
                EMITIDOS_ESTABLECIMIENTO_SELECTORS,
            )
    else:
        if not _seleccionar_en_select(
            page,
            "select#frmPrincipal\:cmbEstablecimiento",
            est_valor,
        ):
            if not _rellenar_input_por_label(
                page,
                ["Establecimiento"],
                est_valor,
                EMITIDOS_ESTABLECIMIENTO_SELECTORS,
            ):
                print(f"[WARN] No se pudo establecer el establecimiento '{est_valor}' en Emitidos.")

    punto_valor = punto_emision
    if est_valor and est_valor.lower() != "todos" and not punto_valor:
        raise RuntimeError("El punto de emision es obligatorio cuando se especifica un establecimiento.")
    if punto_valor:
        punto_selector = "input#frmPrincipal\:txtPuntoEmision"
        punto_ok = False
        try:
            punto_loc = page.locator(punto_selector)
            if punto_loc.count():
                punto_loc.first.fill("")
                punto_loc.first.fill(punto_valor)
                punto_ok = True
        except Exception:
            punto_ok = False
        if not punto_ok:
            if not _rellenar_input_por_label(
                page,
                ["Punto de emision", "Punto emision", "Punto"],
                punto_valor,
                EMITIDOS_PUNTO_SELECTORS,
            ):
                print(f"[WARN] No se pudo establecer el punto de emision '{punto_valor}' en Emitidos.")

    estado_nombre = (estado_visible or "Sin Estado").strip() or "Sin Estado"
    estado_normalizado = unicodedata.normalize("NFKD", estado_nombre).encode("ascii", "ignore").decode("ascii")
    estado_slug = re.sub(r"[^A-Za-z0-9]+", "_", estado_normalizado).strip("_") or "Sin_Estado"

    fecha_dt = _parse_datetime_local(fecha_emision) if fecha_emision else None
    if not fecha_dt:
        fecha_dt = datetime.now()
    anio_dir = f"{fecha_dt.year:04d}"
    mes_dir = _mes_a_texto(fecha_dt.month)
    dia_dir = f"{fecha_dt.day:02d}"
    fecha_token_doc = f"{fecha_dt.year:04d}{fecha_dt.month:02d}{fecha_dt.day:02d}"

    _orden_tipo, _label_tipo, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
    tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
    tipo_slug = _slug_tipo(tipo_visible or tipo)

    carpeta_estado = destino / estado_slug / tipo_dir_nombre / anio_dir / mes_dir
    carpeta_estado.mkdir(parents=True, exist_ok=True)
    carpeta_tipo = carpeta_estado
    xml_dir = carpeta_tipo / "XML"
    descargar_xml_para_reporte = descargar_xml or descargar_pdf
    if descargar_xml_para_reporte:
        xml_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = carpeta_tipo / "PDF"
    if descargar_pdf:
        pdf_dir.mkdir(parents=True, exist_ok=True)

    if not _click_consultar_emitidos(page):
        try:
            page.keyboard.press("Enter")
        except Exception:
            pass

            try:
                page.wait_for_load_state("networkidle", timeout=1000)
            except Exception:
                pass
    time.sleep(0.2)

    tabla_emitidos = page.locator("#frmPrincipal\\:tablaCompEmitidos_data")
    es_rechazado = False
    try:
        tabla_emitidos.wait_for(state="visible", timeout=1000)
    except Exception:
        pass
    if not tabla_emitidos.count():
        tabla_emitidos = page.locator("#frmPrincipal\\:tablaCompRechazados_data")
        if tabla_emitidos.count():
            es_rechazado = True
            try:
                tabla_emitidos.wait_for(state="visible", timeout=1000)
            except Exception:
                pass

    data = []
    if not es_rechazado:
        html = page.content()
        rows = re.findall(
            r"<tr[^>]*>\s*(.*?)\s*</tr>",
            html,
            flags=re.DOTALL
        )

        encabezado_textos = ["EMITIDOS", "", "CLA VE ACCESO"]  # placeholder; se ajustara con row data
        for r in rows:
            cols = re.findall(r"<td[^>]*>(.*?)</td>", r, flags=re.DOTALL)
            textos = [re.sub("<.*?>", "", c).strip() for c in cols]
            if len(textos) < 3:
                continue
            fecha_emision_col = textos[0]
            comprobante_raw = textos[1]
            tipo_detectado = _extraer_tipo_documento(comprobante_raw)
            if tipo_detectado and not _coincide_tipo_documental(tipo_visible or tipo, tipo_detectado):
                continue
            partes_tipo = comprobante_raw.split()
            comprobante = partes_tipo[0] if partes_tipo else ""
            serie = " ".join(partes_tipo[1:]) if len(partes_tipo) > 1 else ""
            clave = textos[2]
            fecha_autorizacion = textos[3] if len(textos) > 3 else ""
            valor_sin_impuestos = textos[5] if len(textos) > 5 else ""
            iva_val = textos[6] if len(textos) > 6 else ""
            importe_total = textos[7] if len(textos) > 7 else ""
            if valor_sin_impuestos and not re.search(r"\d", valor_sin_impuestos):
                valor_sin_impuestos = ""
            if iva_val and not re.search(r"\d", iva_val):
                iva_val = ""
            if importe_total and not re.search(r"\d", importe_total):
                importe_total = ""

            data.append({
                "COMPROBANTE": comprobante,
                "SERIE_COMPROBANTE": serie,
                "CLAVE_ACCESO": clave,
                "FECHA_AUTORIZACION": fecha_autorizacion,
                "FECHA_EMISION": fecha_emision_col,
                "VALOR_SIN_IMPUESTOS": valor_sin_impuestos,
                "IVA": iva_val,
                "IMPORTE_TOTAL": importe_total,
            })
    else:
        filas_rech = tabla_emitidos.locator("tr")
        total_rech = filas_rech.count()
        for idx in range(total_rech):
            fila = filas_rech.nth(idx)
            celdas = fila.locator("td")
            if not celdas.count():
                continue
            try:
                tipo_serie_texto = celdas.nth(1).inner_text().strip()
            except Exception:
                tipo_serie_texto = ""
            tipo_detectado = _extraer_tipo_documento(tipo_serie_texto)
            if tipo_detectado and not _coincide_tipo_documental(tipo_visible or tipo, tipo_detectado):
                continue
            try:
                clave_texto = celdas.nth(2).inner_text().strip()
            except Exception:
                clave_texto = ""
            try:
                fecha_aut = celdas.nth(3).inner_text().strip()
            except Exception:
                fecha_aut = ""

            data.append({
                "COMPROBANTE": tipo_serie_texto,
                "SERIE_COMPROBANTE": "",
                "CLAVE_ACCESO": clave_texto,
                "FECHA_AUTORIZACION": fecha_aut,
                "FECHA_EMISION": fecha_emision,
                "VALOR_SIN_IMPUESTOS": "",
                "IVA": "",
                "IMPORTE_TOTAL": "",
            })

    n_pdf = 0
    n_xml = 0
    pdf_report_rows = []

    info_base = {
        "carpeta_tipo": str(carpeta_tipo),
        "tipo_slug": tipo_slug,
        "tipo_visible": tipo_visible,
        "fecha_filtro": fecha_emision,
        "estado_autorizacion": estado_visible,
        "establecimiento": est_valor,
        "punto_emision": punto_valor,
        "carpeta_estado": str(carpeta_estado),
        "xml_dir": str(xml_dir) if descargar_xml else "",
        "pdf_dir": str(pdf_dir) if descargar_pdf else "",
        "n_xml": n_xml,
        "n_pdf": n_pdf,
        "reporte_xml": "",
        "reporte_pdf": "",
    }

    if not data:
        # Si no hay datos, limpiamos carpetas vacías creadas para el día
        def _limpiar_carpetas_vacias(carpeta_base: Path, stop_at: Path) -> None:
            try:
                current = carpeta_base
                stop_at_resolved = stop_at.resolve()
                while True:
                    if not current.exists():
                        break
                    try:
                        current.rmdir()
                    except OSError:
                        break
                    if current.resolve() == stop_at_resolved:
                        break
                    parent = current.parent
                    if parent == current:
                        break
                    current = parent
            except Exception:
                pass

        _limpiar_carpetas_vacias(carpeta_tipo, destino / estado_slug)
        info_base.update({
            "estado": "sin_resultados",
            "mensaje": "No se encontraron filas en la tabla",
            "n_registros": 0,
        })
        return info_base

    xml_temp_paths = []
    if descargar_pdf or descargar_xml_para_reporte:
        try:
            tabla_emitidos.wait_for(state="visible", timeout=1000)
        except Exception:
            pass
        pagina = 1
        lote_size = 10
        claves_guardadas = set()
        pdf_esperados = set()
        pdf_descargados = set()
        xml_esperados = set()
        xml_descargados = set()
        request_context = page.context.request
        payload_base = _obtener_form_base_emitidos(page)
        def _pdf_report_incompleto(datos: dict, min_campos: int = 4) -> bool:
            if not isinstance(datos, dict) or not datos:
                return True
            ignorar = {"numeroComprobante", "fechaEmision", "claveAcceso"}
            conteo = 0
            for key, valor in datos.items():
                if key in ignorar:
                    continue
                if valor in (None, "", "No Disponible", "N/A", 0, "0"):
                    continue
                conteo += 1
            return conteo < min_campos

        def _obtener_viewstate_actual() -> str:
            try:
                return page.locator("input[name='javax.faces.ViewState']").first.get_attribute("value") or ""
            except Exception:
                return ""

        while True:
            _check_cancel("emitidos_pagina")
            page_inicio = time.perf_counter()
            view_state = _obtener_viewstate_actual()
            filas_snapshot = _snapshot_filas_emitidos(page, es_rechazado=es_rechazado)
            filas = tabla_emitidos.locator("tr")
            total_filas = len(filas_snapshot)
            lote_inicio = time.perf_counter()
            lote_contador = 0
            lote_xml_ok = 0
            lote_pdf_ok = 0
            for idx, fila_data in enumerate(filas_snapshot):
                _check_cancel("emitidos_fila")
                fila = filas.nth(idx)
                total_celdas = int(fila_data.get("total_celdas") or 0)
                if total_celdas < 2:
                    continue
                tipo_serie_texto = str(fila_data.get("tipo_serie_texto") or "").strip()
                tipo_detectado = _extraer_tipo_documento(tipo_serie_texto)
                if tipo_detectado and not _coincide_tipo_documental(tipo_visible or tipo, tipo_detectado):
                    print(
                        f"[WARN] Se omitio una fila de Emitidos porque corresponde a '{tipo_detectado}' y no a '{tipo_visible or tipo}'."
                    )
                    continue
                clave_texto = str(fila_data.get("clave_texto") or "").strip()
                razon_texto = str(fila_data.get("razon_texto") or "").strip()
                fecha_aut_texto = str(fila_data.get("fecha_aut_texto") or "").strip()
                valor_sin_imp_texto = str(fila_data.get("valor_sin_imp_texto") or "").strip()
                iva_texto = str(fila_data.get("iva_texto") or "").strip()
                importe_total_texto = str(fila_data.get("importe_total_texto") or "").strip()
                row_id = clave_texto or f"{pagina}:{idx}:{tipo_serie_texto}"

                tipo_serie_completo = " ".join(
                    fragment for fragment in [tipo_serie_texto, clave_texto] if fragment
                ).strip()
                tipo_slug_archivo = _slug_tipo(tipo_detectado or tipo_visible or tipo) or tipo_slug
                nombre_base_pdf = _nombre_documento_mes(
                    tipo_slug_archivo,
                    fecha_token_doc,
                    tipo_serie_completo or razon_texto or f"emitido_{pagina}_{idx+1}",
                )
                xml_path_report = None

                if es_rechazado:
                    try:
                        if descargar_xml_para_reporte:
                            if clave_texto:
                                xml_esperados.add(row_id)
                                try:
                                    resultado_xml = None
                                    for intento_xml in range(2):
                                        resultado_xml = _descargar_xml_emitido_por_clave(
                                            request_context,
                                            clave_texto,
                                            xml_dir,
                                            nombre_base_pdf,
                                            claves_guardadas,
                                        )
                                        if resultado_xml:
                                            break
                                        time.sleep(0.15)
                                    if resultado_xml:
                                        xml_path_report = resultado_xml
                                        xml_descargados.add(row_id)
                                        if descargar_xml:
                                            n_xml += 1
                                        else:
                                            xml_temp_paths.append(xml_path_report)
                                        lote_xml_ok += 1
                                except Exception as err:
                                    print(f"[WARN] No se pudo obtener XML SOAP para '{nombre_base_pdf}': {err}")
                            else:
                                print(
                                    f"[WARN] La fila '{nombre_base_pdf}' no tiene clave de acceso para solicitar el XML."
                                )

                        if descargar_pdf:
                            pdf_esperados.add(row_id)
                            # En "No Autorizados" el icono XML suele tener id :lnkPdf aunque sea XML.
                            link_id = str(fila_data.get("pdf_link_id") or "").strip()
                            enlace_pdf = fila.locator(
                                "a[id$=':lnkPdf'], a[title*='pdf' i], img[alt*='pdf' i], img[title*='pdf' i]"
                            )
                            if link_id or enlace_pdf.count():
                                contenedor = enlace_pdf.first if enlace_pdf.count() else fila
                                # Si es <img>, buscamos su <a> ancestro
                                if enlace_pdf.count() and contenedor.locator("xpath=ancestor::a[1]").count():
                                    contenedor = contenedor.locator("xpath=ancestor::a[1]").first
                                destino_pdf = pdf_dir / f"{nombre_base_pdf}.pdf"
                                resultado_pdf = None
                                for intento_pdf in range(2):
                                    if link_id and view_state:
                                        resultado_pdf = _descargar_pdf_emitidos_post_con_viewstate(
                                            page, link_id, view_state, destino_pdf
                                        )
                                    if not resultado_pdf:
                                        view_state = _obtener_viewstate_actual() or view_state
                                        if link_id and view_state:
                                            resultado_pdf = _descargar_pdf_emitidos_post_con_viewstate(
                                                page, link_id, view_state, destino_pdf
                                            )
                                    if not resultado_pdf and enlace_pdf.count():
                                        resultado_pdf = _guardar_pdf_desde_jsf(page, contenedor, destino_pdf)
                                    if not resultado_pdf and enlace_pdf.count():
                                        resultado_pdf = _guardar_pdf_desde_enlace(page, contenedor, destino_pdf)
                                    if resultado_pdf:
                                        break
                                    time.sleep(0.15)
                                if resultado_pdf:
                                    pdf_descargados.add(row_id)
                                    n_pdf += 1
                                    lote_pdf_ok += 1
                                    if resultado_pdf.suffix.lower() == ".pdf" and _es_archivo_pdf(resultado_pdf):
                                        datos_pdf = None
                                        if xml_path_report:
                                            try:
                                                if es_retencion:
                                                    datos_xml = _extraer_datos_xml_retencion_emitido(xml_path_report)
                                                elif es_nota_credito:
                                                    datos_xml = _extraer_datos_xml_nota_credito_emitido(xml_path_report)
                                                elif es_nota_debito:
                                                    datos_xml = _extraer_datos_xml_nota_debito_emitido(xml_path_report)
                                                elif es_factura_emitida:
                                                    datos_xml = _extraer_datos_xml_factura_emitido(xml_path_report)
                                                elif es_liquidacion_compra:
                                                    datos_xml = _extraer_datos_xml_liquidacion_compra_emitido(xml_path_report)
                                                else:
                                                    datos_xml = _extraer_datos_xml_pdf_report(xml_path_report)
                                                datos_pdf = datos_xml
                                            except Exception as err:
                                                print(f"[WARN] No se pudo usar XML para el reporte PDF: {err}")
                                        if datos_pdf is None:
                                            if es_retencion:
                                                datos_pdf = _extraer_datos_pdf_retencion_emitido(resultado_pdf)
                                            elif es_nota_credito:
                                                datos_pdf = _extraer_datos_pdf_nota_credito_emitido(resultado_pdf)
                                            elif es_nota_debito:
                                                datos_pdf = _extraer_datos_pdf_nota_debito_emitido(resultado_pdf)
                                            elif es_factura_emitida:
                                                datos_pdf = _extraer_datos_pdf_factura_emitido(resultado_pdf)
                                            elif es_liquidacion_compra:
                                                datos_pdf = _extraer_datos_pdf_liquidacion_compra_emitido(resultado_pdf)
                                            else:
                                                datos_pdf = _extraer_datos_pdf_por_tipo_layout_first(
                                                    resultado_pdf,
                                                    es_retencion=es_retencion,
                                                    es_nota_credito=es_nota_credito,
                                                    es_nota_debito=es_nota_debito,
                                                )
                                        if datos_pdf:
                                            if es_retencion or es_nota_credito or es_nota_debito or es_factura_emitida:
                                                if clave_texto and not datos_pdf.get("Clave de Acceso"):
                                                    datos_pdf["Clave de Acceso"] = clave_texto
                                                if not datos_pdf.get("Número de Autorización") and clave_texto:
                                                    datos_pdf["Número de Autorización"] = clave_texto
                                                if fecha_emision and not datos_pdf.get("Fecha de Emisión"):
                                                    datos_pdf["Fecha de Emisión"] = fecha_emision
                                            else:
                                                if clave_texto and not datos_pdf.get("claveAcceso"):
                                                    datos_pdf["claveAcceso"] = clave_texto
                                                if not datos_pdf.get("numeroComprobante"):
                                                    match = re.search(r"\d{3}-\d{3}-\d{9}", tipo_serie_texto)
                                                    if match:
                                                        datos_pdf["numeroComprobante"] = match.group(0)
                                                if fecha_emision and not datos_pdf.get("fechaEmision"):
                                                    datos_pdf["fechaEmision"] = fecha_emision
                                            pdf_report_rows.append(datos_pdf)
                            else:
                                print(
                                    f"[WARN] No se encontro enlace PDF para '{nombre_base_pdf}' en No Autorizados."
                                )
                    except Exception as err:
                        print(f"[WARN] No se pudo descargar XML/PDF para '{nombre_base_pdf}': {err}")
                    continue

                if descargar_xml_para_reporte and not omitir_soap_xml:
                    if clave_texto:
                        xml_esperados.add(row_id)
                        try:
                            resultado_xml = None
                            for intento_xml in range(2):
                                resultado_xml = _descargar_xml_emitido_por_clave(
                                    request_context,
                                    clave_texto,
                                    xml_dir,
                                    nombre_base_pdf,
                                    claves_guardadas,
                                )
                                if resultado_xml:
                                    break
                                time.sleep(0.15)
                            if resultado_xml:
                                xml_path_report = resultado_xml
                                xml_descargados.add(row_id)
                                if descargar_xml:
                                    n_xml += 1
                                else:
                                    xml_temp_paths.append(xml_path_report)
                                lote_xml_ok += 1
                        except Exception as err:
                            print(f"[WARN] No se pudo obtener XML SOAP para '{nombre_base_pdf}': {err}")
                    else:
                        print(f"[WARN] La fila '{nombre_base_pdf}' no tiene clave de acceso para solicitar el XML.")

                if descargar_pdf:
                    pdf_esperados.add(row_id)
                    link_id = str(fila_data.get("pdf_link_id") or "").strip()
                    link_pdf = fila.locator("a[id$=':lnkPdf']")
                    if not link_id and not link_pdf.count():
                        link_pdf = fila.locator("a[title*='pdf' i], button[title*='pdf' i]")
                    if not link_id and not link_pdf.count():
                        continue

                    destino_pdf = pdf_dir / f"{nombre_base_pdf}.pdf"
                    if not link_id and link_pdf.count():
                        try:
                            link_id = link_pdf.first.get_attribute("id")
                        except Exception:
                            link_id = None
                    resultado_pdf = None
                    for intento_pdf in range(2):
                        if link_id and view_state:
                            resultado_pdf = _descargar_pdf_emitidos_post_con_viewstate(
                                page, link_id, view_state, destino_pdf
                            )
                        if not resultado_pdf:
                            view_state = _obtener_viewstate_actual() or view_state
                            if link_id and view_state:
                                resultado_pdf = _descargar_pdf_emitidos_post_con_viewstate(
                                    page, link_id, view_state, destino_pdf
                                )
                        if not resultado_pdf and link_pdf.count():
                            resultado_pdf = _guardar_pdf_desde_jsf(page, link_pdf.first, destino_pdf)
                        if not resultado_pdf and link_pdf.count():
                            resultado_pdf = _guardar_pdf_desde_enlace(page, link_pdf.first, destino_pdf)
                        if resultado_pdf:
                            break
                        time.sleep(0.15)
                    if resultado_pdf:
                        pdf_descargados.add(row_id)
                        n_pdf += 1
                        lote_pdf_ok += 1
                        if resultado_pdf.suffix.lower() == ".pdf" and _es_archivo_pdf(resultado_pdf):
                            datos_pdf = None
                            if xml_path_report:
                                try:
                                    if es_retencion:
                                        datos_xml = _extraer_datos_xml_retencion_emitido(xml_path_report)
                                    elif es_nota_credito:
                                        datos_xml = _extraer_datos_xml_nota_credito_emitido(xml_path_report)
                                    elif es_nota_debito:
                                        datos_xml = _extraer_datos_xml_nota_debito_emitido(xml_path_report)
                                    elif es_factura_emitida:
                                        datos_xml = _extraer_datos_xml_factura_emitido(xml_path_report)
                                    elif es_liquidacion_compra:
                                        datos_xml = _extraer_datos_xml_liquidacion_compra_emitido(xml_path_report)
                                    else:
                                        datos_xml = _extraer_datos_xml_pdf_report(xml_path_report)
                                    datos_pdf = datos_xml
                                except Exception as err:
                                    print(f"[WARN] No se pudo usar XML para el reporte PDF: {err}")
                            if datos_pdf is None:
                                if es_retencion:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_retencion_emitido(resultado_pdf)
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF de retención para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_nota_credito:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_nota_credito_emitido(resultado_pdf)
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF de nota de crédito para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_nota_debito:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_nota_debito_emitido(resultado_pdf)
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF de nota de débito para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_factura_emitida:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_factura_emitido(resultado_pdf)
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF de factura para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_liquidacion_compra:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_liquidacion_compra_emitido(resultado_pdf)
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF de liquidación de compra para completar el reporte: {err}")
                                        datos_pdf = None
                                else:
                                    datos_dom = _extraer_datos_emitidos_dom(
                                        tipo_visible,
                                        tipo_serie_texto,
                                        clave_texto,
                                        fecha_emision,
                                        fecha_aut_texto,
                                        razon_texto,
                                        valor_sin_imp_texto,
                                        iva_texto,
                                        importe_total_texto,
                                        ruc_emisor=ruc_emisor,
                                    )
                                    detalle_data = None
                                    source_id_detalle = str(fila_data.get("source_id_detalle") or "").strip() or _obtener_source_detalle_emitido(page, idx)
                                    if source_id_detalle:
                                        detalle_data = _obtener_detalle_emitido_xhr(
                                            page,
                                            request_context,
                                            source_id_detalle,
                                            payload_base,
                                            view_state or _obtener_view_state(page),
                                            tipo_visible,
                                            tipo_serie_texto,
                                            clave_texto,
                                            ruc_emisor=ruc_emisor,
                                        )
                                    datos_pdf_archivo = None
                                    try:
                                        datos_pdf_archivo = _extraer_datos_pdf_por_tipo_layout_first(
                                            resultado_pdf,
                                            es_retencion=es_retencion,
                                            es_nota_credito=es_nota_credito,
                                            es_nota_debito=es_nota_debito,
                                        )
                                    except Exception as err:
                                        print(f"[WARN] No se pudo leer el PDF para completar el reporte: {err}")
                                    datos_pdf = _combinar_datos_reporte_emitidos(
                                        datos_dom,
                                        detalle_data,
                                        datos_pdf_archivo,
                                    )
                            if datos_pdf:
                                if es_retencion or es_nota_credito or es_nota_debito or es_factura_emitida:
                                    if clave_texto and not datos_pdf.get("Clave de Acceso"):
                                        datos_pdf["Clave de Acceso"] = clave_texto
                                    if not datos_pdf.get("Número de Autorización") and clave_texto:
                                        datos_pdf["Número de Autorización"] = clave_texto
                                    if fecha_emision and not datos_pdf.get("Fecha de Emisión"):
                                        datos_pdf["Fecha de Emisión"] = fecha_emision
                                else:
                                    if clave_texto and not datos_pdf.get("claveAcceso"):
                                        datos_pdf["claveAcceso"] = clave_texto
                                    if not datos_pdf.get("numeroComprobante"):
                                        match = re.search(r"\d{3}-\d{3}-\d{9}", tipo_serie_texto)
                                        if match:
                                            datos_pdf["numeroComprobante"] = match.group(0)
                                    if fecha_emision and not datos_pdf.get("fechaEmision"):
                                        datos_pdf["fechaEmision"] = fecha_emision
                                pdf_report_rows.append(datos_pdf)
                    else:
                        print(f"[WARN] No se pudo descargar PDF para '{nombre_base_pdf}': no se obtuvo archivo.")

                lote_contador += 1
                if lote_contador >= lote_size or idx == total_filas - 1:
                    duracion_lote = time.perf_counter() - lote_inicio
                    print(
                        f"[INFO] Pag {pagina} lote {((idx // lote_size) + 1)}: "
                        f"{lote_contador} filas, XML {lote_xml_ok}, PDF {lote_pdf_ok}, "
                        f"{duracion_lote:.2f}s"
                    )
                    lote_inicio = time.perf_counter()
                    lote_contador = 0
                    lote_xml_ok = 0
                    lote_pdf_ok = 0

            duracion_pagina = time.perf_counter() - page_inicio
            print(f"[INFO] Pag {pagina} completa: {total_filas} filas en {duracion_pagina:.2f}s")

            boton_siguiente = page.locator("span.ui-paginator-next:not(.ui-state-disabled)")
            if boton_siguiente.count():
                boton_siguiente.first.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=1000)
                except Exception:
                    pass
                time.sleep(0.2)
                pagina += 1
                continue
            break
        if descargar_xml_para_reporte and not omitir_soap_xml:
            faltantes_xml = sorted(xml_esperados - xml_descargados)
            if faltantes_xml:
                print(f"[WARN] Emitidos: faltaron {len(faltantes_xml)} XML respecto a las filas detectadas. Ejemplos: {faltantes_xml[:5]}")
        if descargar_pdf:
            faltantes_pdf = sorted(pdf_esperados - pdf_descargados)
            if faltantes_pdf:
                print(f"[WARN] Emitidos: faltaron {len(faltantes_pdf)} PDF respecto a las filas detectadas. Ejemplos: {faltantes_pdf[:5]}")
        info_base["n_xml"] = n_xml
        info_base["n_pdf"] = n_pdf

        fecha_slug = re.sub(r"[^0-9]+", "", fecha_emision) or "consulta"
        if descargar_xml and n_xml > 0:
            xml_dir.mkdir(parents=True, exist_ok=True)
            xml_report_path = xml_dir / f"emitidos_reporte_xml_{tipo_slug}_{fecha_slug}.xlsx"
            if xml_report_path.exists():
                try:
                    xml_report_path.unlink()
                except PermissionError:
                    sufijo_xml = 1
                    while True:
                        candidato = xml_dir / f"emitidos_reporte_xml_{tipo_slug}_{fecha_slug}_{sufijo_xml}.xlsx"
                        if not candidato.exists():
                            xml_report_path = candidato
                            break
                        sufijo_xml += 1
            try:
                estado_default_reporte = estado_visible if modo_no_autorizados else None
                xml_files_emitidos = _xml_files_por_tipo(carpeta_estado, tipo_prefijo)
                construir_reporte(carpeta_estado, xml_report_path, estado_default_reporte, xml_files=xml_files_emitidos)
                info_base["reporte_xml"] = str(xml_report_path)
            except Exception as err:
                print(f"[WARN] No se pudo construir el reporte XML de emitidos: {err}")

    df = pd.DataFrame(data)
    fecha_slug = re.sub(r"[^0-9]+", "", fecha_emision) or "consulta"
    if descargar_pdf and n_pdf > 0 and pdf_report_rows:
        pdf_dir.mkdir(parents=True, exist_ok=True)
        pdf_report_path = pdf_dir / f"emitidos_reporte_pdf_{tipo_slug}_{fecha_slug}.xlsx"
        if pdf_report_path.exists():
            try:
                pdf_report_path.unlink()
            except PermissionError:
                sufijo_pdf = 1
                while True:
                    candidato = pdf_dir / f"emitidos_reporte_pdf_{tipo_slug}_{fecha_slug}_{sufijo_pdf}.xlsx"
                    if not candidato.exists():
                        pdf_report_path = candidato
                        break
                    sufijo_pdf += 1
        if es_retencion:
            if _guardar_reporte_pdf_retencion_emitidos_excel(pdf_report_rows, pdf_report_path):
                info_base["reporte_pdf"] = str(pdf_report_path)
        elif es_nota_credito:
            if _guardar_reporte_pdf_nota_credito_emitidos_excel(pdf_report_rows, pdf_report_path):
                info_base["reporte_pdf"] = str(pdf_report_path)
        elif es_nota_debito:
            if _guardar_reporte_pdf_nota_debito_emitidos_excel(pdf_report_rows, pdf_report_path):
                info_base["reporte_pdf"] = str(pdf_report_path)
        elif es_factura_emitida:
            if _guardar_reporte_pdf_factura_emitidos_excel(pdf_report_rows, pdf_report_path):
                info_base["reporte_pdf"] = str(pdf_report_path)
        elif _guardar_reporte_pdf_excel(pdf_report_rows, pdf_report_path):
            info_base["reporte_pdf"] = str(pdf_report_path)
    if not df.empty:
        info_base.update({
            "estado": "ok",
            "n_registros": len(df),
        })
    else:
        info_base.update({
            "estado": "sin_resultados",
            "n_registros": 0,
        })
    if not descargar_xml and xml_temp_paths:
        for xml_tmp in xml_temp_paths:
            try:
                Path(xml_tmp).unlink(missing_ok=True)
            except Exception:
                pass
    return info_base

# ============================================================
# FUNCIoN PRINCIPAL
# ============================================================
def descargar_sri(
    ruc: str,
    clave: str,
    anio: int,
    mes: int,
    dia: int,
    tipo: str,
    formatos: list,
    destino: Path,
    mes_fin: Optional[int] = None,
    origen: str = "Recibidos",
    ci_adicional: Optional[str] = None,
    fecha_emitidos: Optional[str] = None,
    estado_emitidos: Optional[str] = None,
    establecimiento: Optional[str] = None,
    punto_emision: Optional[str] = None,
    checkpoint_path: Optional[str] = None,
    resume_download: bool = False,
):
    _check_cancel("inicio_descarga")
    formatos_norm = [(fmt or "").strip().upper() for fmt in (formatos or []) if isinstance(fmt, str)]
    hoy = datetime.now().date()
    aviso_recorte = None
    destino.mkdir(parents=True, exist_ok=True)
    destino_recibidos = destino / "Recibidos"
    destino_emitidos = destino / "Emitidos"
    destino_recibidos.mkdir(parents=True, exist_ok=True)
    destino_emitidos.mkdir(parents=True, exist_ok=True)
    destino_objetivo = destino
    cookies_path = Path(f"cookies_{ruc}.json")
    checkpoint_path_str = str(checkpoint_path or "").strip()
    checkpoint_data = _load_download_checkpoint(checkpoint_path_str) if checkpoint_path_str else None
    progress_data = checkpoint_data.get("progress") if isinstance(checkpoint_data, dict) and isinstance(checkpoint_data.get("progress"), dict) else {}
    try:
        resume_month = int(progress_data.get("next_month") or mes)
    except Exception:
        resume_month = int(mes)
    try:
        resume_day = int(progress_data.get("next_day") or (dia if dia not in (None, "") else 0))
    except Exception:
        resume_day = int(dia or 0)
    if checkpoint_path_str and checkpoint_data:
        _mark_download_checkpoint_running(checkpoint_path_str)

    with sync_playwright() as p:
        launch_kwargs = dict(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
        )
        # Prefer Chrome channel when available (local), fallback to bundled Chromium (Render)
        launch_kwargs["channel"] = "chrome"
        if SLOW_MO > 0:
            launch_kwargs["slow_mo"] = SLOW_MO
        if DEVTOOLS:
            launch_kwargs["devtools"] = True
            launch_kwargs["headless"] = False

        browser = None
        using_persistent_profile = False
        context = None
        if USE_PERSISTENT_PROFILE:
            persistent_profile_dir = Path(USER_DATA_DIR).expanduser()
            if not persistent_profile_dir.is_absolute():
                persistent_profile_dir = Path.cwd() / persistent_profile_dir
            try:
                persistent_profile_dir.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            persistent_kwargs = dict(launch_kwargs)
            persistent_kwargs["accept_downloads"] = True
            try:
                context = p.chromium.launch_persistent_context(
                    str(persistent_profile_dir),
                    **persistent_kwargs,
                )
                using_persistent_profile = True
            except Exception as err:
                print(f"[WARN] No se pudo usar perfil persistente; fallback a contexto normal: {err}")
                persistent_kwargs.pop("channel", None)
                try:
                    context = p.chromium.launch_persistent_context(
                        str(persistent_profile_dir),
                        **persistent_kwargs,
                    )
                    using_persistent_profile = True
                except Exception:
                    context = None

        if context is None:
            try:
                browser = p.chromium.launch(**launch_kwargs)
            except Exception:
                launch_kwargs.pop("channel", None)
                browser = p.chromium.launch(**launch_kwargs)
            context = browser.new_context(accept_downloads=True)

        page = context.pages[0] if context.pages else context.new_page()

        destino_url = PORTAL_HOME if origen in {"Recibidos", "Emitidos"} else URLS.get(origen, URLS["Recibidos"])
        _login(context, page, ruc, clave, cookies_path, destino_url, ci_adicional=ci_adicional)
        _check_cancel("post_login")
        if "auth/realms" in page.url:
            raise RuntimeError("La autenticacion en el SRI fallo, se mantuvo en la pantalla de login.")
        modulo_page = None
        if origen == "Recibidos":
            modulo_page = _abrir_modulo_consultas(page, origen)
            destino_objetivo = destino_recibidos
            consultas_recibidos = 0

            def _reabrir_modulo_recibidos_si_toca() -> None:
                nonlocal modulo_page, consultas_recibidos
                if consultas_recibidos <= 0:
                    return
                if consultas_recibidos % 15 != 0:
                    return
                print(
                    f"[INFO] Recibidos: reabriendo modulo tras {consultas_recibidos} consultas para reducir bloqueos de captcha."
                )
                modulo_page = _abrir_modulo_consultas(page, origen)

            def _consultar_recibidos_dia(mes_actual: int, dia_actual: int):
                nonlocal consultas_recibidos
                _reabrir_modulo_recibidos_si_toca()
                resultado_dia = _flujo_recibidos(
                    modulo_page,
                    destino_objetivo,
                    anio,
                    mes_actual,
                    dia_actual,
                    tipo,
                    formatos,
                )
                consultas_recibidos += 1
                return resultado_dia

            formatos_norm = [(fmt or "").strip().upper() for fmt in (formatos or []) if isinstance(fmt, str)]

            def _recibidos_por_mes(mes_actual: int, dia_actual: int):
                nonlocal aviso_recorte
                if anio > hoy.year:
                    return {"estado": "sin_resultados", "n_xml": 0, "n_pdf": 0}
                if anio == hoy.year and mes_actual > hoy.month:
                    return {"estado": "sin_resultados", "n_xml": 0, "n_pdf": 0}

                dias_en_mes = calendar.monthrange(anio, mes_actual)[1]
                limite_dia = dias_en_mes
                if anio == hoy.year and mes_actual == hoy.month:
                    limite_dia = min(limite_dia, hoy.day)
                    if limite_dia < dias_en_mes:
                        aviso_recorte = (
                            f"Rango ajustado hasta el día actual ({hoy.day:02d}/{hoy.month:02d}/{hoy.year})."
                        )

                if dia_actual in (0, None):
                    resultado_mes = _consultar_recibidos_dia(mes_actual, 0)
                    if resultado_mes is None:
                        return {"estado": "sin_resultados", "n_xml": 0, "n_pdf": 0}
                    resultado_mes = dict(resultado_mes)
                    resultado_mes["fecha_filtro"] = f"01/{mes_actual:02d}/{anio} - {limite_dia:02d}/{mes_actual:02d}/{anio}"
                    resultado_mes["mensaje"] = "Consulta mensual realizada en modo Todos"
                    resultado_mes["detalles_dias"] = []
                    return resultado_mes
                else:
                    dia_int = int(dia_actual)
                    if dia_int > limite_dia:
                        return {"estado": "sin_resultados", "n_xml": 0, "n_pdf": 0}
                    dias_consultar = [dia_int]

                if not dias_consultar:
                    return {"estado": "sin_resultados", "n_xml": 0, "n_pdf": 0}

                total_xml = 0
                total_pdf = 0
                detalle_dias = []
                reportes_pdf_dia = []
                resultado_mes = None

                for idx_dia, dia_iter in enumerate(dias_consultar):
                    _check_cancel("recibidos_dia")
                    resultado_dia = _consultar_recibidos_dia(mes_actual, dia_iter)
                    total_xml += resultado_dia.get("n_xml", 0)
                    total_pdf += resultado_dia.get("n_pdf", 0)
                    detalle_dias.append(
                        {
                            "dia": dia_iter,
                            "estado": resultado_dia.get("estado"),
                            "n_xml": resultado_dia.get("n_xml", 0),
                            "n_pdf": resultado_dia.get("n_pdf", 0),
                        }
                    )
                    reporte_pdf_dia = resultado_dia.get("reporte_pdf")
                    if reporte_pdf_dia and Path(reporte_pdf_dia).exists():
                        reportes_pdf_dia.append(str(reporte_pdf_dia))
                    if checkpoint_path_str and idx_dia < len(dias_consultar) - 1:
                        siguiente_dia = dias_consultar[idx_dia + 1]
                        _update_download_checkpoint_progress(
                            checkpoint_path_str,
                            next_month=int(mes_actual),
                            next_day=int(siguiente_dia),
                            last_completed_day=int(dia_iter),
                            last_completed_label=f"{dia_iter:02d}/{mes_actual:02d}/{anio}",
                        )
                    resultado_mes = resultado_dia

                if dia_actual not in (0, None) and len(dias_consultar) == 1:
                    return resultado_mes

                resultado_mes = dict(resultado_mes or {})
                resultado_mes["n_xml"] = total_xml
                resultado_mes["n_pdf"] = total_pdf
                resultado_mes["estado"] = "sin_resultados" if total_xml == 0 and total_pdf == 0 else "ok"
                resultado_mes["mensaje"] = f"Procesados {len(dias_consultar)} días del mes"
                resultado_mes["detalles_dias"] = detalle_dias
                resultado_mes["fecha_filtro"] = (
                    f"{dias_consultar[0]:02d}/{mes_actual:02d}/{anio}"
                    f" - {dias_consultar[-1]:02d}/{mes_actual:02d}/{anio}"
                )

                tipo_visible = TIPOS_MAP.get(tipo, tipo)
                tipo_slug = resultado_mes.get("tipo_slug", _slug_tipo(tipo_visible or tipo))
                _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
                tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
                base_mes = destino_objetivo / tipo_dir_nombre / f"{anio:04d}" / _mes_a_texto(mes_actual)

                if "XML" in formatos_norm and total_xml > 0 and base_mes.exists():
                    xml_files = _xml_files_por_tipo(base_mes, tipo_prefijo)
                    if xml_files:
                        xml_dir_mes = base_mes / "XML"
                        xml_dir_mes.mkdir(parents=True, exist_ok=True)
                        destino_xml_mes = xml_dir_mes / f"recibidos_reporte_xml_{tipo_slug}_{anio:04d}{mes_actual:02d}.xlsx"
                        try:
                            construir_reporte(base_mes, destino_xml_mes, None, xml_files=xml_files)
                        except Exception as err:
                            print(f"[WARN] No se pudo construir el reporte XML mensual de recibidos: {err}")
                        if destino_xml_mes.exists():
                            resultado_mes["reporte_xml"] = str(destino_xml_mes)
                    resultado_mes["xml_dir"] = str(base_mes / "XML")

                if "PDF" in formatos_norm:
                    sufijos_dia = [f"{anio:04d}{mes_actual:02d}{int(d):02d}" for d in dias_consultar]
                    reportes_pdf_dia = _collect_existing_reports(
                        base_mes / "PDF",
                        "recibidos_reporte_pdf",
                        tipo_slug,
                        sufijos_dia,
                    )
                    if reportes_pdf_dia:
                        pdf_dir_mes = base_mes / "PDF"
                        pdf_dir_mes.mkdir(parents=True, exist_ok=True)
                        destino_pdf_mes = pdf_dir_mes / f"recibidos_reporte_pdf_{tipo_slug}_{anio:04d}{mes_actual:02d}.xlsx"
                        try:
                            pdf_mes = _consolidar_reportes_excel(reportes_pdf_dia, destino_pdf_mes)
                        except Exception as err:
                            print(f"[WARN] No se pudo consolidar reporte PDF mensual de recibidos: {err}")
                            pdf_mes = None
                        if pdf_mes and Path(pdf_mes).exists():
                            resultado_mes["reporte_pdf"] = str(pdf_mes)

                resultado_mes["carpeta_tipo"] = str(base_mes if base_mes.exists() else destino_objetivo)
                return resultado_mes

            mes_fin_val = None
            try:
                if mes_fin not in (None, "", 0):
                    mes_fin_val = int(mes_fin)
            except Exception:
                mes_fin_val = None
            if anio > hoy.year:
                return {
                    "estado": "sin_resultados",
                    "mensaje": "La fecha solicitada es futura.",
                    "n_xml": 0,
                    "n_pdf": 0,
                }
            if anio == hoy.year:
                max_mes = hoy.month
                if mes_fin_val and mes_fin_val > max_mes:
                    aviso_recorte = f"Rango ajustado hasta el mes actual ({hoy.month:02d}/{hoy.year})."
                    mes_fin_val = max_mes
                if mes > max_mes:
                    return {
                        "estado": "sin_resultados",
                        "mensaje": "La fecha solicitada es futura.",
                        "n_xml": 0,
                        "n_pdf": 0,
                        "aviso_recorte": aviso_recorte,
                    }
                if dia not in (0, None) and mes == max_mes and int(dia) > hoy.day:
                    return {
                        "estado": "sin_resultados",
                        "mensaje": "La fecha solicitada es futura.",
                        "n_xml": 0,
                        "n_pdf": 0,
                        "aviso_recorte": aviso_recorte,
                    }
            if mes_fin_val and mes_fin_val >= mes and dia in (0, None):
                reportes_xml = []
                reportes_pdf = []
                detalle_meses = []
                total_xml = 0
                total_pdf = 0
                mes_inicio = int(resume_month if resume_download else mes)
                mes_fin_val = int(mes_fin_val)
                resultado_mes = None
                for mes_actual in range(mes_inicio, mes_fin_val + 1):
                    _check_cancel("recibidos_mes")
                    resultado_mes = _recibidos_por_mes(mes_actual, 0)
                    total_xml += resultado_mes.get("n_xml", 0)
                    total_pdf += resultado_mes.get("n_pdf", 0)
                    detalle_meses.append(
                        {
                            "mes": mes_actual,
                            "estado": resultado_mes.get("estado"),
                            "n_xml": resultado_mes.get("n_xml", 0),
                            "n_pdf": resultado_mes.get("n_pdf", 0),
                        }
                    )
                    if resultado_mes.get("reporte_xml") and Path(resultado_mes.get("reporte_xml")).exists():
                        reportes_xml.append(str(resultado_mes.get("reporte_xml")))
                    reporte_pdf_mes = resultado_mes.get("reporte_pdf")
                    if reporte_pdf_mes and Path(reporte_pdf_mes).exists():
                        reportes_pdf.append(str(reporte_pdf_mes))
                    if checkpoint_path_str:
                        siguiente_mes = mes_actual + 1 if mes_actual < mes_fin_val else None
                        _update_download_checkpoint_progress(
                            checkpoint_path_str,
                            next_month=siguiente_mes,
                            next_day=0,
                            last_completed_day=None,
                            last_completed_label=f"{_mes_a_texto(mes_actual)} {anio}",
                        )
                resultado = dict(resultado_mes or {})
                resultado["n_xml"] = total_xml
                resultado["n_pdf"] = total_pdf
                resultado["estado"] = "sin_resultados" if total_xml == 0 and total_pdf == 0 else "ok"
                resultado["mensaje"] = f"Procesados {mes_fin_val - mes_inicio + 1} meses"
                if aviso_recorte:
                    resultado["aviso_recorte"] = aviso_recorte
                resultado["detalles_meses"] = detalle_meses
                resultado["rango_meses"] = True
                resultado["mes_inicio"] = mes_inicio
                resultado["mes_fin"] = mes_fin_val
                fecha_inicio = f"01/{mes_inicio:02d}/{anio}"
                dia_fin = calendar.monthrange(anio, mes_fin_val)[1]
                if anio == hoy.year and mes_fin_val == hoy.month:
                    dia_fin = min(dia_fin, hoy.day)
                fecha_fin = f"{dia_fin:02d}/{mes_fin_val:02d}/{anio}"
                resultado["fecha_filtro"] = f"{fecha_inicio} - {fecha_fin}"
                resultado["reportes_xml"] = reportes_xml
                tipo_visible = TIPOS_MAP.get(tipo, tipo)
                tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
                tipo_slug = resultado.get(
                    "tipo_slug",
                    _slug_tipo(tipo_visible or tipo),
                )
                if "PDF" in formatos_norm:
                    sufijos_mes = [f"{anio:04d}{m:02d}" for m in range(mes_inicio, mes_fin_val + 1)]
                    carpeta_rango = destino_objetivo / tipo_dir_nombre / f"{anio:04d}" / "PDF"
                    reportes_pdf = _collect_existing_reports(
                        carpeta_rango,
                        "recibidos_reporte_pdf",
                        tipo_slug,
                        sufijos_mes,
                    )
                resultado["reportes_pdf"] = reportes_pdf
                if mes_inicio == 1 and mes_fin_val == 12:
                    if "XML" in formatos_norm:
                        base_anual = destino_objetivo / tipo_dir_nombre / f"{anio:04d}"
                        if base_anual.exists():
                            _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
                            xml_files = _xml_files_por_tipo(base_anual, tipo_prefijo)
                            if xml_files:
                                xml_dir_anual = base_anual / "XML"
                                xml_dir_anual.mkdir(parents=True, exist_ok=True)
                                destino_anual_xml = xml_dir_anual / f"recibidos_reporte_xml_{tipo_slug}_{anio:04d}.xlsx"
                                try:
                                    construir_reporte(base_anual, destino_anual_xml, None, xml_files=xml_files)
                                except Exception as err:
                                    print(f"[WARN] No se pudo construir reporte XML anual (recibidos): {err}")
                                if destino_anual_xml.exists():
                                    resultado["reporte_xml_anual"] = str(destino_anual_xml)
                            elif reportes_xml:
                                xml_dir_anual = base_anual / "XML"
                                xml_dir_anual.mkdir(parents=True, exist_ok=True)
                                destino_anual_xml = xml_dir_anual / f"recibidos_reporte_xml_{tipo_slug}_{anio:04d}.xlsx"
                                anual_xml = _consolidar_reportes_excel(reportes_xml, destino_anual_xml)
                                if anual_xml:
                                    resultado["reporte_xml_anual"] = str(anual_xml)
                    if reportes_pdf:
                        pdf_dir_anual = destino_objetivo / tipo_dir_nombre / f"{anio:04d}" / "PDF"
                        pdf_dir_anual.mkdir(parents=True, exist_ok=True)
                        destino_anual_pdf = pdf_dir_anual / f"recibidos_reporte_pdf_{tipo_slug}_{anio:04d}.xlsx"
                        anual_pdf = _consolidar_reportes_excel(reportes_pdf, destino_anual_pdf)
                        if anual_pdf:
                            resultado["reporte_pdf_anual"] = str(anual_pdf)
                    resultado["anual"] = True
                else:
                    if "XML" in formatos_norm:
                        base_rango = destino_objetivo / tipo_dir_nombre / f"{anio:04d}"
                        if base_rango.exists():
                            _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
                            meses_rango = range(mes_inicio, mes_fin_val + 1)
                            xml_files = _xml_files_por_meses(base_rango, tipo_prefijo, meses_rango)
                            if xml_files:
                                sufijo_rango = f"{anio:04d}{mes_inicio:02d}{mes_fin_val:02d}"
                                xml_dir_rango = base_rango / "XML"
                                xml_dir_rango.mkdir(parents=True, exist_ok=True)
                                destino_rango_xml = xml_dir_rango / f"recibidos_reporte_xml_{tipo_slug}_{sufijo_rango}.xlsx"
                                try:
                                    construir_reporte(base_rango, destino_rango_xml, None, xml_files=xml_files)
                                except Exception as err:
                                    print(f"[WARN] No se pudo construir reporte XML del rango (recibidos): {err}")
                                if destino_rango_xml.exists():
                                    resultado["reporte_xml_rango"] = str(destino_rango_xml)
                resultado.pop("reporte_pdf", None)
                resultado.pop("reporte_xml", None)
                carpeta_rango = destino_objetivo / tipo_dir_nombre / f"{anio:04d}"
                resultado["carpeta_tipo"] = str(carpeta_rango if carpeta_rango.exists() else destino_objetivo)
            else:
                mes_objetivo = int(resume_month if resume_download else mes)
                dia_objetivo = int(resume_day) if resume_download else int(dia)
                if dia in (0, None):
                    resultado = _recibidos_por_mes(mes_objetivo, 0)
                else:
                    resultado = _recibidos_por_mes(mes_objetivo, int(dia_objetivo))
                if aviso_recorte:
                    resultado["aviso_recorte"] = aviso_recorte
        elif origen == "Emitidos":
            modulo_page = _abrir_modulo_consultas(page, origen)
            _resolver_captcha(modulo_page, f"{origen.lower()}_Modulo")
            destino_objetivo = destino_emitidos

            def _reiniciar_emitidos_para_siguiente_dia(fecha_actual: str, total_docs_dia: int) -> None:
                nonlocal modulo_page
                mensaje = (
                    f"[INFO] Emitidos: el dia {fecha_actual} termino con {total_docs_dia} documentos. "
                    "Se reabrira el modulo antes de continuar con el siguiente dia."
                )
                _notificar_usuario_accion(mensaje)
                try:
                    page.goto(PORTAL_HOME, wait_until="domcontentloaded", timeout=15000)
                    page.wait_for_load_state("domcontentloaded", timeout=3000)
                except Exception as err:
                    print(f"[WARN] No se pudo volver al menu principal antes del reinicio de Emitidos: {err}")
                if EMITIDOS_RESET_PAUSE_MS > 0:
                    try:
                        page.wait_for_timeout(EMITIDOS_RESET_PAUSE_MS)
                    except Exception:
                        pass
                modulo_page = _abrir_modulo_consultas(page, origen)
                _resolver_captcha(modulo_page, f"{origen.lower()}_Modulo_Reinicio")

            def _emitidos_por_mes(mes_actual: int, dia_actual: int):
                nonlocal aviso_recorte
                if anio > hoy.year:
                    return {"estado": "sin_descargas", "n_registros": 0, "n_xml": 0, "n_pdf": 0}
                if anio == hoy.year and mes_actual > hoy.month:
                    return {"estado": "sin_descargas", "n_registros": 0, "n_xml": 0, "n_pdf": 0}
                dias_consultar = []
                dias_en_mes = calendar.monthrange(anio, mes_actual)[1]
                limite_dia = dias_en_mes
                if anio == hoy.year and mes_actual == hoy.month:
                    limite_dia = min(limite_dia, hoy.day)
                    if limite_dia < dias_en_mes:
                        aviso_recorte = f"Rango ajustado hasta el día actual ({hoy.day:02d}/{hoy.month:02d}/{hoy.year})."
                if dia_actual in (0, None):
                    dia_inicio = 1
                    if resume_download and mes_actual == resume_month and resume_day not in (0, None):
                        dia_inicio = max(1, min(int(resume_day), limite_dia))
                    dias_consultar = list(range(dia_inicio, limite_dia + 1))
                else:
                    dia_int = int(dia_actual)
                    if dia_int > limite_dia:
                        return {"estado": "sin_descargas", "n_registros": 0, "n_xml": 0, "n_pdf": 0}
                    dias_consultar = [dia_int]
                total_regs = total_xml = total_pdf = 0
                detalle_dias = []
                formatos_norm = [(fmt or "").strip().upper() for fmt in (formatos or []) if isinstance(fmt, str)]
                descargar_pdf_mes = "PDF" in formatos_norm
                reportes_dia = []
                reportes_pdf_generados = []
                reportes_xml_generados = []
                resultado_mes = None
                for idx_dia, dia_iter in enumerate(dias_consultar):
                    _check_cancel("emitidos_dia")
                    fecha_actual = f"{dia_iter:02d}/{mes_actual:02d}/{anio}"
                    resultado_dia = _flujo_emitidos(
                        modulo_page,
                        destino_objetivo,
                        fecha_actual,
                        tipo,
                        estado_emitidos,
                        establecimiento,
                        punto_emision,
                        formatos,
                        ruc_emisor=ruc,
                    )
                    detalle_dias.append(
                        {
                            "dia": dia_iter,
                            "estado": resultado_dia.get("estado"),
                            "n_registros": resultado_dia.get("n_registros", 0),
                        }
                    )
                    total_regs += resultado_dia.get("n_registros", 0)
                    total_xml += resultado_dia.get("n_xml", 0)
                    total_pdf += resultado_dia.get("n_pdf", 0)
                    resultado_mes = resultado_dia
                    reporte_xml_dia = resultado_dia.get("reporte_xml")
                    if reporte_xml_dia and Path(reporte_xml_dia).exists():
                        reportes_xml_generados.append(str(Path(reporte_xml_dia)))
                    if descargar_pdf_mes:
                        reporte_dia = resultado_dia.get("reporte_pdf")
                        if reporte_dia and Path(reporte_dia).exists():
                            reporte_dia = str(Path(reporte_dia))
                            reportes_dia.append(reporte_dia)
                            reportes_pdf_generados.append(reporte_dia)
                    if checkpoint_path_str and idx_dia < len(dias_consultar) - 1:
                        _update_download_checkpoint_progress(
                            checkpoint_path_str,
                            next_month=int(mes_actual),
                            next_day=int(dias_consultar[idx_dia + 1]),
                            last_completed_day=int(dia_iter),
                            last_completed_label=f"{dia_iter:02d}/{mes_actual:02d}/{anio}",
                        )
                    n_registros_dia = int(resultado_dia.get("n_registros", 0) or 0)
                    hay_mas_trabajo = (
                        idx_dia < len(dias_consultar) - 1
                        or (mes_fin_val and int(mes_fin_val) > int(mes_actual) and dia_actual in (0, None))
                    )
                    if hay_mas_trabajo and n_registros_dia >= EMITIDOS_RESET_AFTER_DAY_DOCS:
                        _reiniciar_emitidos_para_siguiente_dia(fecha_actual, n_registros_dia)
                if dia_actual in (0, None):
                    resultado_mes = dict(resultado_mes or {})
                    resultado_mes["n_registros"] = total_regs
                    resultado_mes["n_xml"] = total_xml
                    resultado_mes["n_pdf"] = total_pdf
                    resultado_mes["estado"] = "sin_descargas" if total_regs == 0 else "ok"
                    resultado_mes["mensaje"] = f"Procesados {len(dias_consultar)} días del mes"
                    resultado_mes["detalles_dias"] = detalle_dias
                    estado_nombre = (ESTADOS_EMITIDOS_MAP.get(estado_emitidos, estado_emitidos) or "Sin Estado").strip() or "Sin Estado"
                    estado_normalizado = unicodedata.normalize("NFKD", estado_nombre).encode("ascii", "ignore").decode("ascii")
                    estado_slug = re.sub(r"[^A-Za-z0-9]+", "_", estado_normalizado).strip("_") or "Sin_Estado"
                    estado_norm = estado_normalizado.lower()
                    estado_default_reporte = estado_nombre if "no autoriz" in estado_norm else None
                    tipo_visible = TIPOS_MAP.get(tipo, tipo)
                    tipo_slug = _slug_tipo(tipo_visible or tipo)
                    tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
                    anio_dir = f"{anio:04d}"
                    mes_dir = _mes_a_texto(mes_actual)
                    carpeta_mes = destino_emitidos / estado_slug / tipo_dir_nombre / anio_dir / mes_dir
                    sufijos_dia = [f"{anio:04d}{mes_actual:02d}{int(d):02d}" for d in dias_consultar]
                    reportes_xml_dia = _collect_existing_reports(
                        carpeta_mes / "XML",
                        "emitidos_reporte_xml",
                        tipo_slug,
                        sufijos_dia,
                    )
                    reportes_xml_dia = sorted({*reportes_xml_dia, *reportes_xml_generados})
                    reportes_pdf_dia = _collect_existing_reports(
                        carpeta_mes / "PDF",
                        "emitidos_reporte_pdf",
                        tipo_slug,
                        sufijos_dia,
                    )
                    reportes_pdf_dia = sorted({*reportes_pdf_dia, *reportes_pdf_generados})
                    if "XML" in formatos_norm and total_xml > 0:
                        if carpeta_mes.exists():
                            xml_dir_mes = carpeta_mes / "XML"
                            xml_dir_mes.mkdir(parents=True, exist_ok=True)
                            xml_report_path = xml_dir_mes / f"emitidos_reporte_xml_{tipo_slug}_{anio_dir}{mes_actual:02d}.xlsx"
                            if xml_report_path.exists():
                                try:
                                    xml_report_path.unlink()
                                except PermissionError:
                                    sufijo_xml = 1
                                    while True:
                                        candidato = xml_dir_mes / f"emitidos_reporte_xml_{tipo_slug}_{anio_dir}{mes_actual:02d}_{sufijo_xml}.xlsx"
                                        if not candidato.exists():
                                            xml_report_path = candidato
                                            break
                                        sufijo_xml += 1
                            try:
                                xml_files_mes = _xml_files_por_tipo(carpeta_mes, _prefijo_tipo(tipo_visible or tipo)[2])
                                construir_reporte(carpeta_mes, xml_report_path, estado_default_reporte, xml_files=xml_files_mes)
                                resultado_mes["reporte_xml"] = str(xml_report_path)
                                _delete_report_files(reportes_xml_dia)
                            except Exception as err:
                                print(f"[WARN] No se pudo construir el reporte XML mensual de emitidos: {err}")
                    if descargar_pdf_mes:
                        reportes_dia = list(reportes_pdf_dia)
                    if descargar_pdf_mes and reportes_dia:
                        frames = []
                        for ruta_excel in reportes_dia:
                            try:
                                df_dia = pd.read_excel(ruta_excel)
                            except Exception as err:
                                print(f"[WARN] No se pudo leer reporte diario '{ruta_excel}': {err}")
                                continue
                            if not df_dia.empty:
                                frames.append(df_dia)
                        if frames:
                            df_mes = pd.concat(frames, ignore_index=True)
                            carpeta_mes.mkdir(parents=True, exist_ok=True)
                            pdf_dir_mes = carpeta_mes / "PDF"
                            pdf_dir_mes.mkdir(parents=True, exist_ok=True)
                            pdf_report_path = pdf_dir_mes / f"emitidos_reporte_pdf_{tipo_slug}_{anio_dir}{mes_actual:02d}.xlsx"
                            if pdf_report_path.exists():
                                try:
                                    pdf_report_path.unlink()
                                except PermissionError:
                                    sufijo_pdf = 1
                                    while True:
                                        candidato = pdf_dir_mes / f"emitidos_reporte_pdf_{tipo_slug}_{anio_dir}{mes_actual:02d}_{sufijo_pdf}.xlsx"
                                        if not candidato.exists():
                                            pdf_report_path = candidato
                                            break
                                        sufijo_pdf += 1
                            filas_pdf = df_mes.to_dict(orient="records")
                            if _es_tipo_retencion(tipo_visible or tipo):
                                if _guardar_reporte_pdf_retencion_emitidos_excel(filas_pdf, pdf_report_path):
                                    resultado_mes["reporte_pdf"] = str(pdf_report_path)
                            elif _es_tipo_nota_credito(tipo_visible or tipo):
                                if _guardar_reporte_pdf_nota_credito_emitidos_excel(filas_pdf, pdf_report_path):
                                    resultado_mes["reporte_pdf"] = str(pdf_report_path)
                            elif _es_tipo_nota_debito(tipo_visible or tipo):
                                if _guardar_reporte_pdf_nota_debito_emitidos_excel(filas_pdf, pdf_report_path):
                                    resultado_mes["reporte_pdf"] = str(pdf_report_path)
                            elif _es_tipo_factura(tipo_visible or tipo):
                                if _guardar_reporte_pdf_factura_emitidos_excel(filas_pdf, pdf_report_path):
                                    resultado_mes["reporte_pdf"] = str(pdf_report_path)
                            elif _guardar_reporte_pdf_excel(filas_pdf, pdf_report_path):
                                resultado_mes["reporte_pdf"] = str(pdf_report_path)
                            if resultado_mes.get("reporte_pdf"):
                                _delete_report_files(reportes_dia)
                return resultado_mes

            mes_fin_val = None
            try:
                if mes_fin not in (None, "", 0):
                    mes_fin_val = int(mes_fin)
            except Exception:
                mes_fin_val = None
            if anio > hoy.year:
                return {"estado": "sin_descargas", "n_registros": 0, "n_xml": 0, "n_pdf": 0}
            if anio == hoy.year:
                max_mes = hoy.month
                if mes_fin_val and mes_fin_val > max_mes:
                    aviso_recorte = f"Rango ajustado hasta el mes actual ({hoy.month:02d}/{hoy.year})."
                    mes_fin_val = max_mes
                if mes > max_mes:
                    return {"estado": "sin_descargas", "n_registros": 0, "n_xml": 0, "n_pdf": 0}
            if mes_fin_val and mes_fin_val >= mes and dia in (0, None):
                reportes_xml = []
                reportes_pdf = []
                detalle_meses = []
                total_regs = total_xml = total_pdf = 0
                mes_inicio = int(resume_month if resume_download else mes)
                mes_fin_val = int(mes_fin_val)
                resultado_mes = None
                for mes_actual in range(mes_inicio, mes_fin_val + 1):
                    _check_cancel("emitidos_mes")
                    resultado_mes = _emitidos_por_mes(mes_actual, 0)
                    detalle_meses.append(
                        {
                            "mes": mes_actual,
                            "estado": resultado_mes.get("estado"),
                            "n_registros": resultado_mes.get("n_registros", 0),
                        }
                    )
                    total_regs += resultado_mes.get("n_registros", 0)
                    total_xml += resultado_mes.get("n_xml", 0)
                    total_pdf += resultado_mes.get("n_pdf", 0)
                    if resultado_mes.get("reporte_xml"):
                        reportes_xml.append(resultado_mes["reporte_xml"])
                    if resultado_mes.get("reporte_pdf"):
                        reportes_pdf.append(resultado_mes["reporte_pdf"])
                    if checkpoint_path_str:
                        siguiente_mes = mes_actual + 1 if mes_actual < mes_fin_val else None
                        _update_download_checkpoint_progress(
                            checkpoint_path_str,
                            next_month=siguiente_mes,
                            next_day=0,
                            last_completed_day=None,
                            last_completed_label=f"{_mes_a_texto(mes_actual)} {anio}",
                        )
                resultado = dict(resultado_mes or {})
                resultado["n_registros"] = total_regs
                resultado["n_xml"] = total_xml
                resultado["n_pdf"] = total_pdf
                resultado["estado"] = "sin_descargas" if total_regs == 0 else "ok"
                resultado["mensaje"] = f"Procesados {mes_fin_val - mes_inicio + 1} meses"
                if aviso_recorte:
                    resultado["aviso_recorte"] = aviso_recorte
                resultado["detalles_meses"] = detalle_meses
                resultado["rango_meses"] = True
                resultado["mes_inicio"] = mes_inicio
                resultado["mes_fin"] = mes_fin_val
                fecha_inicio = f"01/{mes_inicio:02d}/{anio}"
                fecha_fin = f"{calendar.monthrange(anio, mes_fin_val)[1]:02d}/{mes_fin_val:02d}/{anio}"
                resultado["fecha_filtro"] = f"{fecha_inicio} - {fecha_fin}"
                resultado["reportes_xml"] = reportes_xml
                estado_nombre = (ESTADOS_EMITIDOS_MAP.get(estado_emitidos, estado_emitidos) or "Sin Estado").strip() or "Sin Estado"
                estado_normalizado = unicodedata.normalize("NFKD", estado_nombre).encode("ascii", "ignore").decode("ascii")
                estado_slug = re.sub(r"[^A-Za-z0-9]+", "_", estado_normalizado).strip("_") or "Sin_Estado"
                tipo_visible = TIPOS_MAP.get(tipo, tipo)
                tipo_dir_nombre = _nombre_carpeta_tipo_visible(tipo_visible or tipo)
                tipo_slug = resultado.get(
                    "tipo_slug",
                    _slug_tipo(tipo_visible or tipo),
                )
                if "PDF" in formatos_norm:
                    sufijos_mes = [f"{anio:04d}{m:02d}" for m in range(mes_inicio, mes_fin_val + 1)]
                    carpeta_rango_pdf = destino_emitidos / estado_slug / tipo_dir_nombre / f"{anio:04d}" / "PDF"
                    reportes_pdf = _collect_existing_reports(
                        carpeta_rango_pdf,
                        "emitidos_reporte_pdf",
                        tipo_slug,
                        sufijos_mes,
                    )
                resultado["reportes_pdf"] = reportes_pdf
                if mes_inicio == 1 and mes_fin_val == 12:
                    if "XML" in formatos_norm:
                        base_anual = destino_emitidos / estado_slug / tipo_dir_nombre / f"{anio:04d}"
                        if base_anual.exists():
                            _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
                            xml_files = _xml_files_por_tipo(base_anual, tipo_prefijo)
                            estado_norm = estado_nombre.lower()
                            estado_default_reporte = estado_nombre if "no autoriz" in estado_norm else None
                            if xml_files:
                                xml_dir_anual = base_anual / "XML"
                                xml_dir_anual.mkdir(parents=True, exist_ok=True)
                                destino_anual_xml = xml_dir_anual / f"emitidos_reporte_xml_{tipo_slug}_{anio:04d}.xlsx"
                                try:
                                    construir_reporte(base_anual, destino_anual_xml, estado_default_reporte, xml_files=xml_files)
                                except Exception as err:
                                    print(f"[WARN] No se pudo construir reporte XML anual (emitidos): {err}")
                                if destino_anual_xml.exists():
                                    resultado["reporte_xml_anual"] = str(destino_anual_xml)
                            elif reportes_xml:
                                xml_dir_anual = base_anual / "XML"
                                xml_dir_anual.mkdir(parents=True, exist_ok=True)
                                destino_anual_xml = xml_dir_anual / f"emitidos_reporte_xml_{tipo_slug}_{anio:04d}.xlsx"
                                anual_xml = _consolidar_reportes_excel(reportes_xml, destino_anual_xml)
                                if anual_xml:
                                    resultado["reporte_xml_anual"] = str(anual_xml)
                    if reportes_pdf:
                        pdf_dir_anual = destino_emitidos / estado_slug / tipo_dir_nombre / f"{anio:04d}" / "PDF"
                        pdf_dir_anual.mkdir(parents=True, exist_ok=True)
                        destino_anual_pdf = pdf_dir_anual / f"emitidos_reporte_pdf_{tipo_slug}_{anio:04d}.xlsx"
                        anual_pdf = _consolidar_reportes_excel(reportes_pdf, destino_anual_pdf)
                        if anual_pdf:
                            resultado["reporte_pdf_anual"] = str(anual_pdf)
                    resultado["anual"] = True
                else:
                    if "XML" in formatos_norm:
                        base_rango = destino_emitidos / estado_slug / tipo_dir_nombre / f"{anio:04d}"
                        if base_rango.exists():
                            _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo)
                            meses_rango = range(mes_inicio, mes_fin_val + 1)
                            xml_files = _xml_files_por_meses(base_rango, tipo_prefijo, meses_rango)
                            estado_norm = estado_nombre.lower()
                            estado_default_reporte = estado_nombre if "no autoriz" in estado_norm else None
                            if xml_files:
                                sufijo_rango = f"{anio:04d}{mes_inicio:02d}{mes_fin_val:02d}"
                                xml_dir_rango = base_rango / "XML"
                                xml_dir_rango.mkdir(parents=True, exist_ok=True)
                                destino_rango_xml = xml_dir_rango / f"emitidos_reporte_xml_{tipo_slug}_{sufijo_rango}.xlsx"
                                try:
                                    construir_reporte(base_rango, destino_rango_xml, estado_default_reporte, xml_files=xml_files)
                                except Exception as err:
                                    print(f"[WARN] No se pudo construir reporte XML del rango (emitidos): {err}")
                                if destino_rango_xml.exists():
                                    resultado["reporte_xml_rango"] = str(destino_rango_xml)
                resultado.pop("reporte_pdf", None)
                resultado.pop("reporte_xml", None)
                carpeta_rango = destino_emitidos / estado_slug / tipo_dir_nombre / f"{anio:04d}"
                resultado["carpeta_tipo"] = str(carpeta_rango if carpeta_rango.exists() else destino_emitidos)
            else:
                mes_objetivo = int(resume_month if resume_download else mes)
                dia_objetivo = int(resume_day) if resume_download else int(dia)
                resultado = _emitidos_por_mes(mes_objetivo, 0 if dia in (0, None) else dia_objetivo)
            if isinstance(resultado, dict) and aviso_recorte:
                resultado.setdefault("aviso_recorte", aviso_recorte)
        else:
            if page.url != destino_url:
                try:
                    page.goto(destino_url, timeout=1000)
                except Exception:
                    pass
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=1000)
                except Exception:
                    pass
            destino_objetivo = destino_recibidos
            resultado = _flujo_recibidos(page, destino_objetivo, anio, mes, dia, tipo, formatos)
        if isinstance(resultado, dict):
            resultado.setdefault("carpeta_base", str(destino_objetivo))

        try:
            paginas_logout = []
            if modulo_page:
                paginas_logout.append(modulo_page)
            paginas_logout.append(page)
            for extra in context.pages:
                if extra not in paginas_logout:
                    paginas_logout.append(extra)
            for objetivo in paginas_logout:
                if _cerrar_sesion(objetivo):
                    break
        except Exception as err:
            print(f"[WARN] No se pudo cerrar la sesion del SRI: {err}")

        try:
            if using_persistent_profile:
                context.close()
            else:
                browser.close()
        except Exception:
            pass
        return resultado
