from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from pathlib import Path
from typing import Optional
import threading
from urllib.parse import urlencode
import pandas as pd
import csv, re, json, os, time, unicodedata, html, calendar, uuid
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

from robot._logging import get_logger

logger = get_logger(__name__)

# Estado compartido y señales movidas a robot/signals.py (Fase 1a del refactor).
# Se re-importan aquí para mantener la API pública estable: `aplicacion.py`
# y otros módulos siguen pudiendo hacer `from robot.downloader import
# set_user_notifier, request_cancel, ...` sin cambios.
from robot.signals import (
    CANCEL_EVENT,
    cancel_requested,
    clear_cancel,
    notify as _notify_user,
    request_cancel,
    set_user_notifier,
    _check_cancel,
    _notificar_usuario_accion,
    _notificar_usuario_captcha,
)

# Captcha (Sub-fase 2b): detección y resolución de captcha de imagen / reCAPTCHA.
# Las funciones se re-importan para mantener cualquier uso interno previo.
from robot.captcha import (
    CAPTCHA_INPUT_QUERY,
    CAPTCHA_INPUT_SELECTORS,
    _captcha_visible,
    _espera_captcha,
    _esperar_captcha_manual_input,
    _esperar_recaptcha_resuelto,
    _localizar_input_captcha,
    _recaptcha_challenge_activo,
    _recaptcha_presente,
    _resolver_captcha,
)

# Helpers de formato/parseo (Sub-fases 2c-ii y 3a-i): parseo de números/fechas,
# filas por defecto y normalización de campos de comprobantes emitidos.
from robot.data_formatters import (
    _emitidos_retencion_default_row,
    _factura_emitidos_default_row,
    _nota_credito_emitidos_default_row,
    _nota_debito_emitidos_default_row,
    _parse_decimal,
    _parse_datetime_local,
    _texto_emitidos_retencion,
    _texto_emitidos_retencion_na,
    _numero_emitidos_retencion,
    _normalizar_ambiente_retencion_emitidos,
    _normalizar_emision_retencion_emitidos,
    _formatear_fecha_autorizacion_retencion_emitidos,
    _detalle_retencion_emitidos,
    _asignar_resumen_retencion_emitidos,
    _map_retencion_legacy_to_emitidos_sample_row,
    _label_tipo_ident_emitidos_nota_credito,
    _label_ambiente_emitidos_retencion,
    _label_emision_emitidos_retencion,
    _label_forma_pago_emitidos_retencion,
)

# Generación de reportes Excel (Sub-fase 2c-ii-b): las 8 funciones que producen
# los .xlsx por tipo de comprobante. aplicacion.py importa 7 de ellas desde
# robot.downloader, así que se re-exportan vía facade.
from robot.reporting import (
    _consolidar_reportes_excel,
    _guardar_reporte_emitidos_excel,
    _guardar_reporte_pdf_excel,
    _guardar_reporte_pdf_factura_emitidos_excel,
    _guardar_reporte_pdf_nota_credito_emitidos_excel,
    _guardar_reporte_pdf_nota_debito_emitidos_excel,
    _guardar_reporte_pdf_retencion_emitidos_excel,
    _guardar_reporte_pdf_retencion_excel,
)

# Parsing de XML de comprobantes emitidos (Sub-fase 3a-ii). aplicacion.py
# importa _extraer_datos_xml_nota_credito_emitido, _extraer_datos_xml_nota_debito_emitido
# y _extraer_datos_xml_factura_emitido desde robot.downloader; se re-exportan.
from robot.xml_extraction import (
    _extraer_datos_xml_factura_emitido,
    _extraer_datos_xml_nota_credito_emitido,
    _extraer_datos_xml_nota_debito_emitido,
    _extraer_xml_emitidos_autorizacion,
    _strip_xml_namespaces,
)

# Parsing de PDF/XML (Sub-fases 3b-A y 3b-B): helpers de bajo nivel +
# parsers principales por tipo de comprobante. Se re-importan vía facade
# porque aplicacion.py y los flujos (_flujo_recibidos/_flujo_emitidos) los
# consumen como si vivieran en robot.downloader.
from robot.pdf_extraction import (
    # Sub-fase 3b-A: helpers de bajo nivel
    _normalizar_texto_pdf,
    _leer_texto_pdf,
    _es_archivo_pdf,
    _extraer_regex,
    _extraer_monto,
    _extraer_forma_pago,
    _extraer_seccion,
    _extraer_tipo_documento,
    _extraer_lineas_layout_pdf,
    _texto_linea_layout,
    _buscar_indice_linea_layout,
    _buscar_indice_linea_layout_exacta,
    _siguiente_linea_layout_no_vacia,
    _fecha_hora_pdf_a_iso,
    # Sub-fase 3b-B: parsers principales + helpers arrastrados
    _extraer_datos_pdf,
    _extraer_datos_pdf_retencion,
    _extraer_datos_xml_retencion,
    _extraer_datos_xml_retencion_emitido,
    _codigo_tipo_identificacion_desde_numero,
    _codigo_documento_sri,
    _combinar_rows_emitidos_especificos,
    _extraer_campos_adicionales_por_layout,
    _extraer_bloque_direccion_layout,
    _extraer_items_emitidos_layout,
    _formatear_descripciones_emitidos,
    _formatear_cantidad_emitidos,
    _formatear_precio_emitidos,
    _extraer_campos_adicionales_emitidos_desde_texto,
    _extraer_datos_pdf_retencion_emitido,
    _map_nota_credito_legada_a_emitidos_row,
    _map_nota_debito_legada_a_emitidos_row,
    _extraer_datos_pdf_nota_credito_emitido,
    _extraer_datos_pdf_nota_debito_emitido,
    _map_factura_legada_a_emitidos_row,
    _extraer_datos_pdf_factura_emitido,
    _extraer_datos_xml_liquidacion_compra_emitido,
    _extraer_datos_pdf_liquidacion_compra_emitido,
    _extraer_datos_xml_pdf_report,
    _extraer_datos_pdf_nota_credito,
    _extraer_datos_pdf_nota_debito,
    _limpiar_cdata,
    _buscar_autorizacion_en_json,
    _extraer_comprobante_desde_autorizacion,
    _valor_reporte_presente,
    _combinar_datos_reporte_emitidos,
    _extraer_datos_pdf_por_tipo_layout_first,
    _extraer_lineas_pdf_layout,
    _extraer_numero_desde_texto,
    _buscar_valor_layout,
    _extraer_datos_pdf_layout,
    _normalizar_token,
)

# Automatización del navegador (Sub-fase 3c): navegación, ViewState/JSF,
# selección de combos, descarga de PDF/XML, parsing de respuestas del portal.
# Se re-importan vía facade: login y los flujos (que siguen en downloader.py)
# las usan como si vivieran aquí.
from robot.browser import (
    _abrir_modulo_consultas,
    _actualizar_view_state_input,
    _asegurar_portal_disponible,
    _capturar_xml_emitido,
    _capturar_xml_emitido_por_dialogo,
    _cerrar_modal_encuesta,
    _cerrar_sesion,
    _click_consultar_emitidos,
    _click_texto,
    _construir_nombre_xml_emitido,
    _descargar_pdf_emitidos_post,
    _descargar_pdf_emitidos_post_con_viewstate,
    _descargar_pdf_recibidos_post,
    _descargar_pdf_recibidos_post_con_viewstate,
    _descargar_xml_emitido_por_clave,
    _es_respuesta_pdf,
    _es_respuesta_xml,
    _esperar_ajax,
    _extraer_autorizacion_desde_partial,
    _extraer_datos_emitidos_dom,
    _extraer_detalle_emitido_desde_partial,
    _fecha_slug,
    _guardar_pdf_desde_enlace,
    _guardar_pdf_desde_jsf,
    _guardar_xml_desde_enlace,
    _inferir_iva_columna,
    _mapear_detalle_emitido_a_pdf,
    _obtener_detalle_emitido_xhr,
    _obtener_form_base_emitidos,
    _obtener_source_detalle_emitido,
    _obtener_view_state,
    _parse_emitido_comprobante,
    _portal_indisponible,
    _primer_texto,
    _rellenar_input_por_label,
    _resolver_autenticacion_persistente,
    _resolver_destino_unico,
    _seleccionar,
    _seleccionar_en_select,
    _seleccionar_por_label,
    _strip_html,
    _texto_probable_comprador_desde_fila,
    _total_slug,
)

# Definiciones de columnas Excel (Sub-fase 2c-i): listas, sets y mapeos que
# describen los reportes generados. Re-exportadas porque aplicacion.py importa
# PDF_REPORT_COLUMNS, RETENCION_REPORT_COLUMNS, EMITIDOS_RETENCION_REPORT_COLUMNS
# y EMITIDOS_FACTURA_REPORT_COLUMNS directamente desde robot.downloader.
from robot.report_columns import (
    EMITIDOS_FACTURA_NUMERIC_COLUMNS,
    EMITIDOS_FACTURA_REPORT_COLUMNS,
    EMITIDOS_FACTURA_TEXT_FORCE_COLUMNS,
    EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS,
    EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS,
    EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS,
    EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL,
    EMITIDOS_NOTA_DEBITO_NUMERIC_COLUMNS,
    EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS,
    EMITIDOS_NOTA_DEBITO_TEXT_FORCE_COLUMNS,
    EMITIDOS_RETENCION_AMBIENTE_LABEL,
    EMITIDOS_RETENCION_DOC_CODE_LABEL,
    EMITIDOS_RETENCION_FORMA_PAGO_LABEL,
    EMITIDOS_RETENCION_NUMERIC_COLUMNS,
    EMITIDOS_RETENCION_REPORT_COLUMNS,
    EMITIDOS_RETENCION_TEXT_FORCE_COLUMNS,
    EMITIDOS_RETENCION_TIPO_EMISION_LABEL,
    PDF_REPORT_COLUMNS,
    RETENCION_REPORT_COLUMNS,
)

# Utilidades de archivos / paths / parsing TXT extraídas a robot/file_utils.py
# (Fase 1b del refactor). Se re-importan para que el resto del módulo siga
# pudiendo usarlas como si vivieran aquí.
from robot.file_utils import (
    _collect_existing_reports,
    _delete_report_files,
    _detectar_delimitador,
    _es_clave,
    _extraer_claves_desde_txt,
    _mes_a_texto,
    _sanear_nombre_archivo,
)

# Helpers de tipos de comprobante extraídos a robot/comprobante_types.py
# (Fase 2a del refactor). Tanto _slug_tipo como _prefijo_tipo son consumidos
# por aplicacion.py via `from robot.downloader import ...`, por lo que se
# re-exportan aquí.
from robot.comprobante_types import (
    TIPO_LABEL_MAP,
    _coincide_tipo_documental,
    _es_tipo_factura,
    _es_tipo_liquidacion_compra,
    _es_tipo_nota_credito,
    _es_tipo_nota_debito,
    _es_tipo_retencion,
    _formatear_label,
    _nombre_carpeta_tipo,
    _nombre_carpeta_tipo_visible,
    _normalizar_label_simple,
    _normalizar_tipo_clave,
    _prefijo_tipo,
    _resolver_tipo_label,
    _slug_tipo,
)

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
# Movida a robot/config.py (Fase 1c del refactor). Se re-importa todo para
# mantener compatible la API previa: aplicacion.py sigue pudiendo hacer
# `from robot.downloader import MANUAL_CONSULTA_RECIBIDOS, TIPOS_MAP,
# ESTADOS_EMITIDOS_MAP` sin cambios.
from robot.config import (
    AUTORIZACION_COMPROBANTES_SOAP_URL,
    CONSULTAS_SELECTOR,
    DEVTOOLS,
    DOC_LABELS,
    DOWNLOAD_TIMEOUT,
    EMITIDOS_RESET_AFTER_DAY_DOCS,
    EMITIDOS_RESET_PAUSE_MS,
    ESTADOS_EMITIDOS_MAP,
    FACTURACION_MENU_SELECTOR,
    HEADLESS,
    MANUAL_CONSULTA_RECIBIDOS,
    MENU_EMITIDOS_TRIGGER_URL,
    MENU_TOGGLE_SELECTOR,
    MENU_URL,
    MENU_URL_ALT,
    MODULO_PRODUCCION_SELECTOR,
    OVERLAY_SELECTORS,
    PAUSE_AT_LOGIN,
    PAUSE_BEFORE_CONSULTAR,
    PAUSE_BEFORE_CONSULTAR_SECONDS,
    PAUSE_PROMPT,
    PORTAL_HOME,
    PORTAL_INDISPONIBLE_MENSAJE,
    RECIBIDOS_AUTO_POST_EXECUTE_MS,
    RECIBIDOS_AUTO_PRE_EXECUTE_MS,
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS,
    RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC,
    RECIBIDOS_CONSULTA_INTENTOS,
    RECIBIDOS_DIRECT_URL,
    RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS,
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO,
    RECIBIDOS_REHIDRATAR_ON_CAPTCHA,
    RECUPERAR_COMPROBANTES_URL,
    SLOW_MO,
    TIPOS_MAP,
    URLS,
    USE_PERSISTENT_PROFILE,
    USER_DATA_DIR,
)

# Las 19 constantes de columnas Excel (PDF_REPORT_COLUMNS, RETENCION_REPORT_COLUMNS,
# EMITIDOS_*_REPORT_COLUMNS, *_TEXT_FORCE_COLUMNS, *_NUMERIC_COLUMNS, *_LABEL)
# fueron movidas a robot/report_columns.py (Sub-fase 2c-i del refactor).
# Se re-importan al final de la cabecera para que `aplicacion.py` siga
# pudiendo hacer `from robot.downloader import PDF_REPORT_COLUMNS, ...`.







def _map_retencion_legada_a_emitidos_row(legacy: dict | None) -> dict:
    return _map_retencion_legacy_to_emitidos_sample_row(legacy)













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

def _nombre_documento_mes(tipo_slug: str, fecha_token: str, nombre_base: str) -> str:
    tipo_parte = _slug_tipo(tipo_slug or "") or "documento"
    fecha_parte = re.sub(r"[^0-9]", "", fecha_token or "") or "00000000"
    base = _sanear_nombre_archivo(nombre_base or "archivo")
    return f"{tipo_parte}__{fecha_parte}__{base}"



# _notificar_usuario_captcha y _notificar_usuario_accion movidos a robot/signals.py
# (Sub-fase 2b del refactor). Quedan re-importados en la cabecera para mantener
# compatibilidad con cualquier llamada interna.



def _es_url_autorizacion(url: str) -> bool:
    if not url:
        return False
    url_lower = url.lower()
    if "sri.gob.ec" not in url_lower:
        return False
    return "autoriz" in url_lower







# _parse_datetime_local movido a robot/data_formatters.py (Sub-fase 2c-ii-b).
# Se re-importa en la cabecera del módulo.

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





















DOWNLOAD_ROW_RETRY_ATTEMPTS = max(1, int(os.getenv("DOWNLOAD_ROW_RETRY_ATTEMPTS", "2")))

def _build_download_row_id(*parts) -> str:
    tokens = []
    for part in parts:
        text = str(part or "").strip()
        if text:
            tokens.append(text)
    return " | ".join(tokens) if tokens else str(uuid.uuid4())

def _build_download_verification(
    registros_esperados: int,
    esperados_xml: set,
    descargados_xml: set,
    esperados_pdf: set,
    descargados_pdf: set,
) -> dict:
    faltantes_xml = max(0, len(esperados_xml) - len(descargados_xml))
    faltantes_pdf = max(0, len(esperados_pdf) - len(descargados_pdf))
    descarga_completa = faltantes_xml == 0 and faltantes_pdf == 0
    partes = [f"Filas detectadas: {int(registros_esperados or 0)}"]
    if esperados_xml:
        partes.append(f"XML {len(descargados_xml)}/{len(esperados_xml)}")
    if esperados_pdf:
        partes.append(f"PDF {len(descargados_pdf)}/{len(esperados_pdf)}")
    mensaje = " | ".join(partes)
    if not descarga_completa:
        mensaje += f" | Faltantes XML: {faltantes_xml}, PDF: {faltantes_pdf}"
    return {
        "registros_esperados": int(registros_esperados or 0),
        "esperados_xml": len(esperados_xml),
        "esperados_pdf": len(esperados_pdf),
        "descargados_xml_verificados": len(descargados_xml),
        "descargados_pdf_verificados": len(descargados_pdf),
        "faltantes_xml": faltantes_xml,
        "faltantes_pdf": faltantes_pdf,
        "descarga_completa": descarga_completa,
        "mensaje_verificacion": mensaje,
    }

def _merge_download_verification(resultados: list[dict]) -> dict:
    registros_esperados = sum(
        int((r or {}).get("registros_esperados", (r or {}).get("n_registros", 0)) or 0)
        for r in resultados
    )
    esperados_xml = sum(int((r or {}).get("esperados_xml", 0) or 0) for r in resultados)
    esperados_pdf = sum(int((r or {}).get("esperados_pdf", 0) or 0) for r in resultados)
    descargados_xml = sum(
        int((r or {}).get("descargados_xml_verificados", (r or {}).get("n_xml", 0)) or 0)
        for r in resultados
    )
    descargados_pdf = sum(
        int((r or {}).get("descargados_pdf_verificados", (r or {}).get("n_pdf", 0)) or 0)
        for r in resultados
    )
    return {
        "registros_esperados": registros_esperados,
        "esperados_xml": esperados_xml,
        "esperados_pdf": esperados_pdf,
        "descargados_xml_verificados": descargados_xml,
        "descargados_pdf_verificados": descargados_pdf,
        "faltantes_xml": max(0, esperados_xml - descargados_xml),
        "faltantes_pdf": max(0, esperados_pdf - descargados_pdf),
        "descarga_completa": descargados_xml >= esperados_xml and descargados_pdf >= esperados_pdf,
        "mensaje_verificacion": (
            f"Filas detectadas: {registros_esperados}"
            + (f" | XML {descargados_xml}/{esperados_xml}" if esperados_xml else "")
            + (f" | PDF {descargados_pdf}/{esperados_pdf}" if esperados_pdf else "")
        ),
    }

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
                        logger.warning("No se pudo accionar el boton 'Ingresar'; el objeto page no expone teclado.")

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
                logger.info(f"Reintentando login por captcha adicional ({captcha_retry}/{CAPTCHA_MAX_ATTEMPTS}).")
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
                logger.info("Recibidos configurado en modo submit unico (evita doble XHR por clic).")
        except Exception as err:
            logger.warning(f"No se pudo configurar submit unico en Recibidos: {err}")

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
            logger.warning(f"No se pudo recargar Recibidos para reintentar captcha: {err}")
            return False
        ok = _aplicar_filtros_recibidos(estricto=False)
        if not ok:
            logger.warning("No se pudieron reaplicar filtros de Recibidos al reintentar captcha.")
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
                        logger.info(f"Esperando {espera:.1f}s antes de reintento manual.")
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
                    logger.info(f"Esperando {espera:.1f}s antes de reintento manual.")
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
                    logger.info(f"Esperando {espera:.1f}s antes de reintentar Recibidos.")
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
                logger.info(f"Sin tabla tras intento {intento}/{intentos}. Espera {espera:.1f}s.")
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
    registros_esperados = 0
    esperados_xml = set()
    esperados_pdf = set()
    descargados_xml = set()
    descargados_pdf = set()

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
                logger.warning(f"No se pudo descargar el reporte TXT/XML: {err}")
        else:
            logger.warning("No se encontro el enlace 'Descargar reporte' para XML.")

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
                clave_fila = _extraer_clave_fila(celdas)
                razon_texto = celdas.nth(1).inner_text().strip()
                bloques = [segmento.strip() for segmento in razon_texto.splitlines() if segmento.strip()]
                razon_social = bloques[-1] if bloques else f"documento_{pagina}_{idx+1}"
                nombre_base = _nombre_documento_mes(tipo_slug, fecha_token_doc, razon_social)
                row_id = _build_download_row_id(
                    clave_fila,
                    tipo_slug,
                    razon_social,
                    f"pag{pagina}",
                    f"fila{idx+1}",
                )
                registros_esperados += 1
                if descargar_xml:
                    esperados_xml.add(row_id)
                if descargar_pdf:
                    esperados_pdf.add(row_id)

                xml_guardado = False
                xml_path_report = None
                if descargar_xml_para_reporte and not xml_guardado:
                    for intento_xml in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
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
                                    descargados_xml.add(row_id)
                                else:
                                    xml_temp_paths.append(xml_path_report)
                                xml_guardado = True
                                lote_xml_ok += 1
                                break

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
                                    descargados_xml.add(row_id)
                                else:
                                    xml_temp_paths.append(xml_path_report)
                                xml_guardado = True
                                lote_xml_ok += 1
                                break
                            except Exception as err:
                                logger.warning(f"No se pudo descargar XML para '{razon_social}' (intento {intento_xml}/{DOWNLOAD_ROW_RETRY_ATTEMPTS}): {err}")
                        if xml_guardado:
                            break
                        if intento_xml < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                            try:
                                page.wait_for_timeout(250)
                            except Exception:
                                pass

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
                        logger.warning(f"No se pudo procesar XML para reporte: {err}")

                if descargar_pdf:
                    pdf_guardado = False
                    for intento_pdf in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
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
                            descargados_pdf.add(row_id)
                            lote_pdf_ok += 1
                            pdf_guardado = True
                            if resultado_pdf.suffix.lower() == ".pdf" and _es_archivo_pdf(resultado_pdf):
                                if not usar_xml_reporte:
                                    datos_pdf = _extraer_datos_pdf_por_tipo_layout_first(
                                        resultado_pdf,
                                        es_retencion=es_retencion,
                                        es_nota_credito=es_nota_credito,
                                        es_nota_debito=es_nota_debito,
                                    )
                                    pdf_report_rows.append(datos_pdf)
                            break
                        if intento_pdf < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                            try:
                                page.wait_for_timeout(250)
                            except Exception:
                                pass
                    if not pdf_guardado:
                        logger.warning(f"No se pudo descargar PDF para '{razon_social}': no se obtuvo archivo.")

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
            logger.info(f"Pag {pagina} completa: {total_filas} filas en {duracion_pagina:.2f}s")

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
                logger.warning("No se pudo construir el reporte PDF de retenciones (recibidos).")
        elif _guardar_reporte_pdf_excel(pdf_report_rows, pdf_report_path):
            reporte_pdf_path = pdf_report_path
        else:
            logger.warning("No se pudo construir el reporte PDF de recibidos.")
    resultado = {
        "estado": "ok",
        "n_xml": n_xml,
        "n_pdf": n_pdf,
        "n_registros": registros_esperados,
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
    resultado.update(
        _build_download_verification(
            registros_esperados,
            esperados_xml,
            descargados_xml,
            esperados_pdf,
            descargados_pdf,
        )
    )
    if not resultado.get("descarga_completa", True):
        logger.warning(f"Verificacion de Recibidos incompleta: {resultado.get('mensaje_verificacion')}")
    else:
        logger.info(f"Verificacion de Recibidos OK: {resultado.get('mensaje_verificacion')}")

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
            logger.warning(f"No se pudo seleccionar el tipo de comprobante '{tipo_visible}' en Emitidos.")

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
                logger.warning(f"No se pudo completar la fecha de emision con '{fecha_emision}' en Emitidos.")

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
                logger.warning(f"No se pudo establecer el establecimiento '{est_valor}' en Emitidos.")

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
                logger.warning(f"No se pudo establecer el punto de emision '{punto_valor}' en Emitidos.")

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
    registros_esperados = 0
    esperados_xml = set()
    esperados_pdf = set()
    descargados_xml = set()
    descargados_pdf = set()

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
        info_base.update(_build_download_verification(0, set(), set(), set(), set()))
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
            filas = tabla_emitidos.locator("tr")
            total_filas = filas.count()
            lote_inicio = time.perf_counter()
            lote_contador = 0
            lote_xml_ok = 0
            lote_pdf_ok = 0
            for idx in range(total_filas):
                _check_cancel("emitidos_fila")
                fila = filas.nth(idx)
                celdas = fila.locator("td")
                try:
                    total_celdas = celdas.count()
                except Exception:
                    total_celdas = 0
                if total_celdas < 2:
                    continue
                try:
                    tipo_serie_texto = celdas.nth(1).inner_text().strip()
                except Exception:
                    tipo_serie_texto = ""
                tipo_detectado = _extraer_tipo_documento(tipo_serie_texto)
                if tipo_detectado and not _coincide_tipo_documental(tipo_visible or tipo, tipo_detectado):
                    print(
                        f"[WARN] Se omitio una fila de Emitidos porque corresponde a '{tipo_detectado}' y no a '{tipo_visible or tipo}'."
                    )
                    continue
                clave_texto = _extraer_clave_fila(celdas)
                try:
                    razon_texto = celdas.nth(4).inner_text().strip() if total_celdas > 4 else ""
                except Exception:
                    razon_texto = ""
                try:
                    fecha_aut_texto = celdas.nth(3).inner_text().strip() if total_celdas > 3 else ""
                except Exception:
                    fecha_aut_texto = ""
                try:
                    valor_sin_imp_texto = celdas.nth(5).inner_text().strip() if total_celdas > 5 else ""
                except Exception:
                    valor_sin_imp_texto = ""
                try:
                    iva_texto = celdas.nth(6).inner_text().strip() if total_celdas > 6 else ""
                except Exception:
                    iva_texto = ""
                try:
                    importe_total_texto = celdas.nth(7).inner_text().strip() if total_celdas > 7 else ""
                except Exception:
                    importe_total_texto = ""

                tipo_serie_completo = " ".join(
                    fragment for fragment in [tipo_serie_texto, clave_texto] if fragment
                ).strip()
                tipo_slug_archivo = _slug_tipo(tipo_detectado or tipo_visible or tipo) or tipo_slug
                nombre_base_pdf = _nombre_documento_mes(
                    tipo_slug_archivo,
                    fecha_token_doc,
                    tipo_serie_completo or razon_texto or f"emitido_{pagina}_{idx+1}",
                )
                row_id = _build_download_row_id(
                    clave_texto,
                    tipo_slug_archivo,
                    tipo_serie_texto,
                    razon_texto,
                    f"pag{pagina}",
                    f"fila{idx+1}",
                )
                registros_esperados += 1
                if descargar_xml and not omitir_soap_xml:
                    esperados_xml.add(row_id)
                if descargar_pdf:
                    esperados_pdf.add(row_id)
                xml_path_report = None

                if es_rechazado:
                    try:
                        if descargar_xml_para_reporte:
                            if clave_texto:
                                for intento_xml in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
                                    try:
                                        resultado_xml = _descargar_xml_emitido_por_clave(
                                            request_context,
                                            clave_texto,
                                            xml_dir,
                                            nombre_base_pdf,
                                            claves_guardadas,
                                        )
                                        if resultado_xml:
                                            xml_path_report = resultado_xml
                                            if descargar_xml:
                                                n_xml += 1
                                                descargados_xml.add(row_id)
                                            else:
                                                xml_temp_paths.append(xml_path_report)
                                            lote_xml_ok += 1
                                            break
                                    except Exception as err:
                                        print(
                                            f"[WARN] No se pudo obtener XML SOAP para '{nombre_base_pdf}' "
                                            f"(intento {intento_xml}/{DOWNLOAD_ROW_RETRY_ATTEMPTS}): {err}"
                                        )
                                    if intento_xml < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                                        try:
                                            page.wait_for_timeout(250)
                                        except Exception:
                                            pass
                            else:
                                print(
                                    f"[WARN] La fila '{nombre_base_pdf}' no tiene clave de acceso para solicitar el XML."
                                )

                        if descargar_pdf:
                            # En "No Autorizados" el icono XML suele tener id :lnkPdf aunque sea XML.
                            enlace_pdf = fila.locator(
                                "a[id$=':lnkPdf'], a[title*='pdf' i], img[alt*='pdf' i], img[title*='pdf' i]"
                            )
                            if enlace_pdf.count():
                                contenedor = enlace_pdf.first
                                # Si es <img>, buscamos su <a> ancestro
                                if contenedor.locator("xpath=ancestor::a[1]").count():
                                    contenedor = contenedor.locator("xpath=ancestor::a[1]").first
                                destino_pdf = pdf_dir / f"{nombre_base_pdf}.pdf"
                                link_id = None
                                try:
                                    link_id = contenedor.get_attribute("id")
                                except Exception:
                                    link_id = None
                                resultado_pdf = None
                                for intento_pdf in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
                                    resultado_pdf = None
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
                                    if not resultado_pdf:
                                        resultado_pdf = _guardar_pdf_desde_jsf(page, contenedor, destino_pdf)
                                    if not resultado_pdf:
                                        resultado_pdf = _guardar_pdf_desde_enlace(page, contenedor, destino_pdf)
                                    if resultado_pdf:
                                        break
                                    if intento_pdf < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                                        try:
                                            page.wait_for_timeout(250)
                                        except Exception:
                                            pass
                                if resultado_pdf:
                                    n_pdf += 1
                                    descargados_pdf.add(row_id)
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
                                                logger.warning(f"No se pudo usar XML para el reporte PDF: {err}")
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
                        logger.warning(f"No se pudo descargar XML/PDF para '{nombre_base_pdf}': {err}")
                    continue

                if descargar_xml_para_reporte and not omitir_soap_xml:
                    if clave_texto:
                        for intento_xml in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
                            try:
                                resultado_xml = _descargar_xml_emitido_por_clave(
                                    request_context,
                                    clave_texto,
                                    xml_dir,
                                    nombre_base_pdf,
                                    claves_guardadas,
                                )
                                if resultado_xml:
                                    xml_path_report = resultado_xml
                                    if descargar_xml:
                                        n_xml += 1
                                        descargados_xml.add(row_id)
                                    else:
                                        xml_temp_paths.append(xml_path_report)
                                    lote_xml_ok += 1
                                    break
                            except Exception as err:
                                print(
                                    f"[WARN] No se pudo obtener XML SOAP para '{nombre_base_pdf}' "
                                    f"(intento {intento_xml}/{DOWNLOAD_ROW_RETRY_ATTEMPTS}): {err}"
                                )
                            if intento_xml < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                                try:
                                    page.wait_for_timeout(250)
                                except Exception:
                                    pass
                    else:
                        logger.warning(f"La fila '{nombre_base_pdf}' no tiene clave de acceso para solicitar el XML.")

                if descargar_pdf:
                    link_pdf = fila.locator("a[id$=':lnkPdf']")
                    if not link_pdf.count():
                        link_pdf = fila.locator("a[title*='pdf' i], button[title*='pdf' i]")
                    if not link_pdf.count():
                        continue

                    destino_pdf = pdf_dir / f"{nombre_base_pdf}.pdf"
                    link_id = None
                    try:
                        link_id = link_pdf.first.get_attribute("id")
                    except Exception:
                        link_id = None
                    resultado_pdf = None
                    for intento_pdf in range(1, DOWNLOAD_ROW_RETRY_ATTEMPTS + 1):
                        resultado_pdf = None
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
                        if not resultado_pdf:
                            resultado_pdf = _guardar_pdf_desde_jsf(page, link_pdf.first, destino_pdf)
                        if not resultado_pdf:
                            resultado_pdf = _guardar_pdf_desde_enlace(page, link_pdf.first, destino_pdf)
                        if resultado_pdf:
                            break
                        if intento_pdf < DOWNLOAD_ROW_RETRY_ATTEMPTS:
                            try:
                                page.wait_for_timeout(250)
                            except Exception:
                                pass
                    if resultado_pdf:
                        n_pdf += 1
                        descargados_pdf.add(row_id)
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
                                    logger.warning(f"No se pudo usar XML para el reporte PDF: {err}")
                            if datos_pdf is None:
                                if es_retencion:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_retencion_emitido(resultado_pdf)
                                    except Exception as err:
                                        logger.warning(f"No se pudo leer el PDF de retención para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_nota_credito:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_nota_credito_emitido(resultado_pdf)
                                    except Exception as err:
                                        logger.warning(f"No se pudo leer el PDF de nota de crédito para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_nota_debito:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_nota_debito_emitido(resultado_pdf)
                                    except Exception as err:
                                        logger.warning(f"No se pudo leer el PDF de nota de débito para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_factura_emitida:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_factura_emitido(resultado_pdf)
                                    except Exception as err:
                                        logger.warning(f"No se pudo leer el PDF de factura para completar el reporte: {err}")
                                        datos_pdf = None
                                elif es_liquidacion_compra:
                                    try:
                                        datos_pdf = _extraer_datos_pdf_liquidacion_compra_emitido(resultado_pdf)
                                    except Exception as err:
                                        logger.warning(f"No se pudo leer el PDF de liquidación de compra para completar el reporte: {err}")
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
                                    source_id_detalle = _obtener_source_detalle_emitido(page, idx)
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
                                        logger.warning(f"No se pudo leer el PDF para completar el reporte: {err}")
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
                        logger.warning(f"No se pudo descargar PDF para '{nombre_base_pdf}': no se obtuvo archivo.")

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
            logger.info(f"Pag {pagina} completa: {total_filas} filas en {duracion_pagina:.2f}s")

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
                logger.warning(f"No se pudo construir el reporte XML de emitidos: {err}")

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
            "n_registros": registros_esperados or len(df),
        })
    else:
        info_base.update({
            "estado": "sin_resultados",
            "n_registros": 0,
        })
    info_base.update(
        _build_download_verification(
            registros_esperados or len(df),
            esperados_xml,
            descargados_xml,
            esperados_pdf,
            descargados_pdf,
        )
    )
    if not info_base.get("descarga_completa", True):
        logger.warning(f"Verificacion de Emitidos incompleta: {info_base.get('mensaje_verificacion')}")
    else:
        logger.info(f"Verificacion de Emitidos OK: {info_base.get('mensaje_verificacion')}")
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
                logger.warning(f"No se pudo usar perfil persistente; fallback a contexto normal: {err}")
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
                resultados_verificacion = []
                resultado_mes = None

                for idx_dia, dia_iter in enumerate(dias_consultar):
                    _check_cancel("recibidos_dia")
                    resultado_dia = _consultar_recibidos_dia(mes_actual, dia_iter)
                    total_xml += resultado_dia.get("n_xml", 0)
                    total_pdf += resultado_dia.get("n_pdf", 0)
                    resultados_verificacion.append(resultado_dia)
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
                resultado_mes.update(_merge_download_verification(resultados_verificacion))
                resultado_mes["n_registros"] = resultado_mes.get("registros_esperados", 0)

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
                            logger.warning(f"No se pudo construir el reporte XML mensual de recibidos: {err}")
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
                            logger.warning(f"No se pudo consolidar reporte PDF mensual de recibidos: {err}")
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
                resultados_verificacion = []
                mes_inicio = int(resume_month if resume_download else mes)
                mes_fin_val = int(mes_fin_val)
                resultado_mes = None
                for mes_actual in range(mes_inicio, mes_fin_val + 1):
                    _check_cancel("recibidos_mes")
                    resultado_mes = _recibidos_por_mes(mes_actual, 0)
                    total_xml += resultado_mes.get("n_xml", 0)
                    total_pdf += resultado_mes.get("n_pdf", 0)
                    resultados_verificacion.append(resultado_mes)
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
                resultado.update(_merge_download_verification(resultados_verificacion))
                resultado["n_registros"] = resultado.get("registros_esperados", 0)
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
                                    logger.warning(f"No se pudo construir reporte XML anual (recibidos): {err}")
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
                                    logger.warning(f"No se pudo construir reporte XML del rango (recibidos): {err}")
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
                    logger.warning(f"No se pudo volver al menu principal antes del reinicio de Emitidos: {err}")
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
                resultados_verificacion = []
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
                    resultados_verificacion.append(resultado_dia)
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
                    resultado_mes.update(_merge_download_verification(resultados_verificacion))
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
                                logger.warning(f"No se pudo construir el reporte XML mensual de emitidos: {err}")
                    if descargar_pdf_mes:
                        reportes_dia = list(reportes_pdf_dia)
                    if descargar_pdf_mes and reportes_dia:
                        frames = []
                        for ruta_excel in reportes_dia:
                            try:
                                df_dia = pd.read_excel(ruta_excel)
                            except Exception as err:
                                logger.warning(f"No se pudo leer reporte diario '{ruta_excel}': {err}")
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
                resultados_verificacion = []
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
                    resultados_verificacion.append(resultado_mes)
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
                resultado.update(_merge_download_verification(resultados_verificacion))
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
                                    logger.warning(f"No se pudo construir reporte XML anual (emitidos): {err}")
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
                                    logger.warning(f"No se pudo construir reporte XML del rango (emitidos): {err}")
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
            logger.warning(f"No se pudo cerrar la sesion del SRI: {err}")

        try:
            if using_persistent_profile:
                context.close()
            else:
                browser.close()
        except Exception:
            pass
        return resultado
