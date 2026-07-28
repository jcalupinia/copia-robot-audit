from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from pathlib import Path
from typing import Optional
import threading
from urllib.parse import urlencode
import pandas as pd
import csv, re, json, os, time, unicodedata, html, calendar, uuid
import tempfile
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
    CAPTCHA_MAX_ATTEMPTS,
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

# Flujos de descarga (Fase 4): el corazón del bot. `descargar_sri` (que sigue
# en este módulo) los invoca tras autenticar. `_xml_files_por_meses` (helper
# que quedó aquí) usa `_xml_files_por_tipo`; aplicacion.py lo importa desde
# robot.downloader, así que se re-exporta vía facade.
from robot.workflows import (
    _flujo_recibidos,
    _flujo_emitidos,
    _xml_files_por_tipo,
    _build_download_verification,
    _build_download_row_id,
    _debe_omitir_soap_xml,
    _extraer_clave_fila,
    _nombre_documento_mes,
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
    RECIBIDOS_CLEAN_GOOGLE_COOKIES,
    PLAYWRIGHT_USE_TEMP_PROFILE,
    PREFER_SYSTEM_CHROME,
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











# ====== Funciones auxiliares ======

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

def _detectar_credenciales_invalidas_sri(page) -> str:
    """Devuelve el texto del mensaje de error si Keycloak reporta credenciales
    invalidas (usuario o clave incorrectos), o cadena vacia si no hay tal error.

    Se chequean los contenedores tipicos de Keycloak: `.alert-error`,
    `.kc-feedback-text`, `#input-error`, mas el body completo como fallback
    cuando los selectores no matchean (algunos themes del SRI usan markup propio).
    """
    selectores = [
        ".alert-error",
        ".kc-feedback-text",
        "#input-error",
        ".pf-c-alert__title",
        ".alert.alert-danger",
        ".error-message",
        ".mensaje-error",
    ]
    for sel in selectores:
        try:
            loc = page.locator(sel)
            if loc.count() and loc.first.is_visible():
                txt = (loc.first.inner_text(timeout=1000) or "").strip()
                if txt:
                    low = txt.lower()
                    if any(k in low for k in (
                        "usuario o clave",
                        "clave incorrec",
                        "usuario incorrec",
                        "credencial",
                        "invalid user",
                        "invalid password",
                        "datos incorrect",
                        "datos invalid",
                    )):
                        return txt
        except Exception:
            continue
    return ""


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
            page.goto(destino_url, timeout=30000)
        except Exception:
            pass
        try:
            page.wait_for_load_state("domcontentloaded", timeout=15000)
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
            timeout=30000,
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
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass

            # Si el SRI ya pinto el mensaje de "Usuario o clave incorrectos",
            # cortamos de una. Se usa un prefijo `[CREDENCIALES]` para que la
            # UI distinga este caso de un timeout/captcha y muestre el modal
            # correspondiente, sin reintentos automaticos.
            msg_cred = _detectar_credenciales_invalidas_sri(page)
            if msg_cred:
                raise RuntimeError(f"[CREDENCIALES] {msg_cred}")

            autenticado = False
            for intento in range(2):
                try:
                    page.wait_for_url(lambda url: "auth/realms" not in url, timeout=15000)
                    autenticado = True
                    break
                except PlaywrightTimeoutError:
                    # Antes de declarar persistente, revisamos si el portal
                    # nos esta diciendo que las credenciales son invalidas.
                    msg_cred = _detectar_credenciales_invalidas_sri(page)
                    if msg_cred:
                        raise RuntimeError(f"[CREDENCIALES] {msg_cred}")
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

            # Ultimo chequeo antes de salir con error generico: si vemos
            # mensaje de credenciales invalidas, lo reportamos especificamente.
            msg_cred = _detectar_credenciales_invalidas_sri(page)
            if msg_cred:
                raise RuntimeError(f"[CREDENCIALES] {msg_cred}")
            raise RuntimeError("No fue posible completar el login del SRI (credenciales o captcha).")
        if "auth/realms" in page.url:
            if not _resolver_autenticacion_persistente(page):
                raise RuntimeError("No fue posible completar el login del SRI (pantalla de autenticacion persistente).")
            try:
                page.wait_for_url(lambda url: "auth/realms" not in url, timeout=15000)
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

# ============================================================
# ?? LECTURA DE TABLA PARA COMPROBANTES EMITIDOS (sin TXT)
# ============================================================

# ============================================================
# APERTURA DEL NAVEGADOR
# ============================================================
def _limpiar_locks_perfil(profile_dir: Path) -> None:
    """Elimina los archivos de lock que Chrome/Chromium deja en un perfil.

    Si una corrida anterior no cerró limpio, quedan `SingletonLock` y
    similares; sin borrarlos, el siguiente `launch_persistent_context`
    falla con "Target page, context or browser has been closed".
    """
    for lock in ("SingletonLock", "SingletonCookie", "SingletonSocket", "lockfile"):
        try:
            (profile_dir / lock).unlink(missing_ok=True)
        except Exception:
            pass


def _limpiar_cookies_google_perfil(profile_dir: Path) -> int:
    """Borra cookies de Google/reCAPTCHA del perfil persistente.

    Resetea el "score base" que reCAPTCHA Enterprise mantiene asociado a
    las cookies del perfil, SIN tocar las cookies del SRI ni de ningun otro
    sitio. Util cuando el perfil acumula mala reputacion tras muchas
    sesiones automatizadas y el captcha vuelve a fallar.

    Solo borra cookies cuyos `host_key` pertenecen a:
      - `*.google.com`, `accounts.google.com`, `apis.google.com`
      - `*.recaptcha.net`
      - `*.gstatic.com`
      - `*.googleapis.com`
      - `*.googleusercontent.com`

    Devuelve la cantidad de cookies borradas. Si el archivo no existe
    o esta bloqueado (Chrome corriendo), devuelve 0 sin error.

    IMPORTANTE: debe llamarse ANTES de lanzar el navegador. Si Chrome
    esta abierto, el archivo Cookies esta bloqueado y la operacion falla
    silenciosamente.
    """
    import sqlite3

    # Chrome guarda cookies en distintas rutas segun la version:
    #   - Chrome <= 120: <profile>/Default/Cookies o <profile>/Cookies
    #   - Chrome >= 121: <profile>/Default/Network/Cookies (movido a Network/)
    # Probamos todas las ubicaciones posibles.
    candidatos = [
        profile_dir / "Default" / "Network" / "Cookies",
        profile_dir / "Default" / "Cookies",
        profile_dir / "Network" / "Cookies",
        profile_dir / "Cookies",
    ]

    # Patrones de dominio (LIKE de SQL). Usamos % al inicio para capturar
    # tanto el dominio base como subdominios (e.g., `accounts.google.com`).
    patrones = [
        "%google.com",
        "%recaptcha.net",
        "%gstatic.com",
        "%googleapis.com",
        "%googleusercontent.com",
        "%ggpht.com",
    ]

    total_borradas = 0
    for db_path in candidatos:
        if not db_path.exists():
            continue
        try:
            # Timeout corto: si la base esta bloqueada (Chrome corriendo)
            # no queremos colgar el arranque por ella.
            conn = sqlite3.connect(str(db_path), timeout=1.5)
            try:
                cur = conn.cursor()
                # Verificar que existe la tabla cookies (Chrome la mantiene
                # estable pero por seguridad).
                cur.execute(
                    "SELECT name FROM sqlite_master "
                    "WHERE type='table' AND name='cookies'"
                )
                if not cur.fetchone():
                    continue
                borradas_en_db = 0
                for patron in patrones:
                    cur.execute(
                        "DELETE FROM cookies WHERE host_key LIKE ?",
                        (patron,),
                    )
                    borradas_en_db += cur.rowcount
                conn.commit()
                total_borradas += borradas_en_db
                if borradas_en_db > 0:
                    logger.info(
                        f"Cookies de Google/reCAPTCHA borradas en {db_path.name}: "
                        f"{borradas_en_db}"
                    )
            finally:
                conn.close()
        except sqlite3.OperationalError as err:
            # "database is locked" — Chrome esta abierto. No es fatal.
            logger.warning(
                f"No se pudieron borrar cookies en {db_path.name} "
                f"(probable Chrome abierto): {err}"
            )
            continue
        except Exception as err:
            logger.warning(
                f"Error inesperado limpiando cookies en {db_path.name}: "
                f"{type(err).__name__}: {err}"
            )
            continue
    return total_borradas


def _navegador_responde(context, timeout_ms: int = 8000) -> bool:
    """Confirma que el navegador está realmente vivo y responde.

    `launch_persistent_context` (y `browser.new_context`) pueden devolver un
    objeto aunque el proceso del navegador haya muerto al instante — el caso
    típico es la colisión de instancia única con el Chrome del sistema, que
    deja en consola un `<launched> pid=...` pero sin ventana usable.

    Navegar a about:blank fuerza una interacción real con el proceso: si está
    muerto, lanza excepción en vez de seguir con una `page` zombie.
    """
    try:
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("about:blank", timeout=timeout_ms)
        return True
    except Exception as err:
        logger.warning(
            f"El navegador se lanzó pero no responde: {type(err).__name__}: {err}"
        )
        return False


def _abrir_navegador(p):
    """Abre el navegador con la mejor estrategia disponible y fallbacks reales.

    Devuelve `(context, browser, using_persistent_profile)` SOLO cuando el
    navegador quedó verificado y funcional. Lanza `RuntimeError` con prefijo
    `[NAVEGADOR]` si no se puede abrir nada, para que la UI lo muestre como
    problema de navegador/perfil y no como un error genérico de descarga.

    Cada intento se valida con `_navegador_responde`: si el proceso se abrió
    pero murió, se descarta y se pasa al siguiente fallback (en vez de
    continuar con una `page` zombie).

    Estrategia (en orden):
      1. Perfil persistente fijo (`browser_profile`) con el Chromium incluido
         en Playwright. NO se usa `channel="chrome"`: el Chrome del sistema es
         de instancia única y, si el usuario lo tiene abierto, el proceso
         lanzado se cierra de inmediato.
      2. Perfil persistente temporal (el fijo puede estar bloqueado/corrupto).
      3. Contexto NO persistente (`browser.launch`), probando primero el Chrome
         del sistema y luego el Chromium incluido.
    """
    base_kwargs = dict(
        headless=HEADLESS,
        args=[
            "--no-sandbox", "--disable-setuid-sandbox",
            "--disable-dev-shm-usage", "--disable-gpu",
            # Quita el flag de "controlado por automatización" de Blink. Sin
            # esto, `navigator.webdriver=true` y Chrome muestra el infobar
            # "Un software de prueba automatizado está controlando Chrome",
            # que reCAPTCHA Enterprise usa como señal fuerte de bot.
            "--disable-blink-features=AutomationControlled",
            # Maximizar la ventana de Chrome para igualar el comportamiento
            # de AMU (que pasa el captcha al primer intento). Una ventana
            # tipica de 1280x854 con `screen.width=1280` huele a viewport
            # de Playwright/automatizacion — un usuario real usa monitor
            # 1920x1080 maximizado. Confirmado con DevTools snapshot del
            # 2026-06-18: AMU tiene window 1920x1032 (maximizado en monitor
            # FHD), nuestro bot tenia 1296x854 (default Chrome unmaxed) +
            # screen 1280x720 (default viewport Playwright). Combinado con
            # `no_viewport=True` abajo, screen.* reporta el monitor real.
            "--start-maximized",
        ],
        # Suprime el switch `--enable-automation` que Playwright añade por
        # defecto. Es el que pinta el infobar y setea webdriver=true en
        # Chrome ≥ 89.
        ignore_default_args=["--enable-automation"],
        # CRITICO para el captcha de Recibidos: NO emular un viewport fijo.
        # Por defecto Playwright fija viewport=1280x720 Y mente sobre
        # `screen.width/height` reportandolos como 1280x720 — algo que un
        # usuario real con monitor desktop NUNCA tiene. reCAPTCHA Enterprise
        # detecta esa inconsistencia (UA de Chrome desktop + pantalla 1280x720)
        # como señal fuerte de automatizacion → score bajo → captcha rechazado.
        # Con `no_viewport=True`, Chrome reporta el monitor fisico real y la
        # ventana ocupa lo que le diga `--start-maximized` arriba. AMU funciona
        # exactamente asi (verificado 2026-06-18).
        no_viewport=True,
    )
    if SLOW_MO > 0:
        base_kwargs["slow_mo"] = SLOW_MO
    if DEVTOOLS:
        base_kwargs["devtools"] = True
        base_kwargs["headless"] = False

    # Init script minimalista. Comparando con la app de referencia del
    # usuario (la que pasa el captcha sin problemas):
    #   - webdriver: false       ← solo viene del flag de Blink + no --enable-automation
    #   - plugins: 5 reales      ← NO sobreescritos, valores reales del Chrome 148
    #   - languages: ['es-419', 'es']  ← reales del SO Windows
    #   - chrome.runtime: NO existe (hasChromeRuntime: false)
    #   - cdc_* keys presentes (típico Selenium, reCAPTCHA los IGNORA)
    #
    # Por eso quitamos las sobreescrituras de plugins/languages/chrome.runtime:
    # solo lograban INCONSISTENCIA con el Chrome real y eran 100% innecesarias.
    # Dejamos solo un fallback defensivo de `navigator.webdriver`, que sólo
    # se aplica si por algún edge case Chrome no honró el flag de Blink.
    _STEALTH_INIT_SCRIPT = r"""
        try {
            // Solo cinturón y tirantes: el flag
            // `--disable-blink-features=AutomationControlled` ya pone webdriver
            // en `false`. Si por alguna razón Chrome lo dejó en `true`, lo
            // forzamos. Apuntamos a `false` (no `undefined`) para coincidir
            // con lo que reporta un Chrome real con el mismo flag.
            if (navigator.webdriver === true) {
                Object.defineProperty(navigator, 'webdriver', { get: () => false });
            }
        } catch (e) {}
    """

    def _aplicar_stealth(context):
        """Inyecta el init script en el contexto. No falla la apertura si por
        algún motivo el script no aplica."""
        try:
            context.add_init_script(_STEALTH_INIT_SCRIPT)
        except Exception as err:
            logger.warning(f"No se pudo aplicar stealth al contexto: {err}")

    errores = []

    def _validar(descripcion, context, browser, persistente):
        """Valida un contexto recién abierto. Devuelve la tupla o None.

        Si el navegador no responde, cierra lo que se haya abierto para no
        dejar procesos zombie y registra el motivo.
        """
        if _navegador_responde(context):
            logger.info(f"Navegador abierto y verificado: {descripcion}.")
            return context, browser, persistente
        errores.append(f"{descripcion}: el navegador se cerró tras abrirse")
        for cerrable in (context, browser):
            if cerrable is not None:
                try:
                    cerrable.close()
                except Exception:
                    pass
        return None

    # --- 1 y 2. Perfil persistente (perfil fijo, luego temporal) ---
    if USE_PERSISTENT_PROFILE:
        persistent_kwargs = dict(base_kwargs, accept_downloads=True)
        rutas = []

        # Si PLAYWRIGHT_USE_TEMP_PROFILE=1, saltamos el perfil persistente
        # fijo y usamos SOLO perfil temporal por sesión. Replica el
        # comportamiento de la app de referencia que crea `scoped_dir_XXXX`
        # nuevo cada vez y pasa el captcha al primer intento (no acumula
        # mala reputación de Google en cookies viejas).
        if not PLAYWRIGHT_USE_TEMP_PROFILE:
            perfil_fijo = Path(USER_DATA_DIR).expanduser()
            if not perfil_fijo.is_absolute():
                perfil_fijo = Path.cwd() / perfil_fijo
            try:
                perfil_fijo.mkdir(parents=True, exist_ok=True)
                _limpiar_locks_perfil(perfil_fijo)
                # Borrar cookies de Google/reCAPTCHA del perfil persistente
                # antes de lanzar Chrome. Resetea el "score base" que reCAPTCHA
                # Enterprise mantiene en esas cookies, evitando que el captcha
                # vuelva a fallar tras muchas sesiones en la misma maquina.
                # NO toca cookies del SRI ni de otros sitios.
                if RECIBIDOS_CLEAN_GOOGLE_COOKIES:
                    try:
                        n_borradas = _limpiar_cookies_google_perfil(perfil_fijo)
                        if n_borradas > 0:
                            logger.info(
                                f"Auto-limpieza captcha: borradas {n_borradas} "
                                f"cookies de Google/reCAPTCHA del perfil persistente "
                                f"(reset de score reCAPTCHA Enterprise sin afectar "
                                f"login SRI)."
                            )
                    except Exception as err:
                        logger.warning(
                            f"Falló la auto-limpieza de cookies de Google: "
                            f"{type(err).__name__}: {err}. Continuamos sin abortar."
                        )
                rutas.append(("perfil fijo", perfil_fijo))
            except Exception as err:
                errores.append(f"perfil fijo no accesible: {type(err).__name__}: {err}")
        else:
            logger.info(
                "PLAYWRIGHT_USE_TEMP_PROFILE=1: saltando perfil persistente "
                "fijo. Solo usaremos perfil temporal por sesión."
            )

        try:
            rutas.append(("perfil temporal", Path(tempfile.mkdtemp(prefix="sri_robot_profile_"))))
        except Exception as err:
            errores.append(f"perfil temporal no creable: {type(err).__name__}: {err}")

        # Probamos primero Chrome del sistema (si PREFER_SYSTEM_CHROME=1),
        # después Chromium bundled. La app de referencia que pasa el captcha
        # usa Chrome 148 del sistema; Chromium bundled de Playwright suele
        # estar varias versiones atrás.
        canales_orden: list[str | None] = []
        if PREFER_SYSTEM_CHROME:
            canales_orden.append("chrome")
        canales_orden.append(None)  # Chromium bundled (sin channel)

        for etiqueta, ruta in rutas:
            for canal in canales_orden:
                kw = dict(persistent_kwargs)
                if canal:
                    kw["channel"] = canal
                etiqueta_canal = f"{etiqueta} ({ruta}) canal={canal or 'chromium-bundled'}"
                try:
                    context = p.chromium.launch_persistent_context(str(ruta), **kw)
                except Exception as err:
                    errores.append(f"{etiqueta_canal}: {type(err).__name__}: {err}")
                    logger.warning(f"launch_persistent_context falló con {etiqueta_canal}: {err}")
                    continue
                _aplicar_stealth(context)
                resultado = _validar(etiqueta_canal, context, None, True)
                if resultado:
                    return resultado

    # --- 3. Contexto NO persistente ---
    for canal in ("chrome", "chromium"):
        kwargs = dict(base_kwargs)
        if canal == "chrome":
            kwargs["channel"] = "chrome"
        try:
            browser = p.chromium.launch(**kwargs)
            context = browser.new_context(accept_downloads=True)
        except Exception as err:
            errores.append(f"launch {canal}: {type(err).__name__}: {err}")
            continue
        _aplicar_stealth(context)
        resultado = _validar(f"contexto no persistente ({canal})", context, browser, False)
        if resultado:
            return resultado

    detalle = " | ".join(errores) or "sin detalle"
    raise RuntimeError(
        "[NAVEGADOR] No se pudo abrir el navegador. Verifica que Google Chrome "
        "esté instalado y cierra las ventanas de Chrome del robot que puedan "
        f"haber quedado abiertas. Detalle técnico: {detalle}"
    )


def _verificar_estado_post_login(page) -> None:
    """Valida que, tras el login, la página esté en un estado utilizable.

    Sin esta verificación, un navegador headless/zombie hace que el login
    "pase" silenciosamente y el fallo recién aparezca mucho después, al no
    encontrar los paneles del portal (Facturación Electrónica, Producción,
    Consultas) — un síntoma confuso para una causa de navegador.

    Lanza `RuntimeError` con un mensaje claro si:
      - el navegador/página se cerró (no se puede leer la URL),
      - seguimos en la pantalla de login (autenticación fallida),
      - el portal devolvió una página de indisponibilidad.
    """
    # 1. ¿La página sigue viva?
    try:
        url_actual = page.url or ""
    except Exception as err:
        raise RuntimeError(
            "[NAVEGADOR] El navegador se cerró durante el inicio de sesión. "
            f"Detalle: {type(err).__name__}: {err}"
        )

    url_low = url_actual.lower()

    # 2. ¿Seguimos en la pantalla de login de Keycloak?
    if "auth/realms" in url_low or "openid-connect" in url_low:
        raise RuntimeError(
            "La autenticacion en el SRI fallo: se mantuvo en la pantalla de "
            "login. Verifica el RUC y la clave e intenta nuevamente."
        )

    # 3. ¿El portal devolvió una página de indisponibilidad?
    try:
        portal_caido = _portal_indisponible(page)
    except Exception:
        portal_caido = False
    if portal_caido:
        raise RuntimeError(PORTAL_INDISPONIBLE_MENSAJE)

    logger.info(f"Sesion iniciada correctamente; pagina actual: {url_actual}")


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
        # Apertura del navegador con fallbacks reales (perfil persistente →
        # perfil temporal → contexto no persistente). Si nada funciona,
        # _abrir_navegador lanza RuntimeError con prefijo [NAVEGADOR].
        context, browser, using_persistent_profile = _abrir_navegador(p)

        page = context.pages[0] if context.pages else context.new_page()

        destino_url = PORTAL_HOME if origen in {"Recibidos", "Emitidos"} else URLS.get(origen, URLS["Recibidos"])
        _login(context, page, ruc, clave, cookies_path, destino_url, ci_adicional=ci_adicional)
        _check_cancel("post_login")
        # Valida explícitamente el estado antes de buscar paneles del portal:
        # navegador vivo, fuera de la pantalla de login, portal disponible.
        _verificar_estado_post_login(page)
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
                    # Fix Recibidos año-completo / rango de meses: entre meses
                    # cerramos el navegador y reabrimos con login fresco. Sin
                    # esto la sesión se degrada mes a mes (ej. Feb baja de 74 a
                    # 24) por captcha score, cookies y estado JSF acumulado.
                    # Solo aplica si aun quedan meses por procesar; la logica
                    # de descarga de un solo mes (_flujo_recibidos) no se toca.
                    if mes_actual < mes_fin_val:
                        _check_cancel("recibidos_reset_navegador")
                        print(
                            f"[INFO] Recibidos: cerrando navegador tras mes {mes_actual:02d} "
                            f"y reabriendo para {mes_actual + 1:02d}.",
                            flush=True,
                        )
                        try:
                            context.close()
                        except Exception as err_close_ctx:
                            print(
                                f"[WARN] Recibidos: error cerrando context: {err_close_ctx}",
                                flush=True,
                            )
                        try:
                            if browser is not None:
                                browser.close()
                        except Exception as err_close_br:
                            print(
                                f"[WARN] Recibidos: error cerrando browser: {err_close_br}",
                                flush=True,
                            )
                        context, browser, using_persistent_profile = _abrir_navegador(p)
                        page = context.pages[0] if context.pages else context.new_page()
                        _login(
                            context,
                            page,
                            ruc,
                            clave,
                            cookies_path,
                            destino_url,
                            ci_adicional=ci_adicional,
                        )
                        _check_cancel("recibidos_reset_post_login")
                        _verificar_estado_post_login(page)
                        modulo_page = _abrir_modulo_consultas(page, origen)
                        consultas_recibidos = 0
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
                # Para resume con granularidad pagina+fila: leer del progress
                # del checkpoint los campos current_page / current_row_index.
                # Solo aplican al PRIMER dia que entre al loop (idx_dia == 0).
                _resume_page_start = 1
                _resume_row_start = 0
                if (
                    resume_download
                    and isinstance(progress_data, dict)
                    and mes_actual == resume_month
                ):
                    try:
                        _resume_page_start = max(1, int(progress_data.get("current_page") or 1))
                        _resume_row_start = max(0, int(progress_data.get("current_row_index") or 0))
                    except (TypeError, ValueError):
                        _resume_page_start, _resume_row_start = 1, 0
                for idx_dia, dia_iter in enumerate(dias_consultar):
                    _check_cancel("emitidos_dia")
                    fecha_actual = f"{dia_iter:02d}/{mes_actual:02d}/{anio}"
                    # El resume fino (pagina/fila) SOLO aplica al primer dia
                    # del loop y SOLO si coincide con el dia donde quedo el
                    # checkpoint. Para los siguientes dias arrancamos en pag 1
                    # fila 0 normalmente.
                    _aplica_resume_fino = (
                        idx_dia == 0
                        and resume_download
                        and dia_iter == resume_day
                    )
                    _page_param = _resume_page_start if _aplica_resume_fino else 1
                    _row_param = _resume_row_start if _aplica_resume_fino else 0
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
                        checkpoint_path=checkpoint_path_str or None,
                        resume_page=_page_param,
                        resume_row_index=_row_param,
                        current_month=int(mes_actual),
                        current_day=int(dia_iter),
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
                            # Reset granularidad fina: el siguiente dia empieza
                            # desde pag 1 fila 0. Si NO reseteamos, el resume
                            # de un dia subsiguiente intentaria saltar a la
                            # pagina del dia previo, fallando.
                            current_page=1,
                            current_row_index=0,
                            total_rows_on_page=0,
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
