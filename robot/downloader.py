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

# _notificar_usuario_captcha y _notificar_usuario_accion movidos a robot/signals.py
# (Sub-fase 2b del refactor). Quedan re-importados en la cabecera para mantener
# compatibilidad con cualquier llamada interna.

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


def _es_url_autorizacion(url: str) -> bool:
    if not url:
        return False
    url_lower = url.lower()
    if "sri.gob.ec" not in url_lower:
        return False
    return "autoriz" in url_lower

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
            logger.warning(f"No aparecio el dialogo XML para {source_id}: {err}")
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
            logger.warning(f"No se encontro el formulario j_idt913 en el dialogo para {source_id}.")
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
            logger.warning(f"Fallo la solicitud POST de XML para {source_id}: {err}")
            raise
        if respuesta.status != 200:
            logger.warning(f"HTTP {respuesta.status} al solicitar XML de emitidos para {source_id}.")
            raise RuntimeError(f"Error HTTP {respuesta.status} al descargar XML de emitidos.")
        try:
            cuerpo_bytes = respuesta.body()
        except Exception as err:
            logger.warning(f"No se pudo leer cuerpo de respuesta XML para {source_id}: {err}")
            raise RuntimeError("No se pudo leer la respuesta del XML de emitidos.") from err
        if not cuerpo_bytes:
            logger.warning(f"Respuesta vacia al descargar XML para {source_id}.")
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
            logger.warning(f"No se pudo interpretar el XML de emitidos descargado: {err}")
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
        logger.warning(f"{mensaje}")
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
        logger.warning(f"{mensaje}")
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
            logger.warning(f"No se pudo interactuar con {descripcion}: {err}")
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
                logger.warning(f"Reintentando acceso directo al formulario ({intento + 1}/3): {err}")
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
            logger.warning(f"No se encontró {descripcion}.")
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
            logger.warning(f"No se pudo expandir {descripcion}: {err}")
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
            logger.warning(f"No se pudo hacer clic en 'Consultas': {err}")
    else:
        logger.warning("No se encontró el botón de Consultas; intentando acceso directo.")

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
