"""Flujos de descarga del SRI: orquestacion de Recibidos y Emitidos.

`_flujo_recibidos` y `_flujo_emitidos` son el corazon del bot: recorren la
tabla de comprobantes del portal, descargan cada PDF/XML, extraen sus datos
y generan los reportes Excel. Reciben la `page` de Playwright ya autenticada.

`descargar_sri` (en robot/downloader.py) crea el navegador, hace login y
delega en estos flujos.

Extraido de `robot/downloader.py` en la Fase 4 del refactor.
"""
from __future__ import annotations

import os
import random
import re
import time
import unicodedata
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from robot._logging import get_logger
from robot.parser import construir_reporte
from robot.browser import (
    _click_consultar_emitidos,
    _descargar_pdf_emitidos_post_con_viewstate,
    _descargar_pdf_recibidos_post_con_viewstate,
    _descargar_xml_emitido_por_clave,
    _esperar_ajax,
    _extraer_datos_emitidos_dom,
    _guardar_pdf_desde_enlace,
    _guardar_pdf_desde_jsf,
    _guardar_xml_desde_enlace,
    _obtener_detalle_emitido_xhr,
    _obtener_form_base_emitidos,
    _obtener_source_detalle_emitido,
    _obtener_view_state,
    _rellenar_input_por_label,
    _resolver_destino_unico,
    _seleccionar,
    _seleccionar_en_select,
    _seleccionar_por_label,
)
from robot.comprobante_types import (
    _coincide_tipo_documental,
    _es_tipo_factura,
    _es_tipo_liquidacion_compra,
    _es_tipo_nota_credito,
    _es_tipo_nota_debito,
    _es_tipo_retencion,
    _nombre_carpeta_tipo_visible,
    _prefijo_tipo,
    _slug_tipo,
)
from robot.config import (
    DOWNLOAD_TIMEOUT,
    ESTADOS_EMITIDOS_MAP,
    MANUAL_CONSULTA_RECIBIDOS,
    PAUSE_BEFORE_CONSULTAR,
    PAUSE_BEFORE_CONSULTAR_SECONDS,
    RECIBIDOS_AUTO_POST_EXECUTE_MS,
    RECIBIDOS_AUTO_PRE_EXECUTE_MS,
    RECIBIDOS_AUTO_RESULT_TIMEOUT_MS,
    RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC,
    RECIBIDOS_CONSULTA_INTENTOS,
    RECIBIDOS_DIRECT_URL,
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MAX,
    RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN,
    RECIBIDOS_HUMANIZAR_CLICK_X_RATIO,
    RECIBIDOS_HUMANIZAR_CLICK_Y_RATIO,
    RECIBIDOS_HUMANIZAR_CLICKS,
    RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS,
    RECIBIDOS_HUMANIZAR_PRE_CLICK,
    RECIBIDOS_RECAPTCHA_TOKEN_TIMEOUT_MS,
    RECIBIDOS_REHIDRATAR_DESDE_INTENTO,
    RECIBIDOS_REHIDRATAR_ON_CAPTCHA,
    TIPOS_MAP,
)
from robot.data_formatters import _parse_datetime_local
from robot.file_utils import _mes_a_texto, _sanear_nombre_archivo
from robot.pdf_extraction import (
    _combinar_datos_reporte_emitidos,
    _es_archivo_pdf,
    _extraer_datos_pdf_factura_emitido,
    _extraer_datos_pdf_liquidacion_compra_emitido,
    _extraer_datos_pdf_nota_credito_emitido,
    _extraer_datos_pdf_nota_debito_emitido,
    _extraer_datos_pdf_por_tipo_layout_first,
    _extraer_datos_pdf_retencion_emitido,
    _extraer_datos_xml_liquidacion_compra_emitido,
    _extraer_datos_xml_pdf_report,
    _extraer_datos_xml_retencion,
    _extraer_datos_xml_retencion_emitido,
    _extraer_tipo_documento,
)
from robot.reporting import (
    _guardar_reporte_pdf_excel,
    _guardar_reporte_pdf_factura_emitidos_excel,
    _guardar_reporte_pdf_nota_credito_emitidos_excel,
    _guardar_reporte_pdf_nota_debito_emitidos_excel,
    _guardar_reporte_pdf_retencion_emitidos_excel,
    _guardar_reporte_pdf_retencion_excel,
)
from robot.signals import _check_cancel, _notificar_usuario_accion
from robot.xml_extraction import (
    _extraer_datos_xml_factura_emitido,
    _extraer_datos_xml_nota_credito_emitido,
    _extraer_datos_xml_nota_debito_emitido,
)


logger = get_logger(__name__)


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

DOWNLOAD_ROW_RETRY_ATTEMPTS = max(1, int(os.getenv("DOWNLOAD_ROW_RETRY_ATTEMPTS", "2")))


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


def _nombre_documento_mes(tipo_slug: str, fecha_token: str, nombre_base: str) -> str:
    tipo_parte = _slug_tipo(tipo_slug or "") or "documento"
    fecha_parte = re.sub(r"[^0-9]", "", fecha_token or "") or "00000000"
    base = _sanear_nombre_archivo(nombre_base or "archivo")
    return f"{tipo_parte}__{fecha_parte}__{base}"


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

    # --------------------------------------------------------------------- #
    # Tracker de responses POST a comprobantesRecibidos.jsf
    # --------------------------------------------------------------------- #
    # Por cada clic en "Consultar", el SRI dispara 2 requests AJAX al mismo
    # endpoint (confirmado en DevTools del portal):
    #   #1: source=frmPrincipal:btnBuscar, g-recaptcha-response VACÍO,
    #       disparado por el `PrimeFaces.ab(...)` del onclick — devuelve la
    #       alerta "No se pudo validar el captcha en google" (es normal,
    #       lo ignoramos).
    #   #2: source=frmPrincipal:j_idtNN, g-recaptcha-response con TOKEN,
    #       disparado por `rcBuscar()` tras el callback de Google — trae
    #       el render de `tablaCompRecibidos`. ESTE es el resultado real.
    #
    # Capturamos cada response para distinguir el #1 del #2 y decidir si
    # esperamos más o si ya tenemos un resultado válido. Guardamos los
    # handles y leemos el cuerpo solo cuando hace falta (es caro).
    #
    # Cacheamos el listener en la página para evitar duplicarlo cuando se
    # llama varias veces dentro de la misma sesión (otro mes, otro tipo).
    listener_attr = "_recibidos_response_listener"
    eventos_attr = "_recibidos_response_eventos"
    if not getattr(page, listener_attr, None):
        eventos_pagina: list[dict] = []

        def _on_response_recibidos(response, _eventos=eventos_pagina) -> None:
            try:
                if "comprobantesRecibidos.jsf" not in (response.url or ""):
                    return
                req = response.request
                if req.method != "POST":
                    return
                _eventos.append({
                    "response": response,
                    "request": req,
                    "ts": time.time(),
                })
            except Exception:
                pass

        page.on("response", _on_response_recibidos)
        setattr(page, listener_attr, _on_response_recibidos)
        setattr(page, eventos_attr, eventos_pagina)

    recibidos_eventos: list[dict] = getattr(page, eventos_attr)

    def _resumir_eventos_recibidos() -> list[dict]:
        """Inspecciona cada response capturada para extraer (source, has_token,
        tiene_render_tabla, tiene_alerta_captcha). Solo se llama bajo demanda
        para no parsear cuerpos innecesariamente.
        """
        from urllib.parse import unquote
        resumen: list[dict] = []
        for ev in recibidos_eventos:
            post = ""
            try:
                post = ev["request"].post_data or ""
            except Exception:
                post = ""
            source = ""
            has_token = False
            for kv in post.split("&"):
                if kv.startswith("javax.faces.source="):
                    try:
                        source = unquote(kv.split("=", 1)[1])
                    except Exception:
                        source = kv.split("=", 1)[1]
                elif kv.startswith("g-recaptcha-response="):
                    try:
                        valor = kv.split("=", 1)[1]
                    except Exception:
                        valor = ""
                    has_token = len(valor) > 50
            tiene_render = False
            tiene_alerta_captcha = False
            status = 0
            try:
                status = int(ev["response"].status or 0)
            except Exception:
                pass
            try:
                # Solo leemos el cuerpo si la response ya terminó (status>0).
                if status:
                    text = ev["response"].text() or ""
                    tiene_render = "tablaCompRecibidos" in text
                    text_low = text.lower()
                    tiene_alerta_captcha = (
                        "validar el captcha" in text_low
                        or "captcha en google" in text_low
                        or "no se pudo validar" in text_low
                    )
            except Exception:
                pass
            resumen.append({
                "source": source or "?",
                "has_token": has_token,
                "tiene_render_tabla": tiene_render,
                "tiene_alerta_captcha": tiene_alerta_captcha,
                "status": status,
            })
        return resumen

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
        """True si el portal rechazó el captcha (cualquier variante conocida)."""
        if not texto:
            return False
        texto_norm = texto.strip().lower()
        # Variantes observadas:
        #   "Captcha incorrecto"
        #   "Captcha inválido"
        #   "No se pudo validar el captcha en google" (reCAPTCHA rechazado)
        #   "Captcha de google inválido"
        return (
            "captcha incorrect" in texto_norm
            or "captcha inval" in texto_norm
            or "no se pudo validar el captcha" in texto_norm
            or "captcha en google" in texto_norm
            or "captcha de google" in texto_norm
            or "recaptcha" in texto_norm
        )

    def _diagnostico_recaptcha_apis() -> dict:
        """Devuelve qué APIs de reCAPTCHA están disponibles en la página.

        Útil para logs: nos dice si el portal cargó executeRecaptcha,
        grecaptcha.enterprise.execute o rcBuscar antes de cada intento.
        """
        try:
            return page.evaluate(
                """() => ({
                    hasExecuteRecaptcha: typeof window.executeRecaptcha === 'function',
                    hasEnterpriseExecute: !!(window.grecaptcha
                        && grecaptcha.enterprise
                        && typeof grecaptcha.enterprise.execute === 'function'),
                    hasGrecaptchaExecute: !!(window.grecaptcha
                        && typeof grecaptcha.execute === 'function'),
                    hasRcBuscar: typeof window.rcBuscar === 'function',
                })"""
            ) or {}
        except Exception:
            return {}

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

    # ----- Constantes del flujo nativo del portal (descubiertas en DevTools) ----- #
    # El botón "Consultar" tiene onclick:
    #   deshabilitarBoton(this);
    #   executeRecaptcha('consulta_cel_recibidos','SI');
    #   PrimeFaces.ab({source:'frmPrincipal:btnBuscar'});
    #   return false;
    # Y el reset del reCAPTCHA del SRI se hace con resetarRecaptcha() (función
    # global del portal), NO con grecaptcha.enterprise.reset() (esa última solo
    # existe para reCAPTCHA v2 checkbox y no aplica al invisible/Enterprise).
    RECAPTCHA_ACTION = "consulta_cel_recibidos"
    RECAPTCHA_FLAG = "SI"

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
            # Usamos resetarRecaptcha() del SRI (función nativa del portal) en
            # vez de grecaptcha.enterprise.reset() porque esta última no existe
            # para reCAPTCHA invisible/Enterprise.
            page.evaluate(
                """() => {
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
                    try {
                        if (typeof window.resetarRecaptcha === 'function') {
                            window.resetarRecaptcha();
                        }
                    } catch (e) {}
                }"""
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

    def _esperar_api_recaptcha_lista(timeout: int = 8000) -> bool:
        """Espera a que `executeRecaptcha` del SRI (wrapper de Google) esté lista."""
        try:
            page.wait_for_function(
                """() => typeof window.executeRecaptcha === 'function'""",
                timeout=timeout,
            )
            return True
        except Exception:
            return False

    def _humanizar_pre_consultar() -> dict:
        """Genera interacción humana antes de hacer click en "Consultar".

        Confirmación empírica del usuario: si hace ~6 clicks rápidos en una
        zona blanca de la página antes de presionar Consultar, el SRI acepta
        el captcha y devuelve la tabla. Sin esos clicks, Google asigna score
        bajo y el SRI rechaza con "Captcha incorrecta".

        Estrategia (replicando la prueba manual del usuario):
          1) Pausa inicial (~1.8s default) para "tiempo en página".
          2) Mouse moves en 3 puntos con `steps` intermedios (trayectoria).
          3) Verificar que las coordenadas (X_RATIO, Y_RATIO) del viewport
             caen sobre un elemento NO interactivo (no input/select/button/
             /anchor/label/role=button, sin onclick). Si no es seguro, log
             y se saltan los clicks (queda el resto de la humanización).
          4) `RECIBIDOS_HUMANIZAR_CLICKS` clicks reales con jitter ±15px y
             `delay` aleatorio entre [DELAY_MIN, DELAY_MAX]. Pausa 80-180ms
             entre clicks.
          5) Pausa 300-700ms post-clicks.
          6) Hover final sobre el botón Consultar (mouseover + mouseenter
             generan eventos que reCAPTCHA reconoce como humano antes del
             click real).

        Devuelve {aplicado, puntos_mouse, clicks_blancos, coordenadas,
        zona_segura, duracion_ms} para logging.
        """
        out = {
            "aplicado": False,
            "puntos_mouse": 0,
            "clicks_blancos": 0,
            "coordenadas": (0, 0),
            "zona_segura": False,
            "duracion_ms": 0,
        }
        if not RECIBIDOS_HUMANIZAR_PRE_CLICK:
            return out
        inicio = time.perf_counter()
        try:
            # (1) Pausa inicial: "tiempo en la página" antes de submitear.
            if RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS > 0:
                try:
                    page.wait_for_timeout(RECIBIDOS_HUMANIZAR_PAUSA_INICIAL_MS)
                except Exception:
                    pass

            # (2) Trayectoria de mouse en zonas seguras (lejos de controles).
            vp = page.viewport_size or {"width": 1280, "height": 720}
            ancho = max(800, int(vp.get("width") or 1280))
            alto = max(600, int(vp.get("height") or 720))
            puntos = [
                (max(80, ancho // 6), max(70, alto // 7)),
                (ancho // 2, max(90, alto // 5)),
                (min(ancho - 80, ancho - ancho // 5), max(80, alto // 6)),
            ]
            for x, y in puntos:
                try:
                    page.mouse.move(int(x), int(y), steps=random.randint(8, 14))
                except Exception:
                    pass
                try:
                    page.wait_for_timeout(random.randint(180, 380))
                except Exception:
                    pass
            out["puntos_mouse"] = len(puntos)

            # (3) Calcular coordenada para los clicks blancos.
            click_x = int(ancho * RECIBIDOS_HUMANIZAR_CLICK_X_RATIO)
            click_y = int(alto * RECIBIDOS_HUMANIZAR_CLICK_Y_RATIO)
            out["coordenadas"] = (click_x, click_y)

            # Verificación de seguridad: que el elemento en (click_x, click_y)
            # NO sea interactivo. Logueamos también el outerHTML truncado y el
            # rect del elemento, para debug si el portal cambia layout.
            zona_segura = False
            elemento_diag = ""
            elemento_html = ""
            if RECIBIDOS_HUMANIZAR_CLICKS > 0:
                try:
                    diag = page.evaluate(
                        """({x, y}) => {
                            const el = document.elementFromPoint(x, y);
                            if (!el) return { segura: true, tag: 'NONE', motivo: 'fuera-de-pagina', html: '' };
                            const tag = (el.tagName || '').toUpperCase();
                            const role = el.getAttribute ? (el.getAttribute('role') || '') : '';
                            const id = el.id ? '#' + el.id : '';
                            const cls = el.className && typeof el.className === 'string'
                                ? '.' + el.className.trim().split(/\\s+/).slice(0,3).join('.') : '';
                            const html = (el.outerHTML || '').slice(0, 120);
                            const desc = tag + id + cls;
                            const noSegurosTag = ['INPUT','SELECT','TEXTAREA','BUTTON','A','LABEL','OPTION','IFRAME'];
                            if (noSegurosTag.includes(tag)) {
                                return { segura: false, tag: desc, motivo: 'tag-interactivo', html };
                            }
                            if (role === 'button' || role === 'link' || role === 'tab') {
                                return { segura: false, tag: desc, motivo: 'role-' + role, html };
                            }
                            if (typeof el.onclick === 'function' && el.onclick !== null) {
                                return { segura: false, tag: desc, motivo: 'tiene-onclick', html };
                            }
                            // Subimos por ancestros buscando button/a clickeable
                            let cur = el;
                            for (let i = 0; cur && i < 4; i++) {
                                const t = (cur.tagName || '').toUpperCase();
                                if (['BUTTON','A','INPUT','SELECT'].includes(t)) {
                                    return { segura: false, tag: desc, motivo: 'ancestor-' + t, html };
                                }
                                cur = cur.parentElement;
                            }
                            return { segura: true, tag: desc, motivo: 'ok', html };
                        }""",
                        arg={"x": click_x, "y": click_y},
                    ) or {}
                    zona_segura = bool(diag.get("segura"))
                    elemento_diag = f"{diag.get('tag','?')}/{diag.get('motivo','?')}"
                    elemento_html = str(diag.get("html") or "")
                except Exception:
                    zona_segura = False
                    elemento_diag = "eval-error"
            out["zona_segura"] = zona_segura

            # (4) Clicks reales en la zona blanca (si es segura).
            # Cambios para parecer más humano (test manual del usuario lo confirma):
            #   - Antes de cada click hacemos `mouse.move(x, y, steps=N)` para
            #     que el cursor viaje (no teleporte) entre puntos. Esto genera
            #     N eventos `mousemove` por trayecto.
            #   - Pausa pre-click (60-140ms) y post-click (220-420ms) más
            #     largas que antes (80-180ms total).
            #   - Jitter de ±30 px (era ±15) para que los clicks no caigan en
            #     la misma columna de pixels.
            if zona_segura and RECIBIDOS_HUMANIZAR_CLICKS > 0:
                logger.info(
                    "Recibidos: humanizar -> intentando %d clicks blancos en "
                    "(%d,%d) sobre %s | html=%r",
                    RECIBIDOS_HUMANIZAR_CLICKS, click_x, click_y, elemento_diag,
                    elemento_html,
                )
                for _ in range(RECIBIDOS_HUMANIZAR_CLICKS):
                    jx = click_x + random.randint(-30, 30)
                    jy = click_y + random.randint(-30, 30)
                    # Viaje gradual hacia la nueva coordenada (genera mousemoves
                    # intermedios, no teleporte).
                    try:
                        page.mouse.move(jx, jy, steps=random.randint(5, 12))
                    except Exception:
                        pass
                    # Pausa breve antes del click (humano se "asienta" sobre el
                    # punto antes de presionar).
                    try:
                        page.wait_for_timeout(random.randint(60, 140))
                    except Exception:
                        pass
                    delay = random.randint(
                        RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MIN,
                        RECIBIDOS_HUMANIZAR_CLICK_DELAY_MS_MAX,
                    )
                    try:
                        page.mouse.click(jx, jy, delay=delay)
                        out["clicks_blancos"] += 1
                    except Exception:
                        pass
                    # Pausa más larga entre clicks (era 80-180ms; subido a
                    # 220-420 porque la prueba manual del usuario tiene ritmo
                    # más lento).
                    try:
                        page.wait_for_timeout(random.randint(220, 420))
                    except Exception:
                        pass
                logger.info(
                    "Recibidos: humanizar -> %d clicks blancos completados.",
                    out["clicks_blancos"],
                )
            elif RECIBIDOS_HUMANIZAR_CLICKS > 0:
                logger.warning(
                    "Recibidos: coords humanizar (%d,%d) caen sobre control "
                    "interactivo (%s) | html=%r. Saltando clicks blancos. "
                    "Ajustá RECIBIDOS_HUMANIZAR_CLICK_X_RATIO/Y_RATIO si querés "
                    "forzar los clicks en otra zona.",
                    click_x, click_y, elemento_diag, elemento_html,
                )

            # (5) Pausa post-clicks (era 300-700ms; subida a 700-1500 para
            # que el modelo de reCAPTCHA vea más "thinking time" antes del
            # click definitivo).
            try:
                page.wait_for_timeout(random.randint(700, 1500))
            except Exception:
                pass

            # (6) Hover final sobre el botón "Consultar".
            try:
                box = boton_consultar.first.bounding_box()
                if box:
                    cx = int(box["x"] + box["width"] / 2)
                    cy = int(box["y"] + box["height"] / 2)
                    page.mouse.move(cx, cy, steps=random.randint(8, 12))
                    page.wait_for_timeout(random.randint(180, 320))
            except Exception:
                pass

            out["aplicado"] = True
        except Exception as err:
            logger.warning(f"Recibidos: humanizar pre-click falló: {err}")
        out["duracion_ms"] = int((time.perf_counter() - inicio) * 1000)
        return out

    def _disparar_consulta_recibidos_automatica(action: str, flag: str) -> dict:
        """Dispara la consulta haciendo click NATURAL sobre el botón.

        Confirmado por inspección en DevTools del portal del SRI:

          function executeRecaptcha(accion, validaRecaptcha) {
              if (validaRecaptcha === "SI") {
                  grecaptcha.enterprise.ready(function() {
                      grecaptcha.enterprise.execute({ action: accion });
                  });
              } else { rcBuscar(); }
          }

        `executeRecaptcha` NO devuelve nada (no Promise). Es fire-and-forget:
        encola el execute para cuando reCAPTCHA esté listo. El token llega
        ASÍNCRONAMENTE; Google lo entrega al callback que internamente llama
        `rcBuscar()`, y `rcBuscar()` dispara un segundo `PrimeFaces.ab` que
        ESE SÍ va con el token correcto. Antes hacíamos un PrimeFaces.ab
        adicional manual — eso solo agregaba ruido.

        Por lo tanto: click natural. El onclick original genera 2 requests:
          #1: source=frmPrincipal:btnBuscar     g-recaptcha-response=""
              (alerta "No se pudo validar el captcha" — normal, lo ignoramos)
          #2: source=frmPrincipal:j_idtNN       g-recaptcha-response=<token>
              (render de tablaCompRecibidos — éxito)

        Antes del click hacemos pre-interacción humana (mouse moves) para
        que reCAPTCHA Enterprise suba el score. Sin esto, Google rechaza
        con "Captcha incorrecta" (score ≈ 0.0 por bot detection).

        Devuelve {modo, click_ok, humanizar} para logging. `modo` es siempre
        'click-nativo' (o '' si el click falló).
        """
        # `action` y `flag` se aceptan solo para logging consistente — no
        # los reusamos para forzar nada, el onclick original ya los conoce.
        del action, flag  # solo para los logs del caller
        humanizar_info = _humanizar_pre_consultar()
        try:
            boton_consultar.first.scroll_into_view_if_needed()
        except Exception:
            pass
        try:
            # `delay=` añade tiempo entre mousedown/mouseup. Lo subimos a
            # 150-300ms (vs 100-160 previo) para parecer más humano. Un
            # mousedown/mouseup tan rápido como 100ms se ve sintético.
            boton_consultar.first.click(delay=random.randint(150, 300))
            return {"modo": "click-nativo", "click_ok": True, "humanizar": humanizar_info}
        except Exception as err:
            logger.warning(f"Recibidos: no se pudo hacer click nativo en Consultar: {err}")
            return {"modo": "", "click_ok": False, "humanizar": humanizar_info}

    def _esperar_resultado_consulta(timeout: int = 300000) -> tuple[str, str]:
        """Espera el resultado válido de la consulta de Recibidos.

        Respeta el flujo de 2 requests del SRI:
          - Si llega request #2 con render de `tablaCompRecibidos` → ÉXITO.
          - Si llega alerta NO captcha (e.g. "sin resultados") → resultado final.
          - La alerta "No se pudo validar el captcha" (del request #1) se
            IGNORA mientras todavía podamos recibir el #2.

        Devuelve un tuple (estado, mensaje):
          ('tabla',    '')            → tabla rendereada
          ('alerta',   '<texto>')     → alerta NO captcha (sin resultados, etc.)
          ('captcha',  '<texto>')     → captcha rechazado de verdad
                                        (no llegó render del #2 o el #2 trajo
                                        la alerta de captcha)
          ('timeout',  '<texto>')     → ni tabla ni alerta no-captcha
        """
        limite = time.time() + (timeout / 1000)
        while time.time() < limite:
            # (a) Tabla visible en el DOM → éxito directo
            try:
                if tabla_datos.is_visible():
                    return ("tabla", "")
            except Exception:
                pass

            eventos = _resumir_eventos_recibidos()

            # (b) Algún response trajo render de tablaCompRecibidos → éxito
            #     (incluso si la tabla todavía no se pintó, llega en ms)
            if any(e.get("tiene_render_tabla") for e in eventos):
                try:
                    page.wait_for_timeout(250)
                except Exception:
                    pass
                try:
                    if tabla_datos.is_visible():
                        return ("tabla", "")
                except Exception:
                    pass
                # El #2 trajo render pero el DOM todavía no actualizó →
                # seguimos esperando un poco más.

            # (c) Si llegó el response #2 (has_token=True) y trajo alerta de
            #     captcha (no render), es captcha rechazado de verdad.
            eventos_con_token = [e for e in eventos if e.get("has_token")]
            if eventos_con_token:
                ultimo_token = eventos_con_token[-1]
                if (
                    ultimo_token.get("tiene_alerta_captcha")
                    and not ultimo_token.get("tiene_render_tabla")
                ):
                    return ("captcha", _texto_alerta() or "Captcha rechazado por SRI/Google.")

            # (d) Alerta NO captcha visible → resultado final no-error
            texto = _texto_alerta()
            if texto and not _es_alerta_captcha(texto):
                return ("alerta", texto)

            time.sleep(0.2)

        # Timeout: lo que haya en pantalla
        texto_final = _texto_alerta()
        eventos_finales = _resumir_eventos_recibidos()
        if any(e.get("tiene_render_tabla") for e in eventos_finales):
            try:
                if tabla_datos.is_visible():
                    return ("tabla", "")
            except Exception:
                pass
        if texto_final and _es_alerta_captcha(texto_final):
            return ("captcha", texto_final)
        if texto_final:
            return ("alerta", texto_final)
        return ("timeout", "")

    def _rehidratar_consulta_recibidos() -> bool:
        """Recarga la pantalla de Recibidos para regenerar el entorno reCAPTCHA."""
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
        # NOTA: no tocamos el onclick original del botón. El flujo nativo del
        # SRI corre cuando hagamos click natural, o nosotros replicamos
        # paso a paso en _disparar_consulta_recibidos_automatica.
        return ok

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
        def _resumen_eventos_log(eventos: list[dict]) -> str:
            """String corto con un resumen de los 2 requests para logs."""
            partes = []
            for i, ev in enumerate(eventos, 1):
                partes.append(
                    f"#{i}(source={ev.get('source','?')} "
                    f"has_token={ev.get('has_token', False)} "
                    f"render_tabla={ev.get('tiene_render_tabla', False)} "
                    f"alerta_captcha={ev.get('tiene_alerta_captcha', False)} "
                    f"status={ev.get('status', 0)})"
                )
            return " | ".join(partes) if partes else "sin-requests"

        if MANUAL_CONSULTA_RECIBIDOS:
            for intento in range(1, intentos + 1):
                recibidos_eventos.clear()
                _notificar_usuario_accion(
                    f"[ACCION] Da clic manual en 'Consultar' (Recibidos). "
                    f"Intento {intento}/{intentos}."
                )
                estado, mensaje = _esperar_resultado_consulta(timeout=300000)
                eventos_log = _resumen_eventos_log(_resumir_eventos_recibidos())
                logger.info(
                    "Recibidos MANUAL intento %d/%d | estado=%s | requests=%s | mensaje=%r",
                    intento, intentos, estado, eventos_log, mensaje,
                )
                if estado == "tabla":
                    return True
                if estado == "alerta":
                    return True
                if estado == "captcha":
                    print(
                        f"[WARN] Captcha incorrecto tras clic manual "
                        f"({intento}/{intentos}): {mensaje!r}"
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
                    continue
                # timeout sin resultado
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

        # ===================== Modo automático ===================== #
        ultima_alerta_captcha = ""
        for intento in range(1, intentos + 1):
            inicio_intento = time.perf_counter()
            # Limpiamos eventos del tracker para este intento.
            recibidos_eventos.clear()
            # Limpieza de estado SIEMPRE (no solo en reintentos): así el
            # token anterior nunca se reutiliza, y resetarRecaptcha() del SRI
            # le pide a Google un score nuevo en el próximo execute.
            _limpiar_estado_consulta()
            rehidratado = False
            if RECIBIDOS_REHIDRATAR_ON_CAPTCHA and intento >= RECIBIDOS_REHIDRATAR_DESDE_INTENTO:
                rehidratado = _rehidratar_consulta_recibidos()
                if rehidratado:
                    recibidos_eventos.clear()  # los responses post-recarga ya no aplican
            api_lista = _esperar_api_recaptcha_lista(timeout=7000)
            apis_diag = _diagnostico_recaptcha_apis()
            if RECIBIDOS_AUTO_PRE_EXECUTE_MS > 0:
                try:
                    page.wait_for_timeout(RECIBIDOS_AUTO_PRE_EXECUTE_MS)
                except Exception:
                    pass
            # Click natural sobre el botón. El onclick original del SRI hace
            # executeRecaptcha → Google → rcBuscar → PrimeFaces.ab con token.
            # Genera 2 requests; nosotros sólo escuchamos.
            disparo = _disparar_consulta_recibidos_automatica(
                RECAPTCHA_ACTION, RECAPTCHA_FLAG
            )
            modo_disparo = disparo.get("modo", "")
            click_ok = bool(disparo.get("click_ok"))
            humanizar_info = disparo.get("humanizar") or {}
            humanizar_aplicado = bool(humanizar_info.get("aplicado"))
            humanizar_dur = int(humanizar_info.get("duracion_ms") or 0)
            humanizar_puntos = int(humanizar_info.get("puntos_mouse") or 0)
            humanizar_clicks = int(humanizar_info.get("clicks_blancos") or 0)
            humanizar_coords = humanizar_info.get("coordenadas") or (0, 0)
            humanizar_zona_segura = bool(humanizar_info.get("zona_segura"))
            if RECIBIDOS_AUTO_POST_EXECUTE_MS > 0:
                try:
                    page.wait_for_timeout(RECIBIDOS_AUTO_POST_EXECUTE_MS)
                except Exception:
                    pass

            estado, mensaje = _esperar_resultado_consulta(
                timeout=RECIBIDOS_AUTO_RESULT_TIMEOUT_MS
            )

            eventos = _resumir_eventos_recibidos()
            eventos_log = _resumen_eventos_log(eventos)
            n_requests = len(eventos)
            algun_token = any(e.get("has_token") for e in eventos)
            algun_render = any(e.get("tiene_render_tabla") for e in eventos)

            dur = time.perf_counter() - inicio_intento
            logger.info(
                "Recibidos intento %d/%d | action=%s flag=%s | modo=%s click_ok=%s | "
                "humanizar=%s (puntos=%d clicks=%d zona_segura=%s coords=%s dur=%dms) | "
                "api_lista=%s executeRecaptcha=%s enterprise.execute=%s rcBuscar=%s | "
                "rehidratado=%s | n_requests=%d algun_token=%s render_tabla=%s | "
                "estado=%s mensaje=%r | requests=%s | dur=%.2fs",
                intento, intentos,
                RECAPTCHA_ACTION, RECAPTCHA_FLAG,
                modo_disparo or "desconocido", click_ok,
                humanizar_aplicado, humanizar_puntos, humanizar_clicks,
                humanizar_zona_segura, humanizar_coords, humanizar_dur,
                api_lista,
                apis_diag.get("hasExecuteRecaptcha", False),
                apis_diag.get("hasEnterpriseExecute", False),
                apis_diag.get("hasRcBuscar", False),
                rehidratado,
                n_requests, algun_token, algun_render,
                estado, mensaje,
                eventos_log,
                dur,
            )
            print(
                f"[INFO] Recibidos intento {intento}/{intentos}: "
                f"action={RECAPTCHA_ACTION!r}, modo={modo_disparo or 'desconocido'}, "
                f"humanizar={humanizar_aplicado} ({humanizar_puntos}m/"
                f"{humanizar_clicks}c@{humanizar_coords}/{humanizar_dur}ms), "
                f"n_requests={n_requests}, algun_token={algun_token}, "
                f"render_tabla={algun_render}, estado={estado}, "
                f"rehidratado={rehidratado}, dur={dur:.2f}s"
            )

            if estado == "tabla":
                return True
            if estado == "alerta":
                # Alerta no-captcha (e.g. sin resultados): la capa superior la procesa.
                return True
            if estado == "captcha":
                ultima_alerta_captcha = mensaje or "Captcha rechazado por SRI/Google."
                logger.warning(
                    "Captcha rechazado de verdad en intento %d/%d "
                    "(n_requests=%d, algun_token=%s, render_tabla=%s).",
                    intento, intentos, n_requests, algun_token, algun_render,
                )
                _limpiar_estado_consulta()
                if intento < intentos and RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                    espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                    logger.info(f"Esperando {espera:.1f}s antes de reintentar Recibidos.")
                    time.sleep(espera)
                continue
            # estado == "timeout"
            if intento < intentos and RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC > 0:
                espera = RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC * intento
                logger.info(
                    f"Sin tabla ni alerta tras intento {intento}/{intentos}. "
                    f"Espera {espera:.1f}s."
                )
                time.sleep(espera)
        # Agotamos todos los intentos. Si el motivo fue captcha rechazado,
        # damos un mensaje claro al usuario.
        if ultima_alerta_captcha:
            mensaje_final = (
                f"El SRI/Google rechazó el captcha tras {intentos} intentos. "
                f"Espera unos minutos y vuelve a intentar, o activa "
                f"RECIBIDOS_MANUAL_CONSULTA=1 para resolverlo manualmente. "
                f"Última alerta del portal: {ultima_alerta_captcha!r}"
            )
            logger.error(mensaje_final)
            print(f"[ERROR] {mensaje_final}")
            try:
                _notificar_usuario_accion(
                    "[CAPTCHA] El SRI/Google rechazó el captcha. "
                    "Espera unos minutos o realiza la consulta manualmente."
                )
            except Exception:
                pass
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
