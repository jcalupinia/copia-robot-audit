"""Automatizacion del navegador (Playwright) sobre el portal del SRI.

Toda la interaccion con el portal: navegacion y modales, manejo de
ViewState/JSF, seleccion de combos e inputs, descarga de PDF y XML
(via enlaces, formularios JSF, POST directo y SOAP), y parsing de las
respuestas parciales/AJAX.

Las funciones reciben el objeto `page` de Playwright como parametro; este
modulo NO crea el navegador ni orquesta el flujo (eso queda en
`robot/downloader.py`: login + _flujo_recibidos/_flujo_emitidos).

Extraido de `robot/downloader.py` en la Fase 3c del refactor.
"""
from __future__ import annotations

import html
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional
from urllib.parse import urlencode

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from robot._logging import get_logger
from robot.config import (
    AUTORIZACION_COMPROBANTES_SOAP_URL,
    CONSULTAS_SELECTOR,
    DOC_LABELS,
    DOWNLOAD_TIMEOUT,
    FACTURACION_MENU_SELECTOR,
    MENU_TOGGLE_SELECTOR,
    MODULO_PRODUCCION_SELECTOR,
    OVERLAY_SELECTORS,
    PORTAL_INDISPONIBLE_MENSAJE,
    RECIBIDOS_DIRECT_URL,
    RECUPERAR_COMPROBANTES_URL,
)
from robot.data_formatters import _parse_decimal
from robot.file_utils import _sanear_nombre_archivo
from robot.pdf_extraction import (
    _extraer_comprobante_desde_autorizacion,
    _limpiar_cdata,
    _normalizar_token,
)
from robot.report_columns import PDF_REPORT_COLUMNS
from robot.signals import _check_cancel, _notificar_usuario_accion, _notificar_usuario_captcha
from robot.xml_extraction import _strip_xml_namespaces


logger = get_logger(__name__)


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


def _ejecutar_post_pdf(
    page,
    referer_url: str,
    form_data: dict,
    base_destino: Path,
) -> Optional[Path]:
    """Ejecuta el POST + valida content-type + guarda el PDF a disco.
    Logica compartida entre la version Recibidos y Emitidos.
    """
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": referer_url,
    }
    try:
        respuesta = page.context.request.post(referer_url, data=form_data, headers=headers)
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
    form_base: dict | None = None,
    referer_url: str | None = None,
) -> Optional[Path]:
    """Descarga PDF de Recibidos via HTTP POST.

    Si `form_base` y `referer_url` se pasan pre-cacheados (al inicio de la
    pagina, una sola vez), evita ~11 round-trips CDP por llamada — gana
    ~300ms por fila. Si vienen `None` se construye el form base inline
    (modo legacy, mantenido por backward compat para llamadas sueltas).

    `form_base` se obtiene con `_obtener_form_base_emitidos(page)` — esa
    funcion es generica (captura todo el form#frmPrincipal con un solo
    `page.evaluate`) y sirve para ambos modulos.
    """
    if not link_id or not view_state:
        return None

    if form_base is None:
        form_base = _obtener_form_base_emitidos(page)
    if referer_url is None:
        referer_url = page.url.split("#")[0]

    form_data = dict(form_base)
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id
    return _ejecutar_post_pdf(page, referer_url, form_data, base_destino)


def _descargar_pdf_emitidos_post_con_viewstate(
    page,
    link_id: str,
    view_state: str,
    base_destino: Path,
    form_base: dict | None = None,
    referer_url: str | None = None,
) -> Optional[Path]:
    """Descarga PDF de Emitidos via HTTP POST.

    Misma optimizacion que `_descargar_pdf_recibidos_post_con_viewstate`:
    si `form_base` y `referer_url` vienen pre-cacheados, evita ~13
    round-trips CDP por llamada (~350ms por fila).
    """
    if not link_id or not view_state:
        return None

    if form_base is None:
        form_base = _obtener_form_base_emitidos(page)
    if referer_url is None:
        referer_url = page.url.split("#")[0]

    form_data = dict(form_base)
    form_data["javax.faces.ViewState"] = view_state
    form_data[link_id] = link_id
    return _ejecutar_post_pdf(page, referer_url, form_data, base_destino)


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


# ===========================================================================
# Verificadores y recuperadores del formulario de Emitidos (anti-cuelgue 30s)
# ===========================================================================
def _asegurar_en_tabla_emitidos(page, timeout: int = 1500) -> bool:
    """Verifica que la pagina sigue mostrando la tabla de comprobantes
    Emitidos (no se fue al home / perfil del SRI por timeout de sesion,
    redirect de JSF, rate limit, etc).

    Devuelve True si:
      - La URL contiene `recuperarComprobantes.jsf` o `comprobantesRecibidos.jsf`
      - El selector `#frmPrincipal:tablaCompEmitidos_data` esta presente y visible

    Devuelve False si:
      - La URL fue redirigida a `/contribuyente/perfil` o similar
      - La tabla no esta en el DOM
      - Hay cualquier error de comunicacion con la pagina

    Esta funcion es la "guardia previa" antes de clickear PDF/XML links;
    evita el cuelgue clasico de 30 segundos x N filas cuando el SRI devuelve
    al navegador a la pantalla principal.
    """
    try:
        url_actual = (page.url or "").lower()
    except Exception:
        return False
    en_formulario_url = (
        "recuperarcomprobantes.jsf" in url_actual
        or "comprobantesrecibidos.jsf" in url_actual
    )
    if not en_formulario_url:
        # En perfil/home u otra pagina — explicitamente NO estamos en el form.
        logger.warning(
            f"_asegurar_en_tabla_emitidos: URL inesperada '{page.url}' "
            f"— el portal devolvio al usuario fuera del formulario."
        )
        return False
    try:
        tabla = page.locator("#frmPrincipal\\:tablaCompEmitidos_data")
        if not tabla.count():
            logger.warning(
                "_asegurar_en_tabla_emitidos: tabla no presente en el DOM."
            )
            return False
        # is_visible con timeout corto — si tarda mas de 1.5s asumimos que
        # la tabla no esta cargada y reportamos False.
        return bool(tabla.first.is_visible(timeout=timeout))
    except Exception as err:
        logger.warning(f"_asegurar_en_tabla_emitidos: error de verificacion: {err}")
        return False


def _navegar_a_pagina_emitidos(page, pagina_destino: int, max_clicks: int = 200) -> bool:
    """Navega la tabla paginada de Emitidos hasta la `pagina_destino` (1-based)
    haciendo click en el boton "siguiente" de PrimeFaces (`>` / next button).

    Devuelve True si llego a la pagina pedida, False si se agotaron los clicks
    o no se pudo identificar el control de paginacion.

    El selector tipico del paginador de PrimeFaces en este formulario es
    `a.ui-paginator-next` dentro del wrapper de la tabla. Tras cada click
    esperamos al AJAX antes de validar.
    """
    if pagina_destino <= 1:
        return True
    paginador_actual_sel = "span.ui-paginator-current"
    next_sel = "#frmPrincipal\\:tablaCompEmitidos_paginator_top a.ui-paginator-next"
    next_sel_alt = "a.ui-paginator-next"
    for click_idx in range(max_clicks):
        # Leer indicador "(X of Y)" — si ya estamos en la pagina destino, salir.
        try:
            ind = page.locator(paginador_actual_sel)
            if ind.count():
                texto = (ind.first.inner_text() or "").strip()
                # Formato esperado: "(1 of 12)" o "(8 de 12)"
                m = re.search(r"\(\s*(\d+)\s*(?:of|de)\s*\d+\s*\)", texto, re.IGNORECASE)
                if m and int(m.group(1)) >= pagina_destino:
                    return True
        except Exception:
            pass
        # Hacer click en "siguiente"
        try:
            nxt = page.locator(next_sel)
            if not nxt.count():
                nxt = page.locator(next_sel_alt)
            if not nxt.count():
                logger.warning(
                    "_navegar_a_pagina_emitidos: no se encontro boton siguiente."
                )
                return False
            # Si el boton next esta deshabilitado, ya estamos en la ultima pagina
            try:
                cls_attr = (nxt.first.get_attribute("class") or "").lower()
                if "ui-state-disabled" in cls_attr:
                    return click_idx + 1 >= pagina_destino  # llegamos al final
            except Exception:
                pass
            nxt.first.click(timeout=3000)
            _esperar_ajax(page, timeout=4000)
            page.wait_for_timeout(200)
        except Exception as err:
            logger.warning(
                f"_navegar_a_pagina_emitidos: click siguiente fallo en iter {click_idx}: {err}"
            )
            return False
    logger.warning(
        f"_navegar_a_pagina_emitidos: agote {max_clicks} clicks sin llegar a pagina {pagina_destino}."
    )
    return False


def _recuperar_formulario_emitidos(
    page,
    *,
    tipo_visible: Optional[str] = None,
    estado_visible: Optional[str] = None,
    fecha_emision: Optional[str] = None,
    establecimiento: Optional[str] = None,
    punto_emision: Optional[str] = None,
    pagina_destino: int = 1,
) -> bool:
    """Recupera el formulario de Emitidos despues de una redireccion no deseada
    (ej. SRI devolvio el navegador a `/contribuyente/perfil`).

    Pasos:
      1. Re-abrir el modulo Consultas → Emitidos (sin re-loguear; las cookies
         vivas suelen alcanzar). Si el SRI exige login, devuelve False y
         el caller debe propagar para que el worker maneje la reautenticacion.
      2. Re-aplicar los filtros que estaban activos (tipo, estado, fecha,
         establecimiento, punto emision).
      3. Hacer click en Consultar para regenerar la tabla.
      4. Navegar paginando hasta `pagina_destino` para retomar la posicion.

    Devuelve True si la tabla quedo lista en la pagina pedida; False si fallo
    en algun paso (el caller debe registrar el error y eventualmente abortar).
    """
    try:
        _abrir_modulo_consultas(page, "Emitidos")
    except Exception as err:
        logger.warning(f"_recuperar_formulario_emitidos: no se pudo abrir modulo: {err}")
        return False

    # Re-aplicar filtros con tolerancia a fallos parciales — algun filtro
    # opcional puede no estar presente segun el caso.
    if tipo_visible:
        try:
            _seleccionar_en_select(
                page, "select#frmPrincipal\\:cmbTipoComprobante", tipo_visible
            )
        except Exception as err:
            logger.warning(f"recuperar: no se pudo set tipo='{tipo_visible}': {err}")
    if estado_visible:
        try:
            _seleccionar_en_select(
                page, "select#frmPrincipal\\:cmbEstadoAutorizacion", estado_visible
            )
        except Exception as err:
            logger.warning(f"recuperar: no se pudo set estado='{estado_visible}': {err}")
    if fecha_emision:
        try:
            fecha_loc = page.locator("input#frmPrincipal\\:calendarFechaDesde_input")
            if fecha_loc.count():
                fecha_loc.first.fill("")
                fecha_loc.first.fill(fecha_emision)
        except Exception as err:
            logger.warning(f"recuperar: no se pudo set fecha='{fecha_emision}': {err}")
    if establecimiento:
        try:
            _seleccionar_en_select(
                page, "select#frmPrincipal\\:cmbEstablecimiento", establecimiento
            )
        except Exception:
            pass
    if punto_emision:
        try:
            pto_loc = page.locator("input#frmPrincipal\\:txtPuntoEmision")
            if pto_loc.count():
                pto_loc.first.fill("")
                pto_loc.first.fill(punto_emision)
        except Exception:
            pass

    # Click Consultar
    try:
        consultar_btn = page.locator(
            "input[type='submit'][value='Consultar'], button:has-text('Consultar')"
        )
        if consultar_btn.count():
            consultar_btn.first.click()
            page.wait_for_load_state("networkidle", timeout=15000)
        _esperar_ajax(page, timeout=5000)
    except Exception as err:
        logger.warning(f"recuperar: click Consultar fallo: {err}")
        return False

    # Validar que la tabla quedo visible
    if not _asegurar_en_tabla_emitidos(page, timeout=3000):
        logger.warning("recuperar: tabla no quedo visible tras Consultar.")
        return False

    # Navegar a la pagina destino
    if pagina_destino > 1:
        if not _navegar_a_pagina_emitidos(page, pagina_destino):
            logger.warning(
                f"recuperar: no se llego a pagina {pagina_destino} despues de consultar."
            )
            return False

    return True


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
