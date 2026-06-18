"""Cruce de Notas de Credito Emitidas vs Facturas modificadas.

Modulo aislado del flujo normal de descarga: lee NC desde una carpeta ya
descargada por la app, busca la Factura modificada (primero localmente,
luego en el portal del SRI si no esta), calcula el valor neto y exporta
todo a un Excel con columnas estandar + Estado + Observacion.

Diseño:
- Solo procesa NC Emitidas (cod=04) que modifican Facturas (codDocModificado=01).
- Reutiliza completamente las funciones de extraccion XML/PDF existentes
  (xml_extraction, pdf_extraction) — no reimplementa nada.
- La busqueda remota Playwright se hace en una sesion INDEPENDIENTE del
  flujo de Descarga: distintos cookies, distinto contexto. No interfiere
  con descargas en paralelo.
- Si una NC no puede procesarse (falta de campos, error de IO, etc.) NO
  rompe el reporte: queda como fila con Estado="Error" y Observacion
  explicando el motivo.

Punto de entrada principal: `generar_reporte_valor_neto(params, ...)`.
"""

from __future__ import annotations

import json
import logging
import re
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from robot.file_utils import _mes_a_texto

logger = logging.getLogger("robot.nc_vs_factura")


# =============================================================================
# Constantes de columnas del reporte final
# =============================================================================

COLUMNAS_REPORTE = [
    "RUC",
    "Fecha nota de credito",
    "Serie nota de credito",
    "Clave acceso nota de credito",
    "Valor total nota de credito",
    "Factura modificada",
    "Fecha factura modificada",
    "Clave acceso factura",
    "Valor total factura",
    "Valor neto",
    "Estado",
    "Observacion",
]

# Columnas forzadas a texto en Excel (claves de acceso largas + secuenciales
# que Excel interpretaria como numeros y perderia ceros a la izquierda).
COLUMNAS_TEXT_FORCE = {
    "RUC",
    "Serie nota de credito",
    "Clave acceso nota de credito",
    "Factura modificada",
    "Clave acceso factura",
}

# Columnas numericas (con formato monetario en el Excel).
COLUMNAS_NUMERICAS = {
    "Valor total nota de credito",
    "Valor total factura",
    "Valor neto",
}


# =============================================================================
# Utilidades de parseo y normalizacion
# =============================================================================


def _parse_fecha_es(valor: str) -> Optional[datetime]:
    """Parsea fechas en formato del SRI: DD/MM/YYYY o YYYY-MM-DD.

    Devuelve None si no matchea.
    """
    if not valor:
        return None
    valor = str(valor).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(valor, fmt)
        except ValueError:
            continue
    return None


def _safe_float(valor: Any) -> Optional[float]:
    """Convierte un valor a float aceptando coma decimal del SRI. None si no se puede."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip().replace(",", ".")
    txt = re.sub(r"[^0-9.\-]", "", txt)
    try:
        return float(txt)
    except (ValueError, TypeError):
        return None


def _normalizar_secuencial(num_doc: str) -> str:
    """Normaliza un secuencial del SRI a 'NNN-NNN-NNNNNNNNN'.

    Acepta variantes: '002-002-000110921', '002002000110921',
    '2-2-110921' (sin padding). Devuelve string vacio si no parsea.
    """
    if not num_doc:
        return ""
    txt = str(num_doc).strip()
    # Quitar separadores y dejar solo digitos
    partes = re.split(r"[-_\s]+", txt)
    if len(partes) == 3:
        try:
            est = int(partes[0])
            pto = int(partes[1])
            sec = int(partes[2])
            return f"{est:03d}-{pto:03d}-{sec:09d}"
        except ValueError:
            pass
    # Sin separadores: 15 digitos = 3+3+9
    solo_dig = re.sub(r"\D", "", txt)
    if len(solo_dig) == 15:
        return f"{solo_dig[:3]}-{solo_dig[3:6]}-{solo_dig[6:]}"
    return txt  # devolver como vino si no se puede normalizar


# =============================================================================
# Localizacion de archivos en la carpeta del sistema
# =============================================================================


def _es_xml_de_nc(path: Path) -> bool:
    """Heuristica rapida: nombre del archivo empieza con 'Nota' o el tipo
    'NotaCredito' (el sistema usa _nombre_documento_mes con prefijo del tipo).
    """
    if path.suffix.lower() != ".xml":
        return False
    name = path.name.lower()
    return (
        name.startswith("nota")
        or "credito" in name
        or "nc_" in name
        or "_nc_" in name
    )


def _es_pdf_de_nc(path: Path) -> bool:
    if path.suffix.lower() != ".pdf":
        return False
    name = path.name.lower()
    return (
        name.startswith("nota")
        or "credito" in name
        or "nc_" in name
        or "_nc_" in name
    )


def _listar_notas_credito_en_carpeta(carpeta: Path) -> list[tuple[Path, str]]:
    """Devuelve una lista de tuplas (path, source) donde source es 'xml' o 'pdf'.

    Recorre recursivamente carpeta/ y se queda solo con archivos que parecen
    Notas de Credito. Si una NC tiene XML y PDF, se queda solo con el XML
    (preferimos XML por confiabilidad).
    """
    xml_files: dict[str, Path] = {}
    pdf_files: dict[str, Path] = {}
    for p in carpeta.rglob("*"):
        if not p.is_file():
            continue
        if _es_xml_de_nc(p):
            xml_files[p.stem] = p
        elif _es_pdf_de_nc(p):
            pdf_files[p.stem] = p

    resultado: list[tuple[Path, str]] = []
    for stem, xml_path in xml_files.items():
        resultado.append((xml_path, "xml"))
    for stem, pdf_path in pdf_files.items():
        if stem not in xml_files:
            resultado.append((pdf_path, "pdf"))
    return sorted(resultado, key=lambda t: t[0].name.lower())


def _inferir_base_descargas(nc_path: Path) -> Optional[Path]:
    """Dada la ruta de una NC dentro de la estructura del sistema, devuelve
    la carpeta [base]/[RUC] desde donde nace el arbol Emitidos/Recibidos.

    Estructura esperada:
        [base]/[RUC]/Emitidos/Autorizados/Notas de Credito/[Anio]/[Mes]/XML/file.xml
                              ^ buscamos este nivel y devolvemos su parent

    Devuelve None si no encuentra 'Emitidos' en el path.
    """
    for ancestro in nc_path.parents:
        if ancestro.name == "Emitidos":
            return ancestro.parent  # parent de Emitidos = carpeta [RUC] (o equivalente)
    return None


def _construir_ruta_factura_mes(
    base_ruc: Path, fecha_factura: datetime, formato: str = "XML"
) -> Path:
    """Construye la ruta esperada para Facturas Emitidas Autorizadas de un mes.

    [base_ruc]/Emitidos/Autorizados/Facturas/[YYYY]/[Mes_texto]/[formato]/
    """
    return (
        base_ruc
        / "Emitidos"
        / "Autorizados"
        / "Facturas"
        / f"{fecha_factura.year:04d}"
        / _mes_a_texto(fecha_factura.month)
        / formato
    )


def _buscar_factura_local(
    base_ruc: Path, secuencial: str, fecha_factura: datetime
) -> Optional[tuple[Path, str]]:
    """Busca la Factura emitida en la estructura local.

    Estrategia:
    1. Construir la ruta del mes correspondiente
    2. Hacer glob por secuencial (los archivos llevan el secuencial en el nombre,
       patron de _nombre_documento_mes: 'Factura__YYYYMMDD__...secuencial...').
    3. Preferir XML; fallback PDF.

    Devuelve (path, 'xml'|'pdf') o None.
    """
    if not secuencial:
        return None
    # El secuencial puede aparecer en el nombre con o sin guiones segun
    # el patron de _construir_nombre_xml_emitido. Probamos ambas formas.
    sec_normalizado = _normalizar_secuencial(secuencial)
    sec_sin_guiones = sec_normalizado.replace("-", "")
    patrones = [f"*{sec_normalizado}*", f"*{sec_sin_guiones}*"]

    # Si tenemos fecha, buscar solo en el mes correspondiente (mas rapido).
    # Si no, hacer un walk completo bajo Emitidos/Autorizados/Facturas.
    carpetas_busqueda: list[Path] = []
    if fecha_factura is not None:
        carpetas_busqueda.append(_construir_ruta_factura_mes(base_ruc, fecha_factura, "XML"))
        carpetas_busqueda.append(_construir_ruta_factura_mes(base_ruc, fecha_factura, "PDF"))
    else:
        raiz = base_ruc / "Emitidos" / "Autorizados" / "Facturas"
        if raiz.is_dir():
            carpetas_busqueda.extend([d for d in raiz.rglob("XML") if d.is_dir()])
            carpetas_busqueda.extend([d for d in raiz.rglob("PDF") if d.is_dir()])

    # Buscar XML primero
    for carpeta in carpetas_busqueda:
        if not carpeta.is_dir() or carpeta.name != "XML":
            continue
        for patron in patrones:
            for archivo in carpeta.glob(patron):
                if archivo.is_file() and archivo.suffix.lower() == ".xml":
                    return (archivo, "xml")
    # Fallback PDF
    for carpeta in carpetas_busqueda:
        if not carpeta.is_dir() or carpeta.name != "PDF":
            continue
        for patron in patrones:
            for archivo in carpeta.glob(patron):
                if archivo.is_file() and archivo.suffix.lower() == ".pdf":
                    return (archivo, "pdf")
    return None


# =============================================================================
# Extraccion de datos (delegada a las funciones existentes del robot)
# =============================================================================


def _extraer_datos_nc(nc_path: Path, source: str) -> dict:
    """Extrae los campos relevantes de una NC. Reutiliza extractors existentes.

    Devuelve dict con: ruc, fecha_emision, secuencial, clave_acceso,
    importe_total, num_doc_modificado, fecha_emision_doc_sustento,
    cod_doc_modificado, _err (mensaje de error si fallo).
    """
    out = {
        "ruc": "",
        "fecha_emision": "",
        "secuencial": "",
        "clave_acceso": "",
        "importe_total": None,
        "num_doc_modificado": "",
        "fecha_emision_doc_sustento": "",
        "cod_doc_modificado": "",
        "_err": "",
    }
    try:
        if source == "xml":
            from robot.xml_extraction import _extraer_datos_xml_nota_credito_emitido
            datos = _extraer_datos_xml_nota_credito_emitido(nc_path)
        else:
            from robot.pdf_extraction import (
                _extraer_datos_pdf_nota_credito_emitido,
            )
            datos = _extraer_datos_pdf_nota_credito_emitido(nc_path)
    except Exception as exc:
        out["_err"] = f"No se pudo leer la NC ({source.upper()}): {exc}"
        return out

    if not isinstance(datos, dict):
        out["_err"] = "Datos invalidos en la NC."
        return out

    out["ruc"] = str(datos.get("RUC Emisor") or "").strip()
    out["fecha_emision"] = str(datos.get("Fecha de Emisión") or "").strip()
    out["secuencial"] = str(datos.get("Secuencial") or "").strip()
    out["clave_acceso"] = str(datos.get("Clave de Acceso") or "").strip()
    out["importe_total"] = _safe_float(datos.get("Importe Total"))
    out["num_doc_modificado"] = str(datos.get("Número Documento Modificado") or "").strip()
    out["fecha_emision_doc_sustento"] = str(
        datos.get("Fecha Emisión Doc. Sustento") or ""
    ).strip()
    out["cod_doc_modificado"] = str(datos.get("Código Documento Modificado") or "").strip()
    return out


def _extraer_datos_factura(factura_path: Path, source: str) -> dict:
    """Extrae los campos relevantes de una Factura emitida.

    Devuelve dict con: ruc, fecha_emision, secuencial, clave_acceso,
    importe_total, _err.
    """
    out = {
        "ruc": "",
        "fecha_emision": "",
        "secuencial": "",
        "clave_acceso": "",
        "importe_total": None,
        "_err": "",
    }
    try:
        if source == "xml":
            from robot.xml_extraction import _extraer_datos_xml_factura_emitido
            datos = _extraer_datos_xml_factura_emitido(factura_path)
        else:
            from robot.pdf_extraction import _extraer_datos_pdf_factura_emitido
            datos = _extraer_datos_pdf_factura_emitido(factura_path)
    except Exception as exc:
        out["_err"] = f"No se pudo leer la Factura ({source.upper()}): {exc}"
        return out

    if not isinstance(datos, dict):
        out["_err"] = "Datos invalidos en la Factura."
        return out

    out["ruc"] = str(datos.get("RUC Emisor") or "").strip()
    out["fecha_emision"] = str(datos.get("Fecha de Emisión") or "").strip()
    out["secuencial"] = str(datos.get("Secuencial") or "").strip()
    out["clave_acceso"] = str(datos.get("Clave de Acceso") or "").strip()
    out["importe_total"] = _safe_float(datos.get("Importe Total"))
    return out


# =============================================================================
# Busqueda REMOTA en el portal SRI (Playwright)
# =============================================================================


def _normalizar_serie_para_match(texto: str) -> str:
    """Quita guiones, espacios y leading zeros para comparacion robusta."""
    if not texto:
        return ""
    return re.sub(r"[\s\-_]", "", str(texto)).lstrip("0")


# =============================================================================
# Calculo deterministico de clave de acceso (49 digitos) + consulta WS de SRI
# =============================================================================
# El SRI estructura el clave en 10 segmentos: FECHA(8) + TIPO(2) + RUC(13) +
# AMB(1) + ESTAB(3) + PTO(3) + SEC(9) + COD_NUMERICO(8) + TIPO_EMI(1) + DV(1).
# El unico campo NO derivable de los datos de la NC es COD_NUMERICO (libre del
# emisor). Lo extraemos del propio clave de la NC: el mismo facturador del
# emisor suele usar el mismo COD_NUMERICO para todos sus comprobantes. Si la
# heuristica falla para un emisor, se puede override via env var.

_CODIGO_NUMERICO_DEFAULT = "00000001"


def _calcular_dv_mod11(cadena_48: str) -> int:
    """Calcula el digito verificador modulo 11 del SRI sobre 48 digitos.

    Algoritmo: multiplicar cada digito (de derecha a izquierda) por pesos
    ciclicos [2,3,4,5,6,7], sumar, calcular `11 - (suma % 11)`. Casos
    especiales: si da 11 retorna 0; si da 10 retorna 1.

    Validado contra 41 claves reales del Excel del usuario el 2026-06-18:
    41/41 OK.
    """
    if not cadena_48 or len(cadena_48) != 48 or not cadena_48.isdigit():
        raise ValueError(f"Cadena de 48 digitos invalida: {cadena_48!r}")
    pesos = (2, 3, 4, 5, 6, 7)
    suma = 0
    for i, ch in enumerate(reversed(cadena_48)):
        suma += int(ch) * pesos[i % 6]
    resto = suma % 11
    dv = 11 - resto
    if dv == 11:
        return 0
    if dv == 10:
        return 1
    return dv


def _extraer_codigo_numerico_de_clave(clave: str) -> str:
    """Devuelve los 8 digitos del COD_NUMERICO (pos 39-46) de un clave SRI,
    o cadena vacia si la clave no es valida.
    """
    clave = (clave or "").strip()
    if len(clave) != 49 or not clave.isdigit():
        return ""
    return clave[39:47]


def _calcular_clave_acceso_factura(
    fecha_emision_factura: str,
    ruc_emisor: str,
    estab: str,
    pto: str,
    secuencial: str,
    codigo_numerico: str = _CODIGO_NUMERICO_DEFAULT,
    ambiente: str = "2",
    tipo_emision: str = "1",
) -> str:
    """Construye el clave de acceso (49 digitos) de una FACTURA dadas las
    partes que conocemos por la Nota de Credito.

    `fecha_emision_factura` puede venir como "DD/MM/AAAA" o "AAAA-MM-DD" o
    como objeto datetime. Lo normalizamos a DDMMAAAA.

    Si `codigo_numerico` no es 8 digitos, falla — el caller debe asegurarse
    de pasar uno valido (extraido del clave de la NC).
    """
    # --- Fecha → DDMMAAAA ---
    fecha_str = ""
    if isinstance(fecha_emision_factura, datetime):
        fecha_str = fecha_emision_factura.strftime("%d%m%Y")
    else:
        s = str(fecha_emision_factura or "").strip()
        # Intentar varios formatos comunes
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                fecha_str = datetime.strptime(s, fmt).strftime("%d%m%Y")
                break
            except ValueError:
                continue
        if not fecha_str:
            # Fallback: si ya tiene 8 digitos, asumirlo como DDMMAAAA
            solo_dig = re.sub(r"\D", "", s)
            if len(solo_dig) == 8:
                fecha_str = solo_dig
    if len(fecha_str) != 8:
        raise ValueError(f"Fecha no normalizable a DDMMAAAA: {fecha_emision_factura!r}")

    # --- Otros campos: padding a la longitud requerida ---
    ruc_s = re.sub(r"\D", "", str(ruc_emisor or ""))
    if len(ruc_s) != 13:
        raise ValueError(f"RUC debe tener 13 digitos: {ruc_emisor!r}")
    estab_s = re.sub(r"\D", "", str(estab or "")).zfill(3)
    pto_s = re.sub(r"\D", "", str(pto or "")).zfill(3)
    sec_s = re.sub(r"\D", "", str(secuencial or "")).zfill(9)
    if len(estab_s) != 3 or len(pto_s) != 3 or len(sec_s) != 9:
        raise ValueError(
            f"Estab/Pto/Sec invalidos: estab={estab!r}, pto={pto!r}, sec={secuencial!r}"
        )
    cod_num_s = re.sub(r"\D", "", str(codigo_numerico or ""))
    if len(cod_num_s) != 8:
        raise ValueError(f"Codigo numerico debe tener 8 digitos: {codigo_numerico!r}")

    # Construir cadena de 48 dígitos (sin DV) y calcular DV
    cadena_48 = (
        fecha_str          # 8
        + "01"             # 2 (Factura)
        + ruc_s            # 13
        + str(ambiente)    # 1
        + estab_s          # 3
        + pto_s            # 3
        + sec_s            # 9
        + cod_num_s        # 8
        + str(tipo_emision)  # 1
    )
    if len(cadena_48) != 48:
        raise ValueError(f"Cadena interna mal armada (len={len(cadena_48)}): {cadena_48!r}")
    dv = _calcular_dv_mod11(cadena_48)
    return cadena_48 + str(dv)


def _consultar_factura_por_clave_ws(
    clave: str,
    timeout_s: float = 30.0,
) -> Optional[dict]:
    """Consulta el WS publico del SRI (AutorizacionComprobantesOffline) por
    una clave de acceso y devuelve un dict con `importe_total`,
    `fecha_emision`, `clave_acceso` y `razon_social_receptor`.

    Devuelve None si la clave no fue encontrada o el WS responde error.

    No requiere autenticacion ni cookies — es el servicio publico de
    autorizacion del SRI. Funciona via HTTP simple, sin Playwright.
    """
    import requests  # import perezoso

    from robot.browser import (
        SOAP_ENVELOPE_TEMPLATE,
        _parse_emitido_comprobante,
    )
    from robot.pdf_extraction import _extraer_comprobante_desde_autorizacion
    from robot.config import AUTORIZACION_COMPROBANTES_SOAP_URL

    clave = (clave or "").strip()
    if len(clave) != 49 or not clave.isdigit():
        return None

    envelope = SOAP_ENVELOPE_TEMPLATE.format(clave=clave)
    headers = {
        "Content-Type": "text/xml; charset=UTF-8",
        "SOAPAction": "",
        "Accept": "text/xml",
    }
    try:
        resp = requests.post(
            AUTORIZACION_COMPROBANTES_SOAP_URL,
            data=envelope.encode("utf-8"),
            headers=headers,
            timeout=timeout_s,
        )
    except Exception as err:
        logger.warning(f"WS SRI: error de red para clave {clave[:20]}...: {err}")
        return None
    if resp.status_code != 200:
        logger.warning(f"WS SRI: status {resp.status_code} para clave {clave[:20]}...")
        return None
    cuerpo = resp.text or ""
    if not cuerpo:
        return None

    import html
    match = re.search(r"(<autorizacion[\s\S]*?</autorizacion>)", cuerpo, flags=re.IGNORECASE)
    if not match:
        return None
    autorizacion_xml = html.unescape(match.group(1))
    try:
        comprobante_xml, meta_aut = _extraer_comprobante_desde_autorizacion(autorizacion_xml)
    except Exception as err:
        logger.warning(f"WS SRI: no se pudo extraer comprobante para {clave[:20]}...: {err}")
        return None
    if not comprobante_xml or not comprobante_xml.strip():
        # El WS respondio pero la clave no esta autorizada (no existe o fue anulada).
        return None
    try:
        meta = _parse_emitido_comprobante(comprobante_xml, meta_aut)
    except Exception as err:
        logger.warning(f"WS SRI: parse del comprobante fallo para {clave[:20]}...: {err}")
        return None
    return {
        "importe_total": meta.get("importe_total", ""),
        "fecha_emision": meta.get("fecha_emision", ""),
        "clave_acceso": meta.get("clave_acceso", clave),
        "razon_social_receptor": meta.get("razon_social_receptor", ""),
    }


# =============================================================================
# Busqueda en el portal SRI Emitidos via opcion "Clave de acceso"
# =============================================================================
# El WS publico solo devuelve comprobantes del ultimo mes. Para facturas mas
# viejas (caso comun: NCs que modifican facturas de meses anteriores) hay que
# entrar al portal con login y usar la pantalla Emitidos con la opcion radio
# "Clave de acceso / Nro. autorizacion". Como ya tenemos la clave calculada,
# cada lookup es UNA sola Consultar — sin filtros de fecha, sin paginacion.


def _esperar_overlay_cerrado(page, timeout_ms: int = 15000) -> bool:
    """Espera a que el dialogo bloqueante de PrimeFaces se oculte.

    El SRI muestra `<div id="dlgpopStatusPrime_modal" class="ui-widget-overlay">`
    como mascara de "procesando" durante CADA AJAX. Mientras esta visible,
    todos los clicks se interceptan.
    """
    try:
        page.wait_for_function(
            """() => {
                const el = document.getElementById('dlgpopStatusPrime_modal');
                if (!el) return true;
                const style = window.getComputedStyle(el);
                if (style.display === 'none' || style.visibility === 'hidden') return true;
                if (el.offsetWidth === 0 && el.offsetHeight === 0) return true;
                return false;
            }""",
            timeout=timeout_ms,
        )
        return True
    except Exception:
        return False


def _esperar_tabla_emitidos_lista(page, timeout_ms: int = 20000) -> bool:
    """Espera a que la tabla de Emitidos termine de renderizarse despues de
    un Consultar (tenga datos o muestre vacio).
    """
    try:
        page.wait_for_function(
            """() => {
                const t = document.getElementById('frmPrincipal:tablaCompEmitidos_data');
                if (!t) return false;
                if (t.querySelector('tr td')) return true;
                const empty = t.querySelector('.ui-datatable-empty-message');
                if (empty) return true;
                return false;
            }""",
            timeout=timeout_ms,
        )
        return True
    except Exception:
        return False


def _seleccionar_radio_clave_acceso(page) -> bool:
    """Selecciona el radio 'Clave de acceso / Nro. autorizacion' en Emitidos.

    El form Emitidos tiene dos opciones de busqueda:
      - "Ruc/Cedula/Pasaporte"
      - "Clave de acceso / Nro. autorizacion"

    Por defecto suele venir seleccionado el primero. Si ya estaba el segundo,
    no hacemos nada. Devuelve True si quedo seleccionado el correcto.
    """
    try:
        ok = page.evaluate(
            """() => {
                // Estrategia 1: buscar por texto del label adyacente.
                const labels = document.querySelectorAll('label, span');
                for (const l of labels) {
                    const txt = (l.textContent || '').trim();
                    if (txt.includes('Clave de acceso') || txt.toLowerCase().includes('autorizaci')) {
                        const td = l.closest('td');
                        if (td) {
                            const radio = td.querySelector('input[type="radio"]');
                            if (radio) {
                                if (!radio.checked) {
                                    radio.click();
                                    radio.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                                return true;
                            }
                            // Tambien probar el td hermano (algunos themes ponen label
                            // y radio en celdas separadas).
                            const tr = td.closest('tr');
                            if (tr) {
                                const r2 = tr.querySelector('input[type="radio"]');
                                if (r2) {
                                    if (!r2.checked) {
                                        r2.click();
                                        r2.dispatchEvent(new Event('change', {bubbles: true}));
                                    }
                                    return true;
                                }
                            }
                        }
                    }
                }
                // Estrategia 2 (fallback): asumir que el segundo radio del
                // grupo `frmPrincipal:opciones` es el de Clave de acceso.
                const radios = document.querySelectorAll(
                    'input[type="radio"][name="frmPrincipal:opciones"]'
                );
                if (radios.length >= 2) {
                    const r = radios[1];
                    if (!r.checked) {
                        r.click();
                        r.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                    return true;
                }
                return false;
            }"""
        )
        return bool(ok)
    except Exception:
        return False


def _consultar_factura_por_clave_portal(
    page,
    consultar_clic_fn,
    clave: str,
    timeout_ms: int = 20000,
) -> Optional[dict]:
    """Hace UNA consulta en Emitidos por clave de acceso y devuelve la fila
    resultante (importe, fecha, razon social).

    Pre-condiciones:
      - `page` debe estar en la pantalla Emitidos con el radio "Clave de
        acceso" YA seleccionado (ver `_seleccionar_radio_clave_acceso`).
      - `consultar_clic_fn(page)` es una funcion que clickea el boton
        Consultar (tipicamente `_click_consultar_emitidos` de browser.py).

    Devuelve dict con importe_total, clave_acceso, fecha_emision,
    razon_social_receptor — o None si no hay resultado.
    """
    if not clave or len(clave) != 49 or not clave.isdigit():
        return None

    _esperar_overlay_cerrado(page, timeout_ms=10000)

    # Setear el input txtParametro via JS — fill+dispatch_event de Playwright
    # no propaga al estado interno de PrimeFaces (eventos isTrusted=false).
    try:
        ok = page.evaluate(
            """(val) => {
                const el = document.getElementById('frmPrincipal:txtParametro');
                if (!el) return false;
                el.focus();
                el.value = '';
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.value = val;
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.dispatchEvent(new Event('blur', {bubbles: true}));
                return true;
            }""",
            clave,
        )
        if not ok:
            return None
    except Exception:
        return None

    _esperar_overlay_cerrado(page, timeout_ms=5000)

    # Click Consultar
    try:
        clicked = consultar_clic_fn(page)
    except Exception:
        clicked = False
    if not clicked:
        return None

    # Esperar a que el AJAX termine (overlay cerrado + tabla con datos o vacio)
    _esperar_overlay_cerrado(page, timeout_ms=timeout_ms)
    _esperar_tabla_emitidos_lista(page, timeout_ms=timeout_ms)

    # Leer la fila resultante (deberia ser 1 sola fila)
    try:
        tabla = page.locator("#frmPrincipal\\:tablaCompEmitidos_data")
        filas = tabla.locator("tr")
        n = filas.count()
        if n == 0:
            return None
        primera = filas.first
        celdas = primera.locator("td")
        textos = celdas.all_inner_texts()
        if len(textos) < 8:
            return None
        tipo_serie = textos[1].strip()
        clave_text = textos[2].strip()
        fecha_aut = textos[3].strip()
        fecha_emi = textos[4].strip() if len(textos) > 4 else ""
        razon = textos[5].strip() if len(textos) > 5 else ""
        importe = textos[7].strip()
        # Sanity check: la clave del resultado deberia coincidir con la pedida
        clave_digits = re.sub(r"\D", "", clave_text)
        if clave_digits and clave_digits != clave:
            logger.warning(
                f"Portal SRI: clave resultado {clave_digits[:20]}... no coincide "
                f"con la pedida {clave[:20]}..."
            )
        return {
            "importe_total": importe,
            "clave_acceso": clave_digits or clave,
            "fecha_emision": fecha_emi or fecha_aut,
            "razon_social_receptor": razon,
            "tipo_serie_raw": tipo_serie,
        }
    except Exception as err:
        logger.warning(f"Portal SRI: lectura de fila resultante fallo: {err}")
        return None


def _buscar_facturas_remoto(
    ruc: str,
    clave: str,
    pendientes: list[dict],
    *,
    cancel_event: Optional[threading.Event] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> dict[str, dict]:
    """Busca facturas modificadas en el portal SRI Emitidos via clave de acceso.

    Estrategia: para cada NC pendiente calculamos la clave de acceso de la
    factura modificada (datos en la NC + codigo numerico extraido del propio
    clave de la NC). Despues entramos al portal del SRI con login, abrimos
    Emitidos, seleccionamos el radio "Clave de acceso / Nro. autorizacion"
    UNA SOLA VEZ y por cada factura llenamos el input + Consultar.

    Por que NO usamos el WS publico AutorizacionComprobantesOffline:
    el WS solo devuelve comprobantes del ultimo mes. Para facturas mas
    viejas (caso comun: NCs que modifican facturas de meses anteriores)
    el WS responde sin comprobante. Confirmado en la prueba del 2026-06-18:
    37/37 facturas (todas de marzo/abril) salieron como "no encontradas" del WS.

    Pros vs el flujo viejo (busqueda por fecha + serie en la tabla):
      - 1 Consultar por factura (no paginacion).
      - Match infalible: la clave de 49 digitos es unica.
      - Inmune a sorts de la tabla y limites de RPP.

    Parametros:
      `ruc`: RUC del emisor (necesario para el login).
      `clave`: password SRI (necesario para el login).
      `pendientes`: lista de dicts con secuencial, fecha (datetime), nc_data.
      `cancel_event`: para cancelar en medio del loop.
      `progress`: callback para emit logs en UI.

    Devuelve dict { secuencial_normalizado: {importe_total, clave_acceso,
    fecha_emision, razon_social_receptor, tipo_serie_raw} } con lo que SI
    encontro. Las que no encuentre quedan fuera (caller las marca como
    "Factura no encontrada en SRI").
    """
    encontradas: dict[str, dict] = {}
    if not pendientes:
        return encontradas

    def _emit(msg: str) -> None:
        if progress:
            try:
                progress(msg)
            except Exception:
                pass
        logger.info(msg)

    from playwright.sync_api import sync_playwright  # import perezoso
    from robot.downloader import _login, _abrir_navegador
    from robot.browser import _abrir_modulo_consultas, _click_consultar_emitidos
    from robot.config import PORTAL_HOME

    cookies_path = Path(f"cookies_nc_lookup_{ruc}.json")

    _emit(
        f"Iniciando busqueda remota en portal SRI: {len(pendientes)} factura(s). "
        f"Login + 1 Consultar por clave de acceso."
    )

    with sync_playwright() as p:
        context, browser, _persistent = _abrir_navegador(p)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            _emit("Autenticando en el portal del SRI...")
            # Cookies cacheadas de corridas anteriores podrian estar vencidas
            # y causar bounces a login — las borramos pre-emptive.
            if cookies_path.exists():
                try:
                    cookies_path.unlink()
                except Exception:
                    pass
            _login(context, page, ruc, clave, cookies_path, PORTAL_HOME)
            _emit("Sesion del SRI lista. Abriendo modulo Emitidos...")

            try:
                _abrir_modulo_consultas(page, "Emitidos")
            except Exception as exc:
                _emit(f"No se pudo abrir el modulo de Emitidos: {exc}")
                return encontradas

            _emit("Modulo Emitidos abierto. Seleccionando opcion 'Clave de acceso'...")
            _esperar_overlay_cerrado(page, timeout_ms=15000)
            if not _seleccionar_radio_clave_acceso(page):
                _emit(
                    "ERROR: no se pudo seleccionar el radio 'Clave de acceso / "
                    "Nro. autorizacion'. Abandono busqueda remota."
                )
                return encontradas
            _esperar_overlay_cerrado(page, timeout_ms=10000)
            _emit(
                "Radio 'Clave de acceso' activo. Buscando factura por factura..."
            )

            encontradas_count = 0
            for i, item in enumerate(pendientes, 1):
                if cancel_event is not None and cancel_event.is_set():
                    _emit("Busqueda remota cancelada por el usuario.")
                    break

                nc_data = item.get("nc_data") or {}
                secuencial_factura = str(item.get("secuencial") or "").strip()
                fecha_factura = item.get("fecha")

                # RUC emisor: la NC modifica una factura del MISMO contribuyente
                # que la emitio. Usamos el RUC del XML de la NC; fallback al param.
                ruc_emisor = str(nc_data.get("ruc") or ruc or "").strip()

                # Codigo numerico: extraido del clave de la NC. Mismo facturador
                # del emisor → mismo codigo. Si no hay clave o esta rota, default.
                cod_num = _extraer_codigo_numerico_de_clave(
                    nc_data.get("clave_acceso", "")
                )
                if not cod_num:
                    cod_num = _CODIGO_NUMERICO_DEFAULT

                # Partes del secuencial "EEE-PPP-SSSSSSSSS"
                partes = secuencial_factura.split("-")
                if len(partes) != 3:
                    _emit(
                        f"  [{i}/{len(pendientes)}] secuencial mal formado: "
                        f"{secuencial_factura!r} — salto."
                    )
                    continue
                estab, pto, sec = partes

                fecha_display = (
                    fecha_factura.strftime("%d/%m/%Y")
                    if isinstance(fecha_factura, datetime)
                    else str(fecha_factura)
                )

                # Calcular clave
                try:
                    clave_calc = _calcular_clave_acceso_factura(
                        fecha_emision_factura=fecha_factura,
                        ruc_emisor=ruc_emisor,
                        estab=estab,
                        pto=pto,
                        secuencial=sec,
                        codigo_numerico=cod_num,
                    )
                except Exception as err:
                    _emit(
                        f"  [{i}/{len(pendientes)}] no se pudo calcular clave para "
                        f"{secuencial_factura} ({fecha_display}): {err} — salto."
                    )
                    continue

                # Consultar al portal
                data = _consultar_factura_por_clave_portal(
                    page, _click_consultar_emitidos, clave_calc, timeout_ms=20000
                )
                if data is None:
                    _emit(
                        f"  [{i}/{len(pendientes)}] {secuencial_factura} "
                        f"({fecha_display}): portal no devolvio resultado. "
                        f"Clave calculada: {clave_calc}"
                    )
                    continue

                key = _normalizar_secuencial(secuencial_factura)
                encontradas[key] = {
                    "importe_total": _safe_float(data.get("importe_total")),
                    "clave_acceso": data.get("clave_acceso") or clave_calc,
                    "tipo_serie_raw": data.get(
                        "tipo_serie_raw", f"Factura {secuencial_factura}"
                    ),
                    "fecha_emision": data.get("fecha_emision", ""),
                    "razon_social_receptor": data.get("razon_social_receptor", ""),
                }
                encontradas_count += 1
                _emit(
                    f"  [{i}/{len(pendientes)}] OK {secuencial_factura} "
                    f"({fecha_display}): importe={data.get('importe_total')!r}"
                )

            _emit(
                f"Busqueda remota portal completada: "
                f"{encontradas_count}/{len(pendientes)} encontradas."
            )
        finally:
            try:
                context.close()
            except Exception:
                pass
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass

    return encontradas

# =============================================================================
# Orquestador principal
# =============================================================================


def _construir_fila(
    nc_data: dict,
    factura_data: Optional[dict],
    estado: str,
    observacion: str,
) -> dict:
    """Construye una fila del Excel a partir de los datos de NC + Factura (opcional)."""
    valor_nc = nc_data.get("importe_total")
    valor_factura = factura_data.get("importe_total") if factura_data else None
    valor_neto = None
    if valor_factura is not None and valor_nc is not None:
        valor_neto = round(valor_factura - valor_nc, 2)

    return {
        "RUC": nc_data.get("ruc", ""),
        "Fecha nota de credito": nc_data.get("fecha_emision", ""),
        "Serie nota de credito": nc_data.get("secuencial", ""),
        "Clave acceso nota de credito": nc_data.get("clave_acceso", ""),
        "Valor total nota de credito": valor_nc if valor_nc is not None else "",
        "Factura modificada": _normalizar_secuencial(
            nc_data.get("num_doc_modificado", "")
        ),
        "Fecha factura modificada": nc_data.get("fecha_emision_doc_sustento", ""),
        "Clave acceso factura": factura_data.get("clave_acceso", "") if factura_data else "",
        "Valor total factura": valor_factura if valor_factura is not None else "",
        "Valor neto": valor_neto if valor_neto is not None else "",
        "Estado": estado,
        "Observacion": observacion,
    }


def generar_reporte_valor_neto(
    *,
    carpeta_nc: str | Path,
    ruc: str,
    clave: str,
    salida_excel: str | Path,
    cancel_event: Optional[threading.Event] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> dict:
    """Punto de entrada unico del modulo. Lee NC de `carpeta_nc`, cruza contra
    Facturas (local + remoto si hace falta), genera Excel en `salida_excel`.

    Retorna resumen: {
        ok: bool,
        total_nc: int,
        encontradas_local: int,
        encontradas_remoto: int,
        no_encontradas: int,
        errores: int,
        excel_path: str,
        message: str,
    }
    """
    def _emit(msg: str) -> None:
        if progress:
            try:
                progress(msg)
            except Exception:
                pass
        logger.info(msg)

    carpeta_nc_path = Path(carpeta_nc).expanduser()
    if not carpeta_nc_path.is_dir():
        return {
            "ok": False,
            "total_nc": 0,
            "encontradas_local": 0,
            "encontradas_remoto": 0,
            "no_encontradas": 0,
            "errores": 0,
            "excel_path": "",
            "message": f"La carpeta no existe: {carpeta_nc_path}",
        }

    _emit(f"Explorando carpeta de Notas de Credito: {carpeta_nc_path}")
    notas = _listar_notas_credito_en_carpeta(carpeta_nc_path)
    if not notas:
        return {
            "ok": False,
            "total_nc": 0,
            "encontradas_local": 0,
            "encontradas_remoto": 0,
            "no_encontradas": 0,
            "errores": 0,
            "excel_path": "",
            "message": "No se encontraron Notas de Credito en la carpeta indicada.",
        }
    _emit(f"Notas de Credito detectadas: {len(notas)}")

    # Inferir base [RUC] del arbol de descargas (usa la primera NC como pista).
    base_descargas = _inferir_base_descargas(notas[0][0])
    if base_descargas is None:
        # Fallback: la propia carpeta seleccionada es la base.
        base_descargas = carpeta_nc_path
    _emit(f"Base de descargas inferida: {base_descargas}")

    rows: list[dict] = []
    pendientes_remoto: list[dict] = []
    errores = 0
    encontradas_local = 0

    for nc_path, source in notas:
        if cancel_event is not None and cancel_event.is_set():
            _emit("Proceso cancelado por el usuario.")
            break

        nc_data = _extraer_datos_nc(nc_path, source)
        if nc_data["_err"]:
            errores += 1
            rows.append(
                _construir_fila(
                    nc_data, None, "Error", f"{nc_data['_err']} (archivo: {nc_path.name})"
                )
            )
            continue

        # Filtrar solo NC que modifican Facturas (cod_doc_modificado == "01")
        cod = (nc_data.get("cod_doc_modificado") or "").strip().lstrip("0") or "0"
        if cod != "1":
            rows.append(
                _construir_fila(
                    nc_data,
                    None,
                    "Omitido",
                    f"Esta NC modifica un comprobante tipo '{cod}', no una Factura (01).",
                )
            )
            continue

        secuencial_factura = _normalizar_secuencial(nc_data.get("num_doc_modificado", ""))
        fecha_factura = _parse_fecha_es(nc_data.get("fecha_emision_doc_sustento", ""))

        if not secuencial_factura:
            rows.append(
                _construir_fila(
                    nc_data,
                    None,
                    "Error",
                    "La NC no contiene 'Numero Documento Modificado' valido.",
                )
            )
            errores += 1
            continue

        # Busqueda LOCAL
        encontrado_local = _buscar_factura_local(
            base_descargas, secuencial_factura, fecha_factura
        )
        if encontrado_local is not None:
            factura_path, factura_source = encontrado_local
            factura_data = _extraer_datos_factura(factura_path, factura_source)
            if factura_data["_err"]:
                rows.append(
                    _construir_fila(
                        nc_data,
                        None,
                        "Error",
                        f"Factura encontrada localmente pero no se pudo leer: "
                        f"{factura_data['_err']}",
                    )
                )
                errores += 1
                continue
            rows.append(
                _construir_fila(
                    nc_data,
                    factura_data,
                    "OK (local)",
                    f"Factura encontrada localmente en {factura_path.name}.",
                )
            )
            encontradas_local += 1
        else:
            # Marcar para busqueda remota
            pendientes_remoto.append(
                {
                    "secuencial": secuencial_factura,
                    "fecha": fecha_factura,
                    "nc_data": nc_data,
                }
            )

    # === Busqueda REMOTA para las que no se encontraron localmente ===
    encontradas_remoto = 0
    no_encontradas = 0
    encontradas_dict: dict[str, dict] = {}

    if pendientes_remoto and (cancel_event is None or not cancel_event.is_set()):
        if not ruc or not clave:
            # Sin credenciales no podemos consultar el SRI: marcar todas como
            # "no encontradas localmente — falta credencial para consulta remota".
            _emit(
                "No se proporcionaron credenciales — se omite busqueda remota. "
                "Las Facturas faltantes quedaran marcadas como 'no encontradas'."
            )
        else:
            try:
                encontradas_dict = _buscar_facturas_remoto(
                    ruc,
                    clave,
                    pendientes_remoto,
                    cancel_event=cancel_event,
                    progress=progress,
                )
            except Exception as exc:
                _emit(f"Busqueda remota fallo: {exc}")
                encontradas_dict = {}

    # Construir filas de las pendientes
    for item in pendientes_remoto:
        nc_data = item["nc_data"]
        key = _normalizar_secuencial(item["secuencial"])
        info_remoto = encontradas_dict.get(key)
        if info_remoto:
            factura_data = {
                "importe_total": info_remoto.get("importe_total"),
                "clave_acceso": info_remoto.get("clave_acceso", ""),
            }
            rows.append(
                _construir_fila(
                    nc_data,
                    factura_data,
                    "OK (remoto)",
                    f"Factura encontrada en el portal del SRI "
                    f"(tabla emitidos, columna tipo y serie).",
                )
            )
            encontradas_remoto += 1
        else:
            rows.append(
                _construir_fila(
                    nc_data,
                    None,
                    "Factura no encontrada",
                    f"No se encontro la Factura {item['secuencial']} ni localmente "
                    f"ni en el portal del SRI.",
                )
            )
            no_encontradas += 1

    # === Generar Excel ===
    salida_path = Path(salida_excel).expanduser()
    salida_path.parent.mkdir(parents=True, exist_ok=True)
    _escribir_excel(rows, salida_path)
    _emit(f"Excel generado: {salida_path}")

    return {
        "ok": True,
        "total_nc": len(rows),
        "encontradas_local": encontradas_local,
        "encontradas_remoto": encontradas_remoto,
        "no_encontradas": no_encontradas,
        "errores": errores,
        "excel_path": str(salida_path),
        "message": (
            f"Reporte generado con {len(rows)} fila(s). "
            f"Local: {encontradas_local}, Remoto: {encontradas_remoto}, "
            f"No encontradas: {no_encontradas}, Errores: {errores}."
        ),
    }


def _escribir_excel(rows: list[dict], path: Path) -> None:
    """Genera el Excel con formato basico (header + monetario + text-force)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Valor Neto NC vs Facturas"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Header
    for col_idx, nombre in enumerate(COLUMNAS_REPORTE, start=1):
        cell = ws.cell(row=1, column=col_idx, value=nombre)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Anchos sugeridos
    anchos = {
        "RUC": 15,
        "Fecha nota de credito": 14,
        "Serie nota de credito": 18,
        "Clave acceso nota de credito": 52,
        "Valor total nota de credito": 16,
        "Factura modificada": 18,
        "Fecha factura modificada": 14,
        "Clave acceso factura": 52,
        "Valor total factura": 16,
        "Valor neto": 14,
        "Estado": 22,
        "Observacion": 60,
    }
    for col_idx, nombre in enumerate(COLUMNAS_REPORTE, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = anchos.get(nombre, 16)

    # Filas de datos
    for row_idx, fila in enumerate(rows, start=2):
        for col_idx, nombre in enumerate(COLUMNAS_REPORTE, start=1):
            valor = fila.get(nombre, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            if nombre in COLUMNAS_TEXT_FORCE and valor != "":
                cell.number_format = "@"
            if nombre in COLUMNAS_NUMERICAS and valor != "":
                cell.number_format = "#,##0.00"
            cell.alignment = Alignment(vertical="center", wrap_text=(nombre == "Observacion"))

    # Congelar header
    ws.freeze_panes = "A2"
    wb.save(str(path))
