"""Parsing de PDFs y XML de comprobantes electronicos del SRI.

Modulo de extraccion: convierte archivos PDF (via pdfplumber + regex +
analisis de layout) y XML descargados del portal en filas de reporte.

Contiene:
- Helpers de bajo nivel (lectura, regex, layout) - Sub-fase 3b-A.
- Parsers principales por tipo de comprobante (factura, nota de credito/debito,
  retencion, liquidacion de compra), tanto recibidos como emitidos - Sub-fase 3b-B.
- Los parsers de XML que dependen de extraccion por regex compartida.

`_extraer_datos_pdf` y `_extraer_datos_pdf_retencion` son las funciones mas
grandes; usan funciones anidadas internas para no contaminar el namespace.

Extraido de `robot/downloader.py` en las Sub-fases 3b-A y 3b-B del refactor.
"""
from __future__ import annotations

import html
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
except Exception:  # pragma: no cover - entorno sin pdfplumber
    pdfplumber = None

try:
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LTTextContainer, LTTextLine
except Exception:  # pragma: no cover - entorno sin pdfminer
    extract_pages = None
    LTTextContainer = None
    LTTextLine = None

try:
    from robot.pdf_layout.main import extract_pdf_fields as _extract_pdf_layout_fields
except Exception:  # pragma: no cover
    _extract_pdf_layout_fields = None

from robot import parser as xml_parser
from robot._logging import get_logger
from robot.comprobante_types import _normalizar_label_simple
from robot.data_formatters import (
    _asignar_resumen_retencion_emitidos,
    _factura_emitidos_default_row,
    _label_ambiente_emitidos_retencion,
    _label_emision_emitidos_retencion,
    _label_tipo_ident_emitidos_nota_credito,
    _map_retencion_legacy_to_emitidos_sample_row,
    _nota_credito_emitidos_default_row,
    _nota_debito_emitidos_default_row,
    _numero_emitidos_retencion,
    _parse_decimal,
    _texto_emitidos_retencion,
    _texto_emitidos_retencion_na,
)
from robot.report_columns import PDF_REPORT_COLUMNS, RETENCION_REPORT_COLUMNS
from robot.xml_extraction import (
    _extraer_xml_emitidos_autorizacion,
    _strip_xml_namespaces,
)


logger = get_logger(__name__)


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


def _fix_replacement_chars_pdf(texto: str) -> str:
    """Reemplaza el REPLACEMENT CHARACTER (U+FFFD) por 'O'.

    pdfplumber inserta `\\ufffd` cuando no logra decodificar una vocal con tilde
    del PDF del SRI. En estos comprobantes ese carácter siempre corresponde a
    Ó/Ú/É/Á/Í dentro de palabras tipo RETENCIÓN, AUTORIZACIÓN, EMISIÓN, DIRECCIÓN,
    IDENTIFICACIÓN, RAZÓN; sustituirlo por 'O' permite que las regex de keywords
    matcheen después de pasar por NFKD + ascii. No afecta los valores extraídos
    porque para esos preservamos el texto original.
    """
    return (texto or "").replace("�", "O")


def _norm_pdf_keyword(texto: str) -> str:
    """Normaliza texto del PDF para matching de keywords: sin tildes, sin
    REPLACEMENT CHAR, mayúsculas. Equivalente a `_normalizar_texto_pdf` pero
    blindado contra `\\ufffd`."""
    base = _fix_replacement_chars_pdf(texto)
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    return base.upper()


def _agrupar_palabras_visuales(words: list, y_tol: float = 3.5) -> list[list[dict]]:
    """Agrupa las palabras devueltas por `page.extract_words` en líneas visuales
    (mismo `top` ± y_tol), ordenadas por `x0` dentro de cada línea.

    Útil cuando `extract_text` mezcla columnas: agrupando por coordenadas
    podemos luego filtrar por rango de X para separar columna izquierda/derecha.
    """
    if not words:
        return []
    sw = sorted(words, key=lambda w: (float(w.get("top", 0)), float(w.get("x0", 0))))
    lineas: list[list[dict]] = []
    actual = [sw[0]]
    top_actual = float(sw[0].get("top", 0))
    for w in sw[1:]:
        t = float(w.get("top", 0))
        if abs(t - top_actual) <= y_tol:
            actual.append(w)
        else:
            lineas.append(sorted(actual, key=lambda x: float(x.get("x0", 0))))
            actual = [w]
            top_actual = t
    lineas.append(sorted(actual, key=lambda x: float(x.get("x0", 0))))
    return lineas


def _texto_palabras_rango(palabras: list[dict], x_min: float | None = None, x_max: float | None = None) -> str:
    """Une el texto de las palabras dentro del rango horizontal [x_min, x_max)."""
    partes = []
    for w in palabras:
        x = float(w.get("x0", 0))
        if x_min is not None and x < x_min:
            continue
        if x_max is not None and x >= x_max:
            continue
        t = (w.get("text") or "").strip()
        if t:
            partes.append(t)
    return " ".join(partes)


def _parse_celda_impuesto_retencion(cell_imp: str | None, cell_porc_val: str | None) -> dict | None:
    """Parsea las 2 celdas finales (base+impuesto / porcentaje+valor) de una fila
    de la tabla de retenciones del SRI.

    `cell_imp` puede venir como:
      - "12.00 IVA"                  (formato IVA en una sola línea)
      - "Impuesto a la\\n80.00\\nRenta"  (formato Renta en 3 líneas)
    `cell_porc_val` es siempre "{porcentaje} {valor}".

    Devuelve {impuesto: 'IVA'|'Renta', base, porcentaje, valor} o None si no parsea.
    """
    if not cell_imp:
        return None
    cell_norm = _norm_pdf_keyword(cell_imp)
    if "IVA" in cell_norm:
        impuesto = "IVA"
        m = re.search(r"([\d]+(?:[.,]\d+)?)\s*IVA", cell_norm)
        if not m:
            m = re.search(r"IVA\s+([\d]+(?:[.,]\d+)?)", cell_norm)
        base = m.group(1) if m else ""
    elif "RENTA" in cell_norm:
        impuesto = "Renta"
        nums = re.findall(r"\d+(?:[.,]\d+)?", cell_norm)
        base = nums[0] if nums else ""
    else:
        return None
    porcentaje = ""
    valor = ""
    if cell_porc_val:
        nums = re.findall(r"\d+(?:[.,]\d+)?", cell_porc_val)
        if len(nums) >= 1:
            porcentaje = nums[0]
        if len(nums) >= 2:
            valor = nums[1]
    return {"impuesto": impuesto, "base": base, "porcentaje": porcentaje, "valor": valor}


def _parse_tabla_retenciones_pdf(page) -> tuple[list[dict], dict]:
    """Lee la tabla de retenciones de la página con `page.extract_tables` y
    devuelve `(filas_impuesto, sustento)`.

    `filas_impuesto`: [{impuesto: IVA|Renta, base, porcentaje, valor}, ...].
    `sustento`: {Comprobante_Sustento, Numero_Sustento, Fecha_Emision_Sustento, Ejercicio_Fiscal}.

    Es mucho más estable que parsear el texto plano porque pdfplumber preserva
    la estructura de columnas; los valores que el `extract_text` mezcla por
    layout aquí quedan en celdas separadas.
    """
    sustento = {
        "Comprobante_Sustento": "",
        "Numero_Sustento": "",
        "Fecha_Emision_Sustento": "",
        "Ejercicio_Fiscal": "",
    }
    filas: list[dict] = []
    try:
        tablas = page.extract_tables() or []
    except Exception:
        return filas, sustento

    for tabla in tablas:
        if not tabla or len(tabla) < 2:
            continue
        header_norm = _norm_pdf_keyword(" ".join((c or "") for c in (tabla[0] or [])))
        if "COMPROBANTE" not in header_norm:
            continue
        for fila in tabla[1:]:
            if not fila:
                continue
            col0 = (fila[0] or "").strip() if len(fila) > 0 else ""
            col1 = (fila[1] or "").strip() if len(fila) > 1 else ""
            col2 = (fila[2] or "").strip() if len(fila) > 2 else ""
            col3 = (fila[3] or "").strip() if len(fila) > 3 else ""
            col4 = (fila[4] or "").strip() if len(fila) > 4 else ""
            col5 = (fila[5] or "").strip() if len(fila) > 5 else ""

            if col0 and not sustento["Comprobante_Sustento"]:
                sustento["Comprobante_Sustento"] = re.sub(r"[ \t]+", " ", col0).strip()
            if col1 and not sustento["Numero_Sustento"]:
                # "0010010000047\n48" → "001001000004748" (concatenamos
                # todos los grupos numéricos sin separador).
                sustento["Numero_Sustento"] = "".join(re.findall(r"\d+", col1))
            if col2 and not sustento["Fecha_Emision_Sustento"]:
                m = re.search(r"(\d{2}/\d{2}/\d{4})", col2)
                if m:
                    sustento["Fecha_Emision_Sustento"] = m.group(1)
            if col3 and not sustento["Ejercicio_Fiscal"]:
                m = re.search(r"(\d{2}/\d{4})", col3)
                if m:
                    sustento["Ejercicio_Fiscal"] = m.group(1)

            fila_imp = _parse_celda_impuesto_retencion(col4, col5)
            if fila_imp:
                filas.append(fila_imp)
        # solo procesamos la primera tabla con header válido
        break
    return filas, sustento


def _extraer_datos_pdf_retencion(pdf_path: Path) -> dict:
    """Extrae los campos del PDF de comprobante de retención.

    Implementación layout-aware (reescrita en 2026-05): usa `extract_words`
    para separar la columna izquierda (datos) de la derecha (labels suffixed
    AUTORIZACIÓN/AMBIENTE/EMISIÓN/CLAVE DE ACCESO), y `extract_tables` para
    parsear la tabla de retenciones. La versión legacy (línea-por-línea +
    regex con tildes) producía cruces de datos: razón social = "No. <numero>",
    nombre comercial = clave de acceso, dirección con label de columna pegado,
    base/porcentaje/valor de IVA tomando dígitos del número de sustento, etc.
    """
    datos = {col: "" for col in RETENCION_REPORT_COLUMNS}
    datos["tipoDocumento"] = "Retencion"

    if pdfplumber is None:
        return datos

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return datos
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
            filas_imp, sustento = _parse_tabla_retenciones_pdf(page)
    except Exception as exc:
        logger.warning("PDF retencion: error abriendo %s: %s", pdf_path, exc)
        return datos

    if not words:
        return datos

    # Líneas visuales con tres vistas: izquierda (x<305), derecha (x>=305), completa.
    SPLIT_X = 305.0
    lineas_words = _agrupar_palabras_visuales(words, y_tol=3.5)
    lineas: list[dict] = []
    for palabras in lineas_words:
        izq = _texto_palabras_rango(palabras, x_max=SPLIT_X)
        der = _texto_palabras_rango(palabras, x_min=SPLIT_X)
        full = _texto_palabras_rango(palabras)
        lineas.append({"izq": izq, "der": der, "full": full})

    der_join = "\n".join(L["der"] for L in lineas)
    full_join = "\n".join(L["full"] for L in lineas)
    der_norm = _norm_pdf_keyword(der_join)
    full_norm = _norm_pdf_keyword(full_join)

    def _limpiar(valor: str) -> str:
        if not valor:
            return ""
        return re.sub(r"[ \t]+", " ", valor).strip()

    # ---- Columna derecha: RUC, número de comprobante, clave, fecha autorización,
    # ambiente, emisión.
    m = re.search(r"R\.?U\.?C\.?\s*:?\s*(\d{10,13})", der_norm)
    if m:
        datos["rucEmisor"] = m.group(1)

    m = re.search(r"No\.\s*(\d{3}-\d{3}-\d{6,9})", der_join)
    if not m:
        m = re.search(r"(\d{3}-\d{3}-\d{6,9})", der_join)
    if m:
        datos["numeroComprobante"] = m.group(1)
        partes = datos["numeroComprobante"].split("-")
        if len(partes) == 3:
            datos["establecimiento"], datos["puntoEmision"], datos["secuencial"] = partes

    m = re.search(r"(\d{49})", der_join)
    if m:
        datos["claveAcceso"] = m.group(1)

    m = re.search(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})", der_join)
    if m:
        datos["fechaAutorizacion"] = m.group(1)

    # Ambiente/Emisión: matcheamos en der_norm (sin tildes) y luego preservamos
    # el texto original (con Ó si la extracción lo decodificó bien).
    if re.search(r"AMBIENTE\s*:?\s*[A-Z]+", der_norm):
        m_orig = re.search(r"AMBIENTE\s*:?\s*(\S+)", der_join, flags=re.IGNORECASE)
        if m_orig:
            datos["ambiente"] = m_orig.group(1).strip().upper()

    m_em = re.search(r"EMISION\s*:?\s*([A-Z]+)", der_norm)
    if m_em:
        datos["emision"] = m_em.group(1)

    # ---- Columna izquierda parte superior: razón social emisor y nombre comercial.
    # Están en las líneas ANTES de "Dirección Matriz:" (de 1 a 2 líneas; suelen
    # repetirse — el reporte de referencia las guarda como duplicado si solo hay 1).
    idx_matriz = None
    for i, L in enumerate(lineas):
        if "DIRECCION MATRIZ" in _norm_pdf_keyword(L["izq"]):
            idx_matriz = i
            break

    razones: list[str] = []
    if idx_matriz is not None:
        for i in range(idx_matriz):
            txt = lineas[i]["izq"]
            if not txt:
                continue
            # Filtra líneas sin al menos 3 letras consecutivas (descarta tokens
            # tipo "LOGO", caracteres sueltos o solo dígitos).
            if not re.search(r"[A-Z]{3,}", _norm_pdf_keyword(txt)):
                continue
            razones.append(_limpiar(txt))

    razones_unicas: list[str] = []
    for r in razones:
        if r not in razones_unicas:
            razones_unicas.append(r)
    if razones_unicas:
        datos["razonSocialEmisor"] = razones_unicas[0]
        datos["nombreComercial"] = razones_unicas[1] if len(razones_unicas) > 1 else razones_unicas[0]
    if not datos["nombreComercial"]:
        datos["nombreComercial"] = "No Disponible"

    # ---- Direcciones matriz/sucursal: solo el texto después de ":" en la línea
    # con el label; ignoramos continuaciones porque el reporte de referencia las
    # ignora.
    def _direccion_de(label_key: str) -> str:
        for L in lineas:
            if label_key in _norm_pdf_keyword(L["izq"]):
                m_dir = re.search(r":\s*(.+)$", L["izq"])
                if m_dir:
                    return _limpiar(m_dir.group(1))
        return ""

    datos["direccionMatrizEmisor"] = _direccion_de("DIRECCION MATRIZ")
    datos["direccionSucursalEmisor"] = _direccion_de("DIRECCION SUCURSAL")

    # ---- Obligado a Llevar Contabilidad ----
    for L in lineas:
        izq_n = _norm_pdf_keyword(L["izq"])
        if "OBLIGADO A LLEVAR CONTABILIDAD" in izq_n:
            m_ob = re.search(r"OBLIGADO A LLEVAR CONTABILIDAD\s+(SI|NO)\b", izq_n)
            if m_ob:
                datos["obligadoContabilidad"] = m_ob.group(1)
            else:
                # El "SI/NO" puede aparecer en otra X de la misma línea.
                m2 = re.search(r"\b(SI|NO)\b", _norm_pdf_keyword(L["full"]))
                if m2:
                    datos["obligadoContabilidad"] = m2.group(1)
            break

    # ---- Número de agente de retención ----
    for L in lineas:
        if "AGENTE DE RETENCION" in _norm_pdf_keyword(L["izq"]):
            full_line_n = _norm_pdf_keyword(L["full"])
            nums = re.findall(r"\b(\d+)\b", full_line_n)
            if nums:
                # El último número es el valor (después de "Resolución No.").
                datos["numeroAgenteRetencion"] = nums[-1]
            break
    if not datos["numeroAgenteRetencion"]:
        datos["numeroAgenteRetencion"] = "No Disponible"

    # ---- Contribuyente Especial ----
    m_ce = re.search(r"CONTRIBUYENTE\s+ESPECIAL[^0-9]*?(\d+)", full_norm)
    if m_ce:
        valor_ce = m_ce.group(1).lstrip("0")
        datos["numeroContribuyenteEspecial"] = valor_ce or "0"
    else:
        datos["numeroContribuyenteEspecial"] = "No Disponible"

    # ---- Bloque del sujeto retenido: razón social, identificación, fecha. ----
    idx_suj = None
    for i, L in enumerate(lineas):
        full_n = _norm_pdf_keyword(L["full"])
        if "RAZON SOCIAL" in full_n and "APELLIDOS" in full_n:
            idx_suj = i
            break

    if idx_suj is not None:
        m_raz = re.search(r"(?i)apellidos\s*:?\s*(.+)$", lineas[idx_suj]["full"])
        if m_raz:
            datos["razonSocialSujetoRetenido"] = _limpiar(m_raz.group(1))

        # Las 2 líneas siguientes contienen "Identificación XXXX" y "Fecha DD/MM/YYYY"
        # (el orden puede invertirse según el agrupamiento por y_tol).
        for offset in (1, 2, 3):
            i = idx_suj + offset
            if i >= len(lineas):
                break
            full_L = lineas[i]["full"]
            full_n = _norm_pdf_keyword(full_L)
            if "IDENTIFICACION" in full_n and not datos["identificacionSujetoRetenido"]:
                m_id = re.search(r"\b(\d{10,13})\b", full_L)
                if m_id:
                    datos["identificacionSujetoRetenido"] = m_id.group(1)
                continue
            if full_n.startswith("FECHA") and not datos["fechaEmision"]:
                m_fe = re.search(r"(\d{2}/\d{2}/\d{4})", full_L)
                if m_fe:
                    datos["fechaEmision"] = m_fe.group(1)

    # ---- Tabla de retenciones (parseada al abrir el PDF) ----
    datos["Comprobante_Sustento"] = sustento["Comprobante_Sustento"]
    datos["Numero_Sustento"] = sustento["Numero_Sustento"]
    datos["Fecha_Emision_Sustento"] = sustento["Fecha_Emision_Sustento"]
    datos["Ejercicio_Fiscal"] = sustento["Ejercicio_Fiscal"]

    iva_items = [f for f in filas_imp if f["impuesto"] == "IVA"]
    renta_items = [f for f in filas_imp if f["impuesto"] == "Renta"]

    def _asignar(items, k_base, k_imp, k_pct, k_val, label):
        if items:
            datos[k_base] = items[0]["base"]
            datos[k_imp] = label
            datos[k_pct] = items[0]["porcentaje"]
            datos[k_val] = items[0]["valor"]

    _asignar(iva_items, "Base_Imponible_Ret_IVA", "Impuesto_Ret_IVA",
             "Porcentaje_Ret_IVA", "Valor_Retenido_IVA", "IVA")
    _asignar(renta_items, "Base_Imponible_Ret_IR", "Impuesto_Ret_IR",
             "Porcentaje_Ret_IR", "Valor_Retenido_IR", "Renta")
    if iva_items:
        _asignar(iva_items, "Base_Imponible_Ret_IVA_1", "Impuesto_Ret_IVA_1",
                 "Porcentaje_Ret_IVA_1", "Valor_Retenido_IVA_1", "IVA")
    if renta_items:
        _asignar(renta_items, "Base_Imponible_Ret_IR_1", "Impuesto_Ret_IR_1",
                 "Porcentaje_Ret_IR_1", "Valor_Retenido_IR_1", "Renta")

    # Si no hay items para un impuesto: marcamos 'No Aplica' en el resumen
    # principal pero dejamos vacíos el detalle _1 (el reporte de referencia
    # los tiene en blanco cuando no aplica).
    if not iva_items:
        datos["Impuesto_Ret_IVA"] = "No Aplica"
        datos["Impuesto_Ret_IVA_1"] = ""
    if not renta_items:
        datos["Impuesto_Ret_IR"] = "No Aplica"
        datos["Impuesto_Ret_IR_1"] = ""

    # ---- Información Adicional ----
    # En estos PDFs el label aparece 2 veces (sección + tabla interna);
    # tomamos lo que viene DESPUÉS del último.
    idx_info = None
    for i, L in enumerate(lineas):
        if "INFORMACION ADICIONAL" in _norm_pdf_keyword(L["full"]):
            idx_info = i
    if idx_info is not None:
        contenido = []
        for L in lineas[idx_info + 1:]:
            txt = (L["full"] or "").strip()
            if txt:
                contenido.append(txt)
        if contenido:
            # Preservamos los saltos de línea — Excel los respeta con wrap-text.
            datos["informacionAdicional"] = "\n".join(contenido)
    if not datos["informacionAdicional"]:
        datos["informacionAdicional"] = "No Disponible"

    # ---- Defaults numéricos para columnas que quedan vacías ----
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
    legacy = _extraer_datos_xml_retencion(xml_path)
    row = _map_retencion_legacy_to_emitidos_sample_row(legacy)
    root, _meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    doc_sustento = root.find(".//docsSustento/docSustento")
    retenciones = doc_sustento.findall("./retenciones/retencion") if doc_sustento is not None else []
    iva_detalles = []
    renta_detalles = []
    for ret in retenciones:
        codigo = _texto_emitidos_retencion(ret.findtext("codigo"))
        detalle = {
            "base": _numero_emitidos_retencion(ret.findtext("baseImponible")),
            "pct": _numero_emitidos_retencion(ret.findtext("porcentajeRetener")),
            "val": _numero_emitidos_retencion(ret.findtext("valorRetenido")),
        }
        if codigo == "2":
            iva_detalles.append(detalle)
        elif codigo == "1":
            renta_detalles.append(detalle)

    if iva_detalles:
        _asignar_resumen_retencion_emitidos(
            row,
            [{"base": d["base"], "imp": "IVA", "pct": d["pct"], "val": d["val"]} for d in iva_detalles],
            base_key="Base_Imponible_Ret_IVA",
            imp_key="Impuesto_Ret_IVA",
            pct_key="Porcentaje_Ret_IVA",
            val_key="Valor_Retenido_IVA",
            det1_base_key="Base_Imponible_Ret_IVA_1",
            det1_imp_key="Impuesto_Ret_IVA_1",
            det1_pct_key="Porcentaje_Ret_IVA_1",
            det1_val_key="Valor_Retenido_IVA_1",
            det2_base_key="Base_Imponible_Ret_IVA_2",
            det2_imp_key="Impuesto_Ret_IVA_2",
            det2_pct_key="Porcentaje_Ret_IVA_2",
            det2_val_key="Valor_Retenido_IVA_2",
            label="IVA",
        )
    if renta_detalles:
        _asignar_resumen_retencion_emitidos(
            row,
            [{"base": d["base"], "imp": "Renta", "pct": d["pct"], "val": d["val"]} for d in renta_detalles],
            base_key="Base_Imponible_Ret_IR",
            imp_key="Impuesto_Ret_IR",
            pct_key="Porcentaje_Ret_IR",
            val_key="Valor_Retenido_IR",
            det1_base_key="Base_Imponible_Ret_IR_1",
            det1_imp_key="Impuesto_Ret_IR_1",
            det1_pct_key="Porcentaje_Ret_IR_1",
            det1_val_key="Valor_Retenido_IR_1",
            det2_base_key="Base_Imponible_Ret_IR_2",
            det2_imp_key="Impuesto_Ret_IR_2",
            det2_pct_key="Porcentaje_Ret_IR_2",
            det2_val_key="Valor_Retenido_IR_2",
            label="Renta",
        )
    return row


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
    return _map_retencion_legacy_to_emitidos_sample_row(_extraer_datos_pdf_retencion(pdf_path))


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
            logger.warning(f"No se pudo extraer por layout visual el PDF '{pdf_path.name}': {err}")

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


def _normalizar_token(texto: str) -> str:
    base = unicodedata.normalize("NFKD", (texto or "")).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", "", base)
