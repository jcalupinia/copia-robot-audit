"""Helpers de bajo nivel para parsing de PDFs de comprobantes del SRI.

Tres familias de utilidades:

- **Lectura de PDF**: `_leer_texto_pdf` (texto plano con pdfplumber),
  `_normalizar_texto_pdf`, `_es_archivo_pdf`.
- **Extraccion por regex/texto**: `_extraer_regex`, `_extraer_monto`,
  `_extraer_forma_pago`, `_extraer_seccion`, `_extraer_tipo_documento`.
- **Extraccion por layout** (posiciones x/y de pdfplumber):
  `_extraer_lineas_layout_pdf` y los helpers `_texto_linea_layout`,
  `_buscar_indice_linea_layout*`, `_siguiente_linea_layout_no_vacia`,
  `_fecha_hora_pdf_a_iso`.

Las funciones principales de extraccion (`_extraer_datos_pdf*`) siguen en
`robot/downloader.py` y se apoyan en estos helpers; se moveran en una
sub-fase posterior.

Extraido en la Sub-fase 3b-A del refactor.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
except Exception:  # pragma: no cover - entorno sin pdfplumber
    pdfplumber = None

from robot.comprobante_types import _normalizar_label_simple


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
