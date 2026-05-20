"""Helpers sobre tipos de comprobante del SRI (Factura, Retención, NC, ND, etc.).

Convierte entre las distintas representaciones que el portal del SRI usa
para un mismo tipo de comprobante:

- **Texto del portal** ("Factura", "Comprobante de retención", "Notas de crédito"…).
- **Slug** (`factura`, `comprobante_de_retencion`) — usado para comparar.
- **Etiqueta canónica** (`Factura`, `Retencion`, `NotaCredito`) + **orden numérico**
  para los nombres de carpeta y reportes (`01_Factura`, `06_Retencion`).
- **Nombre visible** ("Factura", "Comprobante de Retencion", "Nota de Credito").

Sin dependencias de Playwright ni I/O. Originalmente vivía dentro de
`robot/downloader.py`; se extrajo en la Sub-fase 2a del refactor.
"""
from __future__ import annotations

import re
import unicodedata


# --------------------------------------------------------------------------- #
# Normalización a slug / nombre de carpeta
# --------------------------------------------------------------------------- #
def _nombre_carpeta_tipo(tipo: str) -> str:
    """Devuelve un nombre de carpeta válido a partir de un tipo de comprobante.

    Quita acentos, reemplaza no-alfanuméricos por `_` y trimea bordes. Si el
    resultado queda vacío, devuelve `"Otros"`.
    """
    base = unicodedata.normalize("NFKD", (tipo or "")).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_")
    return base or "Otros"


def _slug_tipo(tipo: str) -> str:
    """Versión en minúsculas de `_nombre_carpeta_tipo`, apta para comparaciones."""
    return _nombre_carpeta_tipo(tipo).lower()


def _normalizar_label_simple(texto_label: str) -> str:
    """Normaliza un label visible: sin acentos, mayúsculas, espacios colapsados."""
    base = unicodedata.normalize("NFKD", texto_label or "").encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9 ]+", " ", base).upper()
    base = re.sub(r"\s+", " ", base).strip()
    return base


def _normalizar_tipo_clave(texto: str) -> str:
    """Como `_slug_tipo`, pero devuelve el resultado directamente en lowercase
    sin pasar por la convención de carpeta (no devuelve 'Otros' si está vacío).
    """
    base = unicodedata.normalize("NFKD", (texto or "")).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_").lower()
    return base


def _formatear_label(texto: str) -> str:
    """Capitaliza un slug separado por `_` (ej: `nota_de_credito` → `Nota_De_Credito`)."""
    if not texto:
        return "Documentos"
    partes = [fragment.capitalize() for fragment in texto.split("_") if fragment]
    return "_".join(partes) if partes else (texto or "Documentos")


# --------------------------------------------------------------------------- #
# Detección de tipos específicos
# --------------------------------------------------------------------------- #
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


# --------------------------------------------------------------------------- #
# Mapeo canónico tipo → (orden, etiqueta)
# --------------------------------------------------------------------------- #
# El orden determina el prefijo numérico que se usa en nombres de carpeta y
# de reportes Excel (`01_Factura`, `02_Liquidacion_de_Compra`, etc.).
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


# --------------------------------------------------------------------------- #
# Resolución de etiquetas y nombres visibles
# --------------------------------------------------------------------------- #
def _nombre_carpeta_tipo_visible(tipo_texto: str) -> str:
    """Devuelve el nombre legible para el usuario de un tipo de comprobante."""
    clave = _normalizar_tipo_clave(tipo_texto)
    if clave in {
        "retencion", "retenciones",
        "comprobante_de_retencion", "comprobantes_de_retencion",
        "comprobante_de_retencion_venta",
    }:
        return "Comprobante de Retencion"
    if clave in {"factura", "facturas"}:
        return "Factura"
    if clave in {
        "liquidacion_de_compra",
        "liquidacion_de_compra_de_bienes_y_prestacion_de_servicios",
    }:
        return "Liquidacion de Compra"
    if clave in {"nota_de_credito", "notas_de_credito", "nota_credito", "notas_credito"}:
        return "Nota de Credito"
    if clave in {"nota_de_debito", "notas_de_debito", "nota_debito", "notas_debito"}:
        return "Nota de Debito"
    if clave in {"guia_de_remision", "guias_de_remision", "guia_remision"}:
        return "Guia de Remision"
    return _formatear_label(_nombre_carpeta_tipo(tipo_texto)).replace("_", " ")


def _resolver_tipo_label(tipo_texto: str) -> tuple[int, str]:
    """Resuelve un texto de tipo al par (orden, etiqueta_canonica).

    Para tipos desconocidos devuelve `(99, label_sanitizado)` — el 99 indica
    que cae al final del orden estándar.
    """
    clave = _normalizar_tipo_clave(tipo_texto)
    if clave in TIPO_LABEL_MAP:
        return TIPO_LABEL_MAP[clave]
    # Tolerar plurales que el portal a veces usa.
    if clave.endswith("s"):
        clave_singular = clave.rstrip("s")
        if clave_singular in TIPO_LABEL_MAP:
            return TIPO_LABEL_MAP[clave_singular]
    label_sanitizado = _nombre_carpeta_tipo(tipo_texto)
    label = _formatear_label(label_sanitizado)
    return 99, label


def _coincide_tipo_documental(tipo_esperado: str, tipo_detectado: str) -> bool:
    """True si dos textos de tipo refieren al mismo comprobante.

    Si alguno está vacío, devuelve True (no rechaza por falta de info).
    Para tipos desconocidos cae a comparar slugs.
    """
    if not tipo_esperado or not tipo_detectado:
        return True
    orden_esperado, _ = _resolver_tipo_label(tipo_esperado)
    orden_detectado, _ = _resolver_tipo_label(tipo_detectado)
    if orden_esperado == 99 or orden_detectado == 99:
        return _slug_tipo(tipo_esperado) == _slug_tipo(tipo_detectado)
    return orden_esperado == orden_detectado


def _prefijo_tipo(tipo_texto: str) -> tuple[int, str, str]:
    """Devuelve `(orden, etiqueta, prefijo_zero_padded)` para nombres de carpeta/reporte.

    Ejemplo: `_prefijo_tipo("Factura")` → `(1, "Factura", "01_Factura")`.
    """
    orden, etiqueta = _resolver_tipo_label(tipo_texto)
    prefijo = f"{orden:02d}_{etiqueta}"
    return orden, etiqueta, prefijo
