"""Helpers de formato y normalización de datos para reportes.

Dos grupos de funciones:

- **Parseo de números** (`_parse_decimal`): convierte un texto a `float`
  tolerando los formatos que aparecen en el portal SRI (separador decimal
  con coma, miles con punto, espacios y NBSP).
- **Filas por defecto** (`_*_default_row`): devuelven un dict con las
  columnas del reporte correspondiente prerellenadas con `""`, `0` o
  `"No Disponible"` según la columna. Las funciones de generación de
  reporte usan estos defaults para columnas que no aparecieron en el JSON
  del SRI.

Originalmente vivían dentro de `robot/downloader.py`; extraídas en la
Sub-fase 2c-ii-a del refactor.
"""
from __future__ import annotations

from typing import Optional

from robot.report_columns import (
    EMITIDOS_FACTURA_REPORT_COLUMNS,
    EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS,
    EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS,
    EMITIDOS_RETENCION_NUMERIC_COLUMNS,
    EMITIDOS_RETENCION_REPORT_COLUMNS,
)


# --------------------------------------------------------------------------- #
# Parseo numérico
# --------------------------------------------------------------------------- #
def _parse_decimal(texto: str) -> Optional[float]:
    """Convierte `texto` a float intentando varias normalizaciones.

    El SRI mezcla formatos en distintas pantallas: a veces `1.234,56`,
    a veces `1234.56`, a veces con espacios o NBSP. Esta función prueba
    el texto tal cual, luego sin puntos (asumiendo separador de miles) y
    luego sin comas. Devuelve `None` si ninguno parsea.
    """
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


# --------------------------------------------------------------------------- #
# Filas por defecto para reportes emitidos
# --------------------------------------------------------------------------- #
def _emitidos_retencion_default_row() -> dict:
    """Fila con columnas prerellenadas para reporte de retenciones emitidas."""
    row = {col: "" for col in EMITIDOS_RETENCION_REPORT_COLUMNS}
    for col in (
        "nombreComercial",
        "numeroContribuyenteEspecial",
        "numeroAgenteRetencion",
        "informacionAdicional",
    ):
        row[col] = "No Disponible"
    for col in EMITIDOS_RETENCION_NUMERIC_COLUMNS:
        row[col] = 0
    row["Impuesto_Ret_IVA"] = "No Aplica"
    row["Impuesto_Ret_IR"] = "No Aplica"
    row["tipoDocumento"] = "Retencion"
    return row


def _nota_credito_emitidos_default_row() -> dict:
    """Fila con columnas prerellenadas para reporte de notas de crédito emitidas."""
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
    """Fila con columnas prerellenadas para reporte de facturas emitidas."""
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
    """Fila con columnas prerellenadas para reporte de notas de débito emitidas."""
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
