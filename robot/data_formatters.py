"""Helpers de formato y normalización de datos para reportes.

Grupos de funciones:

- **Parseo de números / fechas** (`_parse_decimal`, `_parse_datetime_local`):
  convierten texto a `float` / `datetime` tolerando los formatos que aparecen
  en el portal SRI (separador decimal con coma, miles con punto, espacios, NBSP).
- **Filas por defecto** (`_*_default_row`): devuelven un dict con las
  columnas del reporte correspondiente prerellenadas con `""`, `0` o
  `"No Disponible"` según la columna.
- **Formato de campos de Emitidos** (`_texto_emitidos_retencion`,
  `_numero_emitidos_retencion`, `_label_*`, `_normalizar_*`): normalizan
  los valores crudos del XML de comprobantes emitidos a los textos finales
  del reporte (etiquetas de ambiente, forma de pago, tipo de identificación,
  resúmenes de retención IVA/Renta).

Originalmente vivían dentro de `robot/downloader.py`; extraídas en las
Sub-fases 2c-ii-a y 3a-i del refactor.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Optional

from robot.report_columns import (
    EMITIDOS_FACTURA_REPORT_COLUMNS,
    EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS,
    EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL,
    EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS,
    EMITIDOS_RETENCION_AMBIENTE_LABEL,
    EMITIDOS_RETENCION_FORMA_PAGO_LABEL,
    EMITIDOS_RETENCION_NUMERIC_COLUMNS,
    EMITIDOS_RETENCION_REPORT_COLUMNS,
    EMITIDOS_RETENCION_TIPO_EMISION_LABEL,
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


def _parse_datetime_local(texto: str) -> Optional[datetime]:
    """Parsea fechas en formatos del portal SRI (DD/MM/YYYY [HH:MM[:SS]]).

    Tolera espacios múltiples internos. Devuelve None si no matchea ninguno
    de los tres formatos esperados.
    """
    bruto = (texto or "").strip()
    if not bruto:
        return None
    bruto = re.sub(r"\s+", " ", bruto)
    formatos = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y")
    for fmt in formatos:
        try:
            return datetime.strptime(bruto, fmt)
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


# --------------------------------------------------------------------------- #
# Formato de campos de comprobantes Emitidos
# --------------------------------------------------------------------------- #
def _texto_emitidos_retencion(valor, default: str = "") -> str:
    """Normaliza un valor a texto: colapsa espacios, devuelve `default` si vacío."""
    if valor is None:
        return default
    texto = re.sub(r"\s+", " ", str(valor).strip())
    return texto or default


def _texto_emitidos_retencion_multilinea(valor, default: str = "") -> str:
    """Como `_texto_emitidos_retencion` pero preserva los saltos de línea.

    Sólo colapsa whitespace horizontal (espacios/tabs) dentro de cada línea, sin
    juntar líneas con espacio. Útil para `informacionAdicional`, donde la
    referencia del SRI guarda cada bloque (eMail/Sistema/Telefono/Correo) en
    una línea propia."""
    if valor is None:
        return default
    texto = str(valor).strip()
    if not texto:
        return default
    # Normaliza \r\n y \r a \n; colapsa runs de espacios/tabs pero respeta \n.
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    # Quitamos espacios al inicio/fin de cada línea y descartamos líneas vacías
    # consecutivas (deja una sola en blanco como separador máximo).
    lineas = [ln.strip() for ln in texto.split("\n")]
    # Colapsa blancos consecutivos
    salida = []
    blanco_previo = False
    for ln in lineas:
        if ln:
            salida.append(ln); blanco_previo = False
        elif not blanco_previo:
            salida.append("")
            blanco_previo = True
    while salida and not salida[-1]:
        salida.pop()
    return "\n".join(salida) or default


def _texto_emitidos_retencion_na(valor) -> str:
    """Como `_texto_emitidos_retencion` pero con default 'No Disponible'."""
    return _texto_emitidos_retencion(valor, "No Disponible")


def _numero_emitidos_retencion(valor, default=0):
    """Convierte `valor` a número; devuelve `default` si no se puede parsear."""
    if valor in ("", None):
        return default
    if isinstance(valor, (int, float)):
        return valor
    parsed = _parse_decimal(str(valor))
    return parsed if parsed is not None else default


def _normalizar_ambiente_retencion_emitidos(valor) -> str:
    """Normaliza el ambiente a 'PRODUCCIÓN' / 'PRUEBAS' (o el texto original)."""
    texto = _texto_emitidos_retencion(valor)
    if not texto:
        return ""
    valor_norm = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").upper()
    if "PRODUCCION" in valor_norm:
        return "PRODUCCIÓN"
    if "PRUEBA" in valor_norm:
        return "PRUEBAS"
    return texto


def _normalizar_emision_retencion_emitidos(valor) -> str:
    """Normaliza el tipo de emisión a 'NORMAL' / 'CONTINGENCIA'."""
    texto = _texto_emitidos_retencion(valor)
    if not texto:
        return ""
    valor_norm = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").upper()
    if "NORMAL" in valor_norm:
        return "NORMAL"
    if "INDISPONIBILIDAD" in valor_norm or "CONTINGENCIA" in valor_norm:
        return "CONTINGENCIA"
    return texto


def _formatear_fecha_autorizacion_retencion_emitidos(valor) -> str:
    """Convierte la fecha de autorización a 'DD/MM/YYYY HH:MM:SS'.

    Tolera varios formatos de entrada (ISO con/sin zona, formato local).
    Si no matchea ninguno, devuelve el texto original.
    """
    texto = _texto_emitidos_retencion(valor)
    if not texto:
        return ""
    for fmt_in in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            dt = datetime.strptime(texto, fmt_in)
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            continue
    match = re.search(r"(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})", texto)
    if match:
        try:
            dt = datetime.strptime(f"{match.group(1)} {match.group(2)}", "%Y-%m-%d %H:%M:%S")
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            pass
    return texto


def _detalle_retencion_emitidos(legacy: dict, prefijo: str) -> list[dict]:
    """Extrae los detalles de retención (IVA o IR) de un dict legacy.

    `prefijo` es "IVA" o cualquier otro (interpretado como Renta). Devuelve
    una lista de dicts {base, imp, pct, val}, sin entradas vacías ni duplicados
    consecutivos.
    """
    if prefijo == "IVA":
        label = "IVA"
        keys = [
            ("Base_Imponible_Ret_IVA", "Impuesto_Ret_IVA", "Porcentaje_Ret_IVA", "Valor_Retenido_IVA"),
            ("Base_Imponible_Ret_IVA_1", "Impuesto_Ret_IVA_1", "Porcentaje_Ret_IVA_1", "Valor_Retenido_IVA_1"),
        ]
    else:
        label = "Renta"
        keys = [
            ("Base_Imponible_Ret_IR", "Impuesto_Ret_IR", "Porcentaje_Ret_IR", "Valor_Retenido_IR"),
            ("Base_Imponible_Ret_IR_1", "Impuesto_Ret_IR_1", "Porcentaje_Ret_IR_1", "Valor_Retenido_IR_1"),
        ]

    detalles: list[dict] = []
    for base_key, imp_key, pct_key, val_key in keys:
        base = _numero_emitidos_retencion(legacy.get(base_key))
        pct = _numero_emitidos_retencion(legacy.get(pct_key))
        val = _numero_emitidos_retencion(legacy.get(val_key))
        imp = _texto_emitidos_retencion(legacy.get(imp_key))
        if imp.upper() == "NO APLICA":
            continue
        if not any([base, pct, val]) and not imp:
            continue
        detalle = {"base": base, "imp": label, "pct": pct, "val": val}
        if not detalles or detalle != detalles[-1]:
            detalles.append(detalle)
    return detalles


def _asignar_resumen_retencion_emitidos(
    row: dict,
    detalles: list[dict],
    *,
    base_key: str,
    imp_key: str,
    pct_key: str,
    val_key: str,
    det1_base_key: str,
    det1_imp_key: str,
    det1_pct_key: str,
    det1_val_key: str,
    det2_base_key: str,
    det2_imp_key: str,
    det2_pct_key: str,
    det2_val_key: str,
    label: str,
    no_aplica_si_vacio: bool = True,
) -> None:
    """Vuelca los `detalles` de retención en las columnas resumen de `row`.

    Maneja 0, 1 o varios detalles; con varios marca el impuesto como 'Varios'
    y suma bases/valores.
    """
    if not detalles:
        row[base_key] = 0
        row[pct_key] = 0
        row[val_key] = 0
        row[det1_base_key] = 0
        row[det1_pct_key] = 0
        row[det1_val_key] = 0
        row[det2_base_key] = 0
        row[det2_pct_key] = 0
        row[det2_val_key] = 0
        row[imp_key] = "No Aplica" if no_aplica_si_vacio else ""
        row[det1_imp_key] = ""
        row[det2_imp_key] = ""
        return

    if len(detalles) == 1:
        base = detalles[0]["base"]
        pct = detalles[0]["pct"]
        val = detalles[0]["val"]
        row[base_key] = base
        row[imp_key] = label
        row[pct_key] = pct
        row[val_key] = val
        row[det1_base_key] = base
        row[det1_imp_key] = label
        row[det1_pct_key] = pct
        row[det1_val_key] = val
        row[det2_base_key] = 0
        row[det2_imp_key] = ""
        row[det2_pct_key] = 0
        row[det2_val_key] = 0
        return

    row[base_key] = sum(d["base"] for d in detalles)
    row[imp_key] = "Varios"
    row[pct_key] = 0
    row[val_key] = sum(d["val"] for d in detalles)
    first = detalles[0]
    second = detalles[1]
    row[det1_base_key] = first["base"]
    row[det1_imp_key] = label
    row[det1_pct_key] = first["pct"]
    row[det1_val_key] = first["val"]
    row[det2_base_key] = second["base"]
    row[det2_imp_key] = label
    row[det2_pct_key] = second["pct"]
    row[det2_val_key] = second["val"]


def _map_retencion_legacy_to_emitidos_sample_row(legacy: dict | None) -> dict:
    """Convierte un dict de retención en formato legacy a una fila de reporte
    de retenciones emitidas, normalizando todos los campos.
    """
    row = _emitidos_retencion_default_row()
    if not isinstance(legacy, dict):
        return row

    row["rucEmisor"] = _texto_emitidos_retencion(legacy.get("rucEmisor"))
    row["razonSocialEmisor"] = _texto_emitidos_retencion(legacy.get("razonSocialEmisor"))
    row["nombreComercial"] = _texto_emitidos_retencion(legacy.get("nombreComercial"), "No Disponible")
    row["direccionMatrizEmisor"] = _texto_emitidos_retencion(legacy.get("direccionMatrizEmisor"))
    row["direccionSucursalEmisor"] = _texto_emitidos_retencion(legacy.get("direccionSucursalEmisor"))
    row["obligadoContabilidad"] = _texto_emitidos_retencion(legacy.get("obligadoContabilidad"))
    row["numeroContribuyenteEspecial"] = _texto_emitidos_retencion(
        legacy.get("numeroContribuyenteEspecial"), "No Disponible"
    )
    row["numeroAgenteRetencion"] = _texto_emitidos_retencion(
        legacy.get("numeroAgenteRetencion"), "No Disponible"
    )
    row["fechaAutorizacion"] = _formatear_fecha_autorizacion_retencion_emitidos(
        legacy.get("fechaAutorizacion")
    )
    row["ambiente"] = _normalizar_ambiente_retencion_emitidos(legacy.get("ambiente"))
    row["emision"] = _normalizar_emision_retencion_emitidos(legacy.get("emision"))
    row["numeroComprobante"] = _texto_emitidos_retencion(legacy.get("numeroComprobante"))
    row["establecimiento"] = _texto_emitidos_retencion(legacy.get("establecimiento"))
    row["puntoEmision"] = _texto_emitidos_retencion(legacy.get("puntoEmision"))
    row["secuencial"] = _texto_emitidos_retencion(legacy.get("secuencial"))
    row["fechaEmision"] = _texto_emitidos_retencion(legacy.get("fechaEmision"))
    row["razonSocialSujetoRetenido"] = _texto_emitidos_retencion(
        legacy.get("razonSocialSujetoRetenido")
    )
    row["identificacionSujetoRetenido"] = _texto_emitidos_retencion(
        legacy.get("identificacionSujetoRetenido")
    )
    row["claveAcceso"] = _texto_emitidos_retencion(legacy.get("claveAcceso"))
    row["Comprobante_Sustento"] = _texto_emitidos_retencion(legacy.get("Comprobante_Sustento"))
    row["Numero_Sustento"] = _texto_emitidos_retencion(legacy.get("Numero_Sustento")).replace("-", "")
    row["Fecha_Emision_Sustento"] = _texto_emitidos_retencion(legacy.get("Fecha_Emision_Sustento"))
    row["Ejercicio_Fiscal"] = _texto_emitidos_retencion(legacy.get("Ejercicio_Fiscal"))
    row["informacionAdicional"] = _texto_emitidos_retencion_multilinea(
        legacy.get("informacionAdicional"), "No Disponible"
    )
    row["tipoDocumento"] = _texto_emitidos_retencion(legacy.get("tipoDocumento"), "Retencion")

    iva_detalles = _detalle_retencion_emitidos(legacy, "IVA")
    renta_detalles = _detalle_retencion_emitidos(legacy, "IR")
    _asignar_resumen_retencion_emitidos(
        row,
        iva_detalles,
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
    _asignar_resumen_retencion_emitidos(
        row,
        renta_detalles,
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


def _label_tipo_ident_emitidos_nota_credito(valor: str) -> str:
    """Mapea un código de tipo de identificación a su etiqueta legible."""
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL:
        return EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL[valor]
    return valor or "No Disponible"


def _label_ambiente_emitidos_retencion(valor: str) -> str:
    """Mapea un código/texto de ambiente a su etiqueta legible."""
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_RETENCION_AMBIENTE_LABEL:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL[valor]
    valor_norm = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii").upper()
    if "PRODUCCION" in valor_norm:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL["2"]
    if "PRUEBA" in valor_norm:
        return EMITIDOS_RETENCION_AMBIENTE_LABEL["1"]
    return valor


def _label_emision_emitidos_retencion(valor: str) -> str:
    """Mapea un código/texto de tipo de emisión a su etiqueta legible."""
    valor = _texto_emitidos_retencion(valor)
    if valor in EMITIDOS_RETENCION_TIPO_EMISION_LABEL:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL[valor]
    valor_norm = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii").upper()
    if "NORMAL" in valor_norm:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL["1"]
    if "INDISPONIBILIDAD" in valor_norm or "CONTINGENCIA" in valor_norm:
        return EMITIDOS_RETENCION_TIPO_EMISION_LABEL["2"]
    return valor


def _label_forma_pago_emitidos_retencion(valor: str) -> str:
    """Mapea un código de forma de pago a su etiqueta legible."""
    valor = _texto_emitidos_retencion(valor)
    if not valor:
        return "No Disponible"
    if valor in EMITIDOS_RETENCION_FORMA_PAGO_LABEL:
        return EMITIDOS_RETENCION_FORMA_PAGO_LABEL[valor]
    return valor
