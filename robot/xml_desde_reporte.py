"""Reconstruye XML de comprobantes EMITIDOS a partir del reporte del robot.

Para que sirve
--------------
El WS del SRI deja de servir los comprobantes emitidos al mes de autorizados.
Pasado ese plazo, si el contribuyente perdio los XML, lo unico que queda es el
reporte que genero el robot. Este modulo hace el camino inverso: reporte -> XML.

Lo que NO hace, y no puede hacer
--------------------------------
El XML resultante **no lleva firma digital (XAdES-BES)**. Esa firma se calcula
con la clave privada del contribuyente sobre los bytes exactos del original: no
es un campo que se pueda rellenar, y cualquier diferencia de un byte da otra
firma. Sin ella el archivo NO es un comprobante electronico: no sustenta credito
tributario ni sirve ante una fiscalizacion.

Sirve para cargar a un sistema contable, conciliar y auditar. Cada XML sale
marcado como reconstruido en `infoAdicional` para que nunca se confunda con un
original.

Alcance
-------
Solo EMITIDOS. Factura esta validada contra datos reales; nota de credito y
nota de debito comparten 37 de sus 38 columnas con factura y estan
implementadas, pero todavia sin contrastar contra un reporte real de ese tipo.
Retencion queda afuera a proposito: su reporte sale sin los campos de
identificacion porque el extractor de retenciones esta roto, y reconstruir
sobre eso seria inventar.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Callable, Iterable, Optional
from xml.etree import ElementTree as ET

import pandas as pd

from robot._logging import get_logger

logger = get_logger(__name__)

try:  # La validacion contra el XSD es opcional: sin lxml se genera igual.
    from lxml import etree as _lxml_etree
except Exception:  # pragma: no cover - depende del entorno
    _lxml_etree = None


# =============================================================================
# Tablas del SRI
# =============================================================================

# Tabla 17 de la ficha tecnica: codigo del porcentaje de IVA.
_CODIGO_PORCENTAJE_IVA = {
    "0": "0",
    "5": "5",
    "12": "2",
    "13": "10",
    "14": "3",
    "15": "4",
}
_COD_IVA = "2"          # Tabla 16: codigo del impuesto IVA
_COD_NO_OBJETO = "6"
_COD_EXENTO = "7"

CARPETA_XSD = Path(__file__).resolve().parent.parent / "FichaTecnica"

# codDoc -> (nombre, subcarpeta del XSD, archivo)
TIPOS = {
    "01": ("factura", "XML y XSD Factura/XML y XSD Factura", "factura_V1.1.0.xsd"),
    "04": ("nota de credito", "XML y XSD Nota de Crédito/XML y XSD Nota de Crédito", "NotaCredito_V1.1.0.xsd"),
    "05": ("nota de debito", "XML y XSD Nota de Débito/XML y XSD Nota de Débito", "NotaDebito_V1.0.0.xsd"),
}

# Tipos que el modulo reconoce pero todavia no reconstruye, con el motivo.
TIPOS_NO_SOPORTADOS = {
    "07": (
        "comprobante de retencion",
        "el reporte de retenciones emitidas sale sin rucEmisor, claveAcceso ni "
        "fechaEmision porque el extractor de retenciones esta roto: no hay de "
        "donde reconstruir",
    ),
    "06": ("guia de remision", "el robot no genera reporte de guias en Emitidos"),
    "03": ("liquidacion de compra", "la ficha tecnica del repo no trae su XSD"),
}


# =============================================================================
# Lectura de los valores del reporte
# =============================================================================


def _texto(valor) -> str:
    """Normaliza una celda a texto limpio. Los vacios del reporte tambien."""
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in {"nan", "none", "no disponible", "n/a", "sd"}:
        return ""
    return texto


def _numero(valor, defecto: float = 0.0) -> float:
    texto = _texto(valor)
    if not texto:
        return defecto
    texto = re.sub(r"[^\d,.\-]", "", texto)
    # "1.234,56" -> "1234.56";  "1,234.56" -> "1234.56"
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".") if texto.rfind(",") > texto.rfind(".") else texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return defecto


def _dec(valor, decimales: int = 2) -> str:
    """Formato de importe del SRI: punto decimal y cantidad fija de decimales."""
    return f"{_numero(valor):.{decimales}f}"


def _codigo_de_etiqueta(valor) -> str:
    """De "2 - Produccion" saca "2"; de "04 - RUC" saca "04"."""
    texto = _texto(valor)
    match = re.match(r"\s*(\d{1,2})\s*-", texto)
    if match:
        return match.group(1)
    return texto if texto.isdigit() else ""


def _sin_acentos(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


def _columna(fila: pd.Series, *nombres: str) -> str:
    """Busca una columna tolerando acentos y mayusculas."""
    indice = {_sin_acentos(str(c)).strip().lower(): c for c in fila.index}
    for nombre in nombres:
        clave = _sin_acentos(nombre).strip().lower()
        if clave in indice:
            return _texto(fila[indice[clave]])
    return ""


def descomponer_clave_acceso(clave: str) -> dict:
    """Los 49 digitos de la clave traen casi toda la infoTributaria.

    Es la fuente mas confiable que hay: no depende de que el extractor haya
    leido bien una etiqueta del RIDE.
    """
    clave = re.sub(r"\D", "", clave or "")
    if len(clave) != 49:
        return {}
    return {
        "fechaEmision": f"{clave[0:2]}/{clave[2:4]}/{clave[4:8]}",
        "codDoc": clave[8:10],
        "ruc": clave[10:23],
        "ambiente": clave[23:24],
        "estab": clave[24:27],
        "ptoEmi": clave[27:30],
        "secuencial": clave[30:39],
        "codigoNumerico": clave[39:47],
        "tipoEmision": clave[47:48],
        "digitoVerificador": clave[48:49],
    }


_PATRON_ITEM = re.compile(
    r"C.?digo:\s*(?P<codigo>.*?),\s*"
    r"Aux:\s*(?P<aux>.*?),\s*"
    r"Cant:\s*(?P<cantidad>[\d.,]+),\s*"
    r"Desc:\s*(?P<descripcion>.*?),\s*"
    r"P\.Unit:\s*(?P<precio>[\d.,]+),.*?"
    r"Descuento:\s*(?P<descuento>[\d.,]+),\s*"
    r"P\.Total:\s*(?P<total>[\d.,]+)",
    re.DOTALL,
)


def parsear_descripciones(texto: str) -> list[dict]:
    """Reconstruye el detalle linea por linea desde la columna Descripciones.

    El reporte guarda cada item como "Codigo: X, Aux: Y, Cant: N, Desc: ..." y
    los separa con "|". Trae todo lo que el XSD pide para `<detalle>` salvo el
    impuesto por linea, que se deduce aparte.
    """
    items = []
    for parte in (texto or "").split("|"):
        parte = parte.strip()
        if not parte:
            continue
        match = _PATRON_ITEM.search(parte)
        if not match:
            continue
        datos = match.groupdict()
        items.append(
            {
                "codigoPrincipal": _texto(datos["codigo"]) or "SIN-CODIGO",
                "codigoAuxiliar": _texto(datos["aux"]),
                "descripcion": _texto(datos["descripcion"]),
                "cantidad": _numero(datos["cantidad"]),
                "precioUnitario": _numero(datos["precio"]),
                "descuento": _numero(datos["descuento"]),
                "precioTotalSinImpuesto": _numero(datos["total"]),
            }
        )
    return items


def parsear_formas_pago(texto: str) -> list[dict]:
    """De "19 - TARJETA DE CREDITO - 26.00" saca (codigo, total)."""
    pagos = []
    for parte in re.split(r"[|;]", texto or ""):
        parte = parte.strip()
        if not parte:
            continue
        match = re.match(r"\s*(\d{2})\s*-\s*(.+?)\s*-\s*([\d.,]+)\s*$", parte)
        if match:
            pagos.append({"formaPago": match.group(1), "total": _numero(match.group(3))})
    return pagos


def parsear_campos_adicionales(texto: str) -> list[tuple[str, str]]:
    """De "Telefono: 099...; Correo: x@y.z" saca los pares nombre/valor."""
    campos = []
    for parte in re.split(r"[;|]", texto or ""):
        parte = parte.strip()
        if not parte or ":" not in parte:
            continue
        nombre, _, valor = parte.partition(":")
        nombre, valor = nombre.strip(), valor.strip()
        if nombre and valor:
            campos.append((nombre, valor))
    return campos


_FORMA_IDENTIFICACION = re.compile(r"[A-Za-z0-9\-]{5,20}")
_CONSUMIDOR_FINAL = "9" * 13


def identificacion_comprador(fila: pd.Series) -> tuple[str, str]:
    """Devuelve (identificacion, origen) blindando contra reportes viejos.

    Los reportes generados antes del arreglo del extractor traen aca el texto
    de otra linea del RIDE ("Fecha 06/03/2026 Placa / Matricula: Guia"). El XSD
    no lo rechaza -- el campo admite cualquier cadena -- asi que la unica
    defensa es mirar la forma antes de escribirlo.
    """
    valor = _columna(fila, "Identificación Comprador", "identificacionComprador")
    if _FORMA_IDENTIFICACION.fullmatch(valor):
        return valor, DIRECTO
    # Consumidor final tiene identificacion conocida, asi que se recupera.
    razon = _sin_acentos(_columna(fila, "Razón Social Comprador")).upper()
    if "CONSUMIDOR FINAL" in razon:
        return _CONSUMIDOR_FINAL, DEDUCIDO
    return "", FALTANTE


def codigo_porcentaje_iva(tarifa: str) -> str:
    """De "15%" saca el codigo 4 de la tabla 17."""
    texto = _sin_acentos(_texto(tarifa)).upper()
    if "NO OBJETO" in texto:
        return _COD_NO_OBJETO
    if "EXENT" in texto:
        return _COD_EXENTO
    match = re.search(r"(\d{1,2})\s*%", texto)
    if match:
        return _CODIGO_PORCENTAJE_IVA.get(match.group(1), "")
    return ""


def tarifas_declaradas(tarifa: str) -> list[str]:
    """Cuantas tarifas distintas declara el comprobante."""
    return [t for t in re.split(r"[,;|]", _texto(tarifa)) if t.strip()]


# =============================================================================
# Cobertura: de donde salio cada campo
# =============================================================================

DIRECTO = "directo"      # copiado tal cual del reporte
DEDUCIDO = "deducido"    # calculado a partir de otros campos
FALTANTE = "faltante"    # el XSD lo pide y no hay de donde sacarlo


class Cobertura:
    """Registra el origen de cada campo para el informe final."""

    def __init__(self) -> None:
        self.campos: dict[str, str] = {}

    def anota(self, campo: str, origen: str) -> None:
        self.campos[campo] = origen

    def cuenta(self, origen: str) -> int:
        return sum(1 for v in self.campos.values() if v == origen)

    def lista(self, origen: str) -> str:
        return ", ".join(sorted(c for c, v in self.campos.items() if v == origen))


def _sub(padre: ET.Element, etiqueta: str, valor, cobertura: Cobertura,
         origen: str = DIRECTO, obligatorio: bool = False) -> Optional[ET.Element]:
    """Agrega un subelemento y anota de donde salio su valor."""
    texto = _texto(valor)
    if not texto:
        if obligatorio:
            cobertura.anota(etiqueta, FALTANTE)
        return None
    nodo = ET.SubElement(padre, etiqueta)
    nodo.text = texto
    cobertura.anota(etiqueta, origen)
    return nodo


# =============================================================================
# Constructores
# =============================================================================


def _info_tributaria(raiz: ET.Element, fila: pd.Series, clave: dict,
                     cobertura: Cobertura) -> None:
    info = ET.SubElement(raiz, "infoTributaria")
    ambiente = _codigo_de_etiqueta(_columna(fila, "Ambiente")) or clave.get("ambiente", "")
    emision = _codigo_de_etiqueta(_columna(fila, "Tipo Emisión", "Tipo Emision")) or clave.get("tipoEmision", "")

    _sub(info, "ambiente", ambiente, cobertura, DIRECTO, True)
    _sub(info, "tipoEmision", emision, cobertura, DIRECTO, True)
    _sub(info, "razonSocial", _columna(fila, "Razón Social Emisor", "razonSocialEmisor"), cobertura, DIRECTO, True)
    _sub(info, "nombreComercial", _columna(fila, "Nombre Comercial", "nombreComercial"), cobertura)
    _sub(info, "ruc", _columna(fila, "RUC Emisor", "rucEmisor") or clave.get("ruc", ""), cobertura, DIRECTO, True)
    _sub(info, "claveAcceso", _columna(fila, "Clave de Acceso", "claveAcceso"), cobertura, DIRECTO, True)
    # Los cuatro siguientes salen de la clave de acceso, que es mas confiable
    # que la columna: no depende de como el extractor leyo el RIDE.
    _sub(info, "codDoc", clave.get("codDoc") or _codigo_de_etiqueta(_columna(fila, "Código del Documento")), cobertura, DEDUCIDO, True)
    _sub(info, "estab", clave.get("estab") or _columna(fila, "Establecimiento").zfill(3), cobertura, DEDUCIDO, True)
    _sub(info, "ptoEmi", clave.get("ptoEmi") or _columna(fila, "Punto de Emisión").zfill(3), cobertura, DEDUCIDO, True)
    _sub(info, "secuencial", clave.get("secuencial") or _columna(fila, "Secuencial").zfill(9), cobertura, DEDUCIDO, True)
    _sub(info, "dirMatriz", _columna(fila, "Dirección Matriz", "direccionMatrizEmisor"), cobertura, DIRECTO, True)


def _total_con_impuestos(padre: ET.Element, fila: pd.Series, cobertura: Cobertura) -> None:
    """Arma totalConImpuestos desde las bases y el monto de IVA del reporte."""
    contenedor = ET.SubElement(padre, "totalConImpuestos")
    tarifa = _columna(fila, "Tarifas IVA")
    codigo_pct = codigo_porcentaje_iva(tarifa)
    base_gravada = _numero(_columna(fila, "Base Gravada"))
    base_no_gravada = _numero(_columna(fila, "Base No Gravada"))
    monto_iva = _numero(_columna(fila, "Monto IVA"))

    def _bloque(cod_pct: str, base: float, valor: float) -> None:
        nodo = ET.SubElement(contenedor, "totalImpuesto")
        ET.SubElement(nodo, "codigo").text = _COD_IVA
        ET.SubElement(nodo, "codigoPorcentaje").text = cod_pct
        ET.SubElement(nodo, "baseImponible").text = _dec(base)
        ET.SubElement(nodo, "valor").text = _dec(valor)

    if base_gravada > 0 and codigo_pct not in ("", "0"):
        _bloque(codigo_pct, base_gravada, monto_iva)
        cobertura.anota("totalConImpuestos", DIRECTO)
    if base_no_gravada > 0:
        _bloque("0", base_no_gravada, 0.0)
        cobertura.anota("totalConImpuestos", DIRECTO)
    if not len(contenedor):
        # Sin bases utilizables: se declara el total sin impuestos a tarifa 0
        # para no dejar el bloque vacio, que el XSD no admite.
        _bloque(codigo_pct or "0", _numero(_columna(fila, "Total Sin Impuestos")), monto_iva)
        cobertura.anota("totalConImpuestos", DEDUCIDO)


def _detalles(padre: ET.Element, fila: pd.Series, cobertura: Cobertura,
              avisos: list[str]) -> int:
    contenedor = ET.SubElement(padre, "detalles")
    items = parsear_descripciones(_columna(fila, "Descripciones"))
    if not items:
        cobertura.anota("detalles", FALTANTE)
        avisos.append("no se pudo reconstruir el detalle de items")
        return 0

    tarifa = _columna(fila, "Tarifas IVA")
    codigo_pct = codigo_porcentaje_iva(tarifa) or "0"
    multiples = len(tarifas_declaradas(tarifa)) > 1
    if multiples:
        avisos.append(
            "el comprobante declara mas de una tarifa de IVA: el impuesto por "
            "linea se repartio con la primera y es una suposicion"
        )
    total_sin_imp = sum(i["precioTotalSinImpuesto"] for i in items)
    monto_iva = _numero(_columna(fila, "Monto IVA"))

    for item in items:
        nodo = ET.SubElement(contenedor, "detalle")
        ET.SubElement(nodo, "codigoPrincipal").text = item["codigoPrincipal"]
        if item["codigoAuxiliar"] and item["codigoAuxiliar"].upper() != "S/N":
            ET.SubElement(nodo, "codigoAuxiliar").text = item["codigoAuxiliar"]
        ET.SubElement(nodo, "descripcion").text = item["descripcion"] or "SIN DESCRIPCION"
        ET.SubElement(nodo, "cantidad").text = _dec(item["cantidad"], 6)
        ET.SubElement(nodo, "precioUnitario").text = _dec(item["precioUnitario"], 6)
        ET.SubElement(nodo, "descuento").text = _dec(item["descuento"])
        ET.SubElement(nodo, "precioTotalSinImpuesto").text = _dec(item["precioTotalSinImpuesto"])
        # El impuesto por linea no viaja en el reporte: se reparte el IVA del
        # comprobante en proporcion al peso de cada linea.
        impuestos = ET.SubElement(nodo, "impuestos")
        imp = ET.SubElement(impuestos, "impuesto")
        proporcion = (item["precioTotalSinImpuesto"] / total_sin_imp) if total_sin_imp else 0
        ET.SubElement(imp, "codigo").text = _COD_IVA
        ET.SubElement(imp, "codigoPorcentaje").text = codigo_pct
        ET.SubElement(imp, "tarifa").text = _dec(_tarifa_numerica(tarifa))
        ET.SubElement(imp, "baseImponible").text = _dec(item["precioTotalSinImpuesto"])
        ET.SubElement(imp, "valor").text = _dec(monto_iva * proporcion)

    cobertura.anota("detalles", DIRECTO)
    cobertura.anota("detalle/impuestos", DEDUCIDO)
    return len(items)


def _tarifa_numerica(tarifa: str) -> float:
    match = re.search(r"(\d{1,2})\s*%", _texto(tarifa))
    return float(match.group(1)) if match else 0.0


def _info_adicional(raiz: ET.Element, fila: pd.Series, cobertura: Cobertura) -> None:
    campos = parsear_campos_adicionales(_columna(fila, "Campos Adicionales", "informacionAdicional"))
    nodo = ET.SubElement(raiz, "infoAdicional")
    for nombre, valor in campos:
        campo = ET.SubElement(nodo, "campoAdicional", {"nombre": nombre[:300]})
        campo.text = valor[:300]
    if campos:
        cobertura.anota("infoAdicional", DIRECTO)
    # Marca permanente: este archivo no es el comprobante original.
    marca = ET.SubElement(nodo, "campoAdicional", {"nombre": "reconstruido"})
    marca.text = (
        "XML reconstruido desde el reporte del robot. SIN FIRMA DIGITAL: "
        "no sustituye al comprobante autorizado ni tiene validez tributaria."
    )


def construir_factura(fila: pd.Series) -> tuple[ET.Element, Cobertura, list[str]]:
    cobertura, avisos = Cobertura(), []
    clave = descomponer_clave_acceso(_columna(fila, "Clave de Acceso", "claveAcceso"))
    if not clave:
        avisos.append("la clave de acceso no tiene 49 digitos: infoTributaria sale del reporte")

    raiz = ET.Element("factura", {"id": "comprobante", "version": "1.1.0"})
    _info_tributaria(raiz, fila, clave, cobertura)

    info = ET.SubElement(raiz, "infoFactura")
    _sub(info, "fechaEmision", _columna(fila, "Fecha de Emisión", "fechaEmision") or clave.get("fechaEmision", ""), cobertura, DIRECTO, True)
    _sub(info, "dirEstablecimiento", _columna(fila, "Dir. Establecimiento", "direccionSucursalEmisor"), cobertura)
    obligado = _columna(fila, "Obligado Contabilidad", "obligadoContabilidad").upper()
    _sub(info, "obligadoContabilidad", obligado if obligado in {"SI", "NO"} else "", cobertura)
    _sub(info, "tipoIdentificacionComprador", _codigo_de_etiqueta(_columna(fila, "Tipo Identificación Comprador")), cobertura, DEDUCIDO, True)
    _sub(info, "razonSocialComprador", _columna(fila, "Razón Social Comprador"), cobertura, DIRECTO, True)
    ident, origen_ident = identificacion_comprador(fila)
    _sub(info, "identificacionComprador", ident, cobertura, origen_ident, True)
    if origen_ident == FALTANTE:
        avisos.append(
            "la identificacion del comprador no vino en el reporte; el XML "
            "queda sin ella y hay que completarla a mano"
        )
    elif origen_ident == DEDUCIDO:
        avisos.append("identificacion deducida: consumidor final")
    _sub(info, "direccionComprador", _columna(fila, "Dirección Comprador"), cobertura)
    _sub(info, "totalSinImpuestos", _dec(_columna(fila, "Total Sin Impuestos")), cobertura, DIRECTO, True)
    _sub(info, "totalDescuento", _dec(_columna(fila, "Total Descuento")), cobertura, DIRECTO, True)
    _total_con_impuestos(info, fila, cobertura)
    _sub(info, "propina", _dec(_columna(fila, "Propina")), cobertura)
    _sub(info, "importeTotal", _dec(_columna(fila, "Importe Total")), cobertura, DIRECTO, True)
    _sub(info, "moneda", _columna(fila, "Moneda"), cobertura)

    pagos = parsear_formas_pago(_columna(fila, "Forma Pago"))
    if pagos:
        contenedor = ET.SubElement(info, "pagos")
        for pago in pagos:
            nodo = ET.SubElement(contenedor, "pago")
            ET.SubElement(nodo, "formaPago").text = pago["formaPago"]
            ET.SubElement(nodo, "total").text = _dec(pago["total"])
        cobertura.anota("pagos", DIRECTO)
    else:
        cobertura.anota("pagos", FALTANTE)
        avisos.append("no se pudo leer la forma de pago")

    _detalles(raiz, fila, cobertura, avisos)
    _info_adicional(raiz, fila, cobertura)
    return raiz, cobertura, avisos


def construir_nota(fila: pd.Series, es_credito: bool) -> tuple[ET.Element, Cobertura, list[str]]:
    """Nota de credito o de debito.

    Comparte 37 de las 38 columnas con la factura; lo propio es el documento
    que modifica. Implementado pero AUN SIN CONTRASTAR contra un reporte real
    de este tipo.
    """
    cobertura, avisos = Cobertura(), []
    avisos.append("tipo implementado pero no validado con un reporte real")
    clave = descomponer_clave_acceso(_columna(fila, "Clave de Acceso", "claveAcceso"))

    etiqueta = "notaCredito" if es_credito else "notaDebito"
    raiz = ET.Element(etiqueta, {"id": "comprobante", "version": "1.0.0"})
    _info_tributaria(raiz, fila, clave, cobertura)

    info = ET.SubElement(raiz, "infoNotaCredito" if es_credito else "infoNotaDebito")
    _sub(info, "fechaEmision", _columna(fila, "Fecha de Emisión") or clave.get("fechaEmision", ""), cobertura, DIRECTO, True)
    _sub(info, "dirEstablecimiento", _columna(fila, "Dir. Establecimiento"), cobertura)
    _sub(info, "tipoIdentificacionComprador", _codigo_de_etiqueta(_columna(fila, "Tipo Identificación Comprador")), cobertura, DEDUCIDO, True)
    _sub(info, "razonSocialComprador", _columna(fila, "Razón Social Comprador"), cobertura, DIRECTO, True)
    ident, origen_ident = identificacion_comprador(fila)
    _sub(info, "identificacionComprador", ident, cobertura, origen_ident, True)
    if origen_ident == FALTANTE:
        avisos.append("la identificacion del comprador no vino en el reporte")
    _sub(info, "codDocModificado", _codigo_de_etiqueta(_columna(fila, "Código Documento Modificado")), cobertura, DIRECTO, True)
    _sub(info, "numDocModificado", _columna(fila, "Número Documento Modificado"), cobertura, DIRECTO, True)
    _sub(info, "fechaEmisionDocSustento", _columna(fila, "Fecha Emisión Doc. Sustento"), cobertura, DIRECTO, True)
    _sub(info, "totalSinImpuestos", _dec(_columna(fila, "Total Sin Impuestos")), cobertura, DIRECTO, True)
    if es_credito:
        _sub(info, "valorModificacion", _dec(_columna(fila, "Valor Modificación", "Importe Total")), cobertura, DIRECTO, True)
        _sub(info, "moneda", _columna(fila, "Moneda"), cobertura)
        _total_con_impuestos(info, fila, cobertura)
        _sub(info, "motivo", _columna(fila, "Motivo"), cobertura)
    else:
        _total_con_impuestos(info, fila, cobertura)
        _sub(info, "valorTotal", _dec(_columna(fila, "Importe Total")), cobertura, DIRECTO, True)

    if es_credito:
        _detalles(raiz, fila, cobertura, avisos)
    _info_adicional(raiz, fila, cobertura)
    return raiz, cobertura, avisos


# =============================================================================
# Validacion y escritura
# =============================================================================


def _ruta_xsd(cod_doc: str) -> Optional[Path]:
    if cod_doc not in TIPOS:
        return None
    _, subcarpeta, archivo = TIPOS[cod_doc]
    ruta = CARPETA_XSD / subcarpeta / archivo
    return ruta if ruta.exists() else None


def validar_contra_xsd(xml_bytes: bytes, cod_doc: str) -> tuple[bool, str]:
    """Valida contra el XSD oficial. Sin lxml devuelve (True, aviso)."""
    if _lxml_etree is None:
        return True, "lxml no esta instalado: no se valido contra el XSD"
    ruta = _ruta_xsd(cod_doc)
    if ruta is None:
        return True, f"no se encontro el XSD para codDoc {cod_doc}"
    try:
        esquema = _lxml_etree.XMLSchema(_lxml_etree.parse(str(ruta)))
        esquema.assertValid(_lxml_etree.fromstring(xml_bytes))
        return True, ""
    except Exception as err:
        return False, str(err).replace("\n", " ")[:300]


def _escribir_xml(raiz: ET.Element, destino: Path) -> bytes:
    ET.indent(raiz, space="  ")
    cuerpo = ET.tostring(raiz, encoding="utf-8")
    contenido = b'<?xml version="1.0" encoding="UTF-8"?>\n' + cuerpo
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(contenido)
    return contenido


def _nombre_archivo(fila: pd.Series, cod_doc: str) -> str:
    clave = re.sub(r"\D", "", _columna(fila, "Clave de Acceso", "claveAcceso"))
    if clave:
        return f"{clave}.xml"
    nombre = TIPOS.get(cod_doc, ("comprobante",))[0].replace(" ", "_")
    secuencial = _columna(fila, "Secuencial") or "sin_secuencial"
    return f"{nombre}_{secuencial}.xml"


# =============================================================================
# Entrada principal
# =============================================================================


def generar_xml_desde_reporte(
    *,
    reportes: Iterable[str | Path],
    destino: str | Path,
    informe_cobertura: Optional[str | Path] = None,
    validar: bool = True,
    progress: Optional[Callable[[str], None]] = None,
) -> dict:
    """Reconstruye un XML por cada fila de los reportes de Emitidos indicados.

    Devuelve un resumen con lo generado, lo que fallo y la ruta del informe de
    cobertura.
    """

    def _avisar(mensaje: str) -> None:
        if progress:
            try:
                progress(mensaje)
            except Exception:
                pass
        logger.info(mensaje)

    destino = Path(destino)
    resumen = {
        "ok": False,
        "total_filas": 0,
        "repetidos": 0,
        "generados": 0,
        "fallidos": 0,
        "invalidos": 0,
        "por_tipo": {},
        "no_soportados": {},
        "destino": str(destino),
        "informe_cobertura": "",
        "message": "",
    }
    filas_informe: list[dict] = []
    # Un mismo comprobante puede venir en el reporte diario y en el mensual del
    # mismo mes. Sin esto se procesaria dos veces: el archivo se sobrescribe --
    # el nombre es la clave de acceso -- pero los conteos y el informe saldrian
    # inflados, que es peor que perder tiempo.
    claves_vistas: set[str] = set()

    rutas = [Path(r) for r in reportes]
    existentes = [r for r in rutas if r.exists()]
    if not existentes:
        resumen["message"] = "No se encontro ninguno de los reportes indicados."
        return resumen

    for ruta in existentes:
        try:
            # `dtype=str` no es opcional: sin el, pandas infiere las columnas de
            # identificadores como numeros y se come el cero inicial. La cedula
            # 0503174286 se leia como 503174286, y todas las cedulas de las
            # provincias 01 a 09 quedaban mal. El reporte las guarda bien, con
            # formato de texto; el problema era leerlas.
            df = pd.read_excel(ruta, dtype=str)
        except Exception as err:
            _avisar(f"No se pudo leer '{ruta.name}': {err}")
            continue
        _avisar(f"{ruta.name}: {len(df)} fila(s)")

        for indice, fila in df.iterrows():
            resumen["total_filas"] += 1
            clave_txt = _columna(fila, "Clave de Acceso", "claveAcceso")
            clave_norm = re.sub(r"\D", "", clave_txt)
            if clave_norm and clave_norm in claves_vistas:
                resumen["repetidos"] += 1
                continue
            if clave_norm:
                claves_vistas.add(clave_norm)
            clave = descomponer_clave_acceso(clave_txt)
            cod_doc = clave.get("codDoc") or _codigo_de_etiqueta(
                _columna(fila, "Código del Documento")
            )
            if not cod_doc and _columna(fila, "tipoDocumento"):
                cod_doc = {"factura": "01", "retencion": "07"}.get(
                    _sin_acentos(_columna(fila, "tipoDocumento")).lower(), ""
                )

            if cod_doc in TIPOS_NO_SOPORTADOS:
                nombre, motivo = TIPOS_NO_SOPORTADOS[cod_doc]
                resumen["no_soportados"].setdefault(nombre, motivo)
                resumen["fallidos"] += 1
                filas_informe.append(
                    {
                        "Reporte": ruta.name,
                        "Fila": indice + 2,
                        "Clave de acceso": clave_txt,
                        "Tipo": nombre,
                        "Resultado": "no soportado",
                        "Detalle": motivo,
                    }
                )
                continue

            if cod_doc not in TIPOS:
                resumen["fallidos"] += 1
                filas_informe.append(
                    {
                        "Reporte": ruta.name,
                        "Fila": indice + 2,
                        "Clave de acceso": clave_txt,
                        "Tipo": f"codDoc {cod_doc or '?'}",
                        "Resultado": "no reconocido",
                        "Detalle": "no se pudo determinar el tipo de comprobante",
                    }
                )
                continue

            try:
                if cod_doc == "01":
                    raiz, cobertura, avisos = construir_factura(fila)
                else:
                    raiz, cobertura, avisos = construir_nota(fila, cod_doc == "04")
            except Exception as err:
                resumen["fallidos"] += 1
                filas_informe.append(
                    {
                        "Reporte": ruta.name,
                        "Fila": indice + 2,
                        "Clave de acceso": clave_txt,
                        "Tipo": TIPOS[cod_doc][0],
                        "Resultado": "error",
                        "Detalle": str(err)[:250],
                    }
                )
                continue

            nombre_tipo = TIPOS[cod_doc][0]
            salida = destino / nombre_tipo.replace(" ", "_") / _nombre_archivo(fila, cod_doc)
            contenido = _escribir_xml(raiz, salida)

            valido, detalle_xsd = (True, "")
            if validar:
                valido, detalle_xsd = validar_contra_xsd(contenido, cod_doc)
            if not valido:
                resumen["invalidos"] += 1
                avisos.append(f"no valida contra el XSD: {detalle_xsd}")

            resumen["generados"] += 1
            resumen["por_tipo"][nombre_tipo] = resumen["por_tipo"].get(nombre_tipo, 0) + 1
            filas_informe.append(
                {
                    "Reporte": ruta.name,
                    "Fila": indice + 2,
                    "Clave de acceso": clave_txt,
                    "Tipo": nombre_tipo,
                    "Resultado": "generado" if valido else "generado (no valida XSD)",
                    "Archivo": str(salida),
                    "Campos del reporte": cobertura.cuenta(DIRECTO),
                    "Campos deducidos": cobertura.cuenta(DEDUCIDO),
                    "Campos faltantes": cobertura.cuenta(FALTANTE),
                    "Deducidos": cobertura.lista(DEDUCIDO),
                    "Faltantes": cobertura.lista(FALTANTE),
                    "Detalle": "; ".join(avisos),
                }
            )

    if informe_cobertura and filas_informe:
        ruta_informe = Path(informe_cobertura)
        ruta_informe.parent.mkdir(parents=True, exist_ok=True)
        _escribir_informe(filas_informe, resumen, ruta_informe)
        resumen["informe_cobertura"] = str(ruta_informe)

    resumen["ok"] = resumen["generados"] > 0
    resumen["message"] = (
        f"{resumen['generados']} XML generados de {resumen['total_filas']} fila(s)."
        + (
            f" {resumen['repetidos']} fila(s) repetidas entre reportes, omitidas."
            if resumen["repetidos"]
            else ""
        )
        + (f" {resumen['fallidos']} sin generar." if resumen["fallidos"] else "")
        + (f" {resumen['invalidos']} no validan contra el XSD." if resumen["invalidos"] else "")
    )
    _avisar(resumen["message"])
    return resumen


def _escribir_informe(filas: list[dict], resumen: dict, path: Path) -> None:
    """Dos hojas: el resumen para decidir, el detalle para auditar."""
    generales = [
        ("Filas leidas", resumen["total_filas"]),
        ("Repetidas entre reportes", resumen["repetidos"]),
        ("XML generados", resumen["generados"]),
        ("Sin generar", resumen["fallidos"]),
        ("No validan contra el XSD", resumen["invalidos"]),
    ]
    for tipo, cantidad in sorted(resumen["por_tipo"].items()):
        generales.append((f"  de tipo {tipo}", cantidad))
    for tipo, motivo in sorted(resumen["no_soportados"].items()):
        generales.append((f"  {tipo}: no soportado", motivo))
    generales.append(
        (
            "ADVERTENCIA",
            "Los XML no llevan firma digital. No sustituyen al comprobante "
            "autorizado ni tienen validez tributaria.",
        )
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(generales, columns=["Concepto", "Valor"]).to_excel(
            writer, sheet_name="Resumen", index=False
        )
        pd.DataFrame(filas).to_excel(writer, sheet_name="Detalle por comprobante", index=False)
