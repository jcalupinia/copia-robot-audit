"""Parsing de XML de comprobantes electrónicos emitidos.

Extrae los datos de los archivos XML descargados del SRI hacia las filas
de reporte (dicts con las columnas de `report_columns`). Cubre:

- Facturas emitidas (`_extraer_datos_xml_factura_emitido`).
- Notas de crédito (`_extraer_datos_xml_nota_credito_emitido`).
- Notas de débito (`_extraer_datos_xml_nota_debito_emitido`).

`_extraer_xml_emitidos_autorizacion` es el helper común que abre el XML,
separa el sobre de autorización del comprobante interno y normaliza los
namespaces; `_strip_xml_namespaces` hace esa última parte.

Los parsers de retención emitida, liquidación de compra y el XML de
recibidos siguen en `robot/downloader.py` por ahora — dependen de helpers
de extracción por regex compartidos con los parsers de PDF, que se
moverán en Fase 3b.

Extraído en la Sub-fase 3a-ii del refactor.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path

from robot.data_formatters import (
    _factura_emitidos_default_row,
    _label_ambiente_emitidos_retencion,
    _label_emision_emitidos_retencion,
    _label_forma_pago_emitidos_retencion,
    _label_tipo_ident_emitidos_nota_credito,
    _nota_credito_emitidos_default_row,
    _nota_debito_emitidos_default_row,
    _numero_emitidos_retencion,
    _texto_emitidos_retencion,
    _texto_emitidos_retencion_na,
)
from robot.report_columns import EMITIDOS_RETENCION_DOC_CODE_LABEL


# --------------------------------------------------------------------------- #
# Helpers de namespaces / apertura del XML
# --------------------------------------------------------------------------- #
def _strip_xml_namespaces(element: ET.Element):
    """Elimina prefijos de namespace de tags y atributos, in-place."""
    if element is None:
        return
    for node in element.iter():
        if '}' in node.tag:
            node.tag = node.tag.split('}', 1)[1]
        if ':' in node.tag:
            node.tag = node.tag.split(':', 1)[1]
        if node.attrib:
            node.attrib = {
                key.split('}', 1)[-1].split(':', 1)[-1]: val
                for key, val in node.attrib.items()
            }


def _extraer_xml_emitidos_autorizacion(xml_path: Path) -> tuple[ET.Element | None, dict]:
    """Abre un XML de comprobante emitido y devuelve (root_comprobante, meta).

    Si el archivo es un sobre de autorización, extrae el comprobante interno
    y los metadatos de autorización (estado, número, fecha, ambiente).
    Devuelve (None, {}) o (None, meta) si el XML no se puede parsear.
    """
    try:
        contenido = xml_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None, {}
    if not contenido:
        return None, {}
    try:
        root = ET.fromstring(contenido)
    except ET.ParseError:
        return None, {}
    meta = {}
    comprobante_xml = contenido
    if root.tag.lower().endswith("autorizacion"):
        meta = {
            "estado": _texto_emitidos_retencion(root.findtext("estado")),
            "numero_autorizacion": _texto_emitidos_retencion(root.findtext("numeroAutorizacion")),
            "fecha_autorizacion": _texto_emitidos_retencion(root.findtext("fechaAutorizacion")),
            "ambiente": _texto_emitidos_retencion(root.findtext("ambiente")),
        }
        comprobante_xml = root.findtext("comprobante") or ""
    try:
        comprobante_root = ET.fromstring(comprobante_xml)
    except ET.ParseError:
        return None, meta
    _strip_xml_namespaces(comprobante_root)
    return comprobante_root, meta


# --------------------------------------------------------------------------- #
# Nota de Crédito
# --------------------------------------------------------------------------- #
def _extraer_datos_xml_nota_credito_emitido(xml_path: Path) -> dict:
    row = _nota_credito_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_nc = root.find("infoNotaCredito")
    detalles = root.findall(".//detalles/detalle")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(info_trib.findtext("nombreComercial"))
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, cod_doc or "No Disponible")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_nc is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(info_nc.findtext("dirEstablecimiento"))
        row["Obligado Contabilidad"] = _texto_emitidos_retencion_na(info_nc.findtext("obligadoContabilidad"))
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_nc.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_nc.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_nc.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_nc.findtext("razonSocialComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_nc.findtext("moneda"))
        row["Código Documento Modificado"] = _texto_emitidos_retencion_na(info_nc.findtext("codDocModificado"))
        row["Número Documento Modificado"] = _texto_emitidos_retencion_na(info_nc.findtext("numDocModificado"))
        row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(info_nc.findtext("fechaEmisionDocSustento"))
        row["Motivo"] = _texto_emitidos_retencion_na(info_nc.findtext("motivo"))
        row["Valor Modificación"] = _numero_emitidos_retencion(info_nc.findtext("valorModificacion"))
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_nc.findtext("totalSinImpuestos"))

    detalle_textos = []
    base_gravada = 0
    base_no_gravada = 0
    monto_iva = 0
    base_gravada_15 = 0
    monto_iva_15 = 0
    tarifas = []
    for detalle in detalles:
        codigo = _texto_emitidos_retencion(detalle.findtext("codigoInterno") or detalle.findtext("codigoPrincipal"))
        descripcion = (detalle.findtext("descripcion") or "").strip()
        cantidad = _texto_emitidos_retencion(detalle.findtext("cantidad"))
        precio_unitario = _texto_emitidos_retencion(detalle.findtext("precioUnitario"))
        partes = []
        if codigo:
            partes.append(f"Código: {codigo}")
        if descripcion:
            partes.append(f"Desc: {descripcion}")
        if cantidad:
            partes.append(f"Cant: {cantidad}")
        if precio_unitario:
            partes.append(f"P.Unit: {precio_unitario}")
        if partes:
            detalle_textos.append(", ".join(partes))

        for imp in detalle.findall("./impuestos/impuesto"):
            codigo = _texto_emitidos_retencion(imp.findtext("codigo"))
            codigo_pct = _texto_emitidos_retencion(imp.findtext("codigoPorcentaje"))
            tarifa = _numero_emitidos_retencion(imp.findtext("tarifa"))
            base = _numero_emitidos_retencion(imp.findtext("baseImponible"))
            valor = _numero_emitidos_retencion(imp.findtext("valor"))
            if codigo == "2":
                if codigo_pct == "0":
                    base_no_gravada += base
                else:
                    base_gravada += base
                    monto_iva += valor
                    if codigo_pct == "4":
                        base_gravada_15 += base
                        monto_iva_15 += valor
                if tarifa:
                    etiqueta = f"{int(tarifa) if float(tarifa).is_integer() else tarifa}%"
                    if etiqueta not in tarifas:
                        tarifas.append(etiqueta)

    row["Descripciones"] = " | ".join(detalle_textos)
    row["Forma Pago"] = "No Disponible - No Disponible"
    row["Total Sin Impuestos"] = row["Total Sin Impuestos"] or base_gravada or base_no_gravada
    row["Base Gravada"] = base_gravada
    row["Base No Gravada"] = base_no_gravada
    row["Tarifas IVA"] = ", ".join(tarifas)
    row["Monto IVA"] = monto_iva
    row["Importe Total"] = row["Valor Modificación"] or (row["Total Sin Impuestos"] + row["Monto IVA"])
    row["Total Pago"] = 0
    row["Base Gravada 15%"] = base_gravada_15
    row["Monto IVA 15%"] = monto_iva_15

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)

    return row


# --------------------------------------------------------------------------- #
# Nota de Débito
# --------------------------------------------------------------------------- #
def _extraer_datos_xml_nota_debito_emitido(xml_path: Path) -> dict:
    row = _nota_debito_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_nd = root.find("infoNotaDebito")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(
            info_trib.findtext("nombreComercial") or info_trib.findtext("razonSocial")
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, "05 - NOTA DE DÉBITO")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_nd is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
            info_nd.findtext("dirEstablecimiento") or row["Dirección Matriz"]
        )
        obligado = _texto_emitidos_retencion(info_nd.findtext("obligadoContabilidad"))
        row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_nd.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_nd.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_nd.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_nd.findtext("razonSocialComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_nd.findtext("moneda") or "DOLAR")
        row["Código Documento Modificado"] = _texto_emitidos_retencion_na(info_nd.findtext("codDocModificado"))
        row["Número Documento Modificado"] = _texto_emitidos_retencion_na(info_nd.findtext("numDocModificado"))
        row["Fecha Emisión Doc. Sustento"] = _texto_emitidos_retencion_na(info_nd.findtext("fechaEmisionDocSustento"))
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_nd.findtext("totalSinImpuestos"))
        row["Importe Total"] = _numero_emitidos_retencion(info_nd.findtext("valorTotal"))

        pago = info_nd.find("./pagos/pago")
        if pago is not None:
            forma = _label_forma_pago_emitidos_retencion(pago.findtext("formaPago"))
            row["Forma Pago"] = f"{forma} - {forma}" if forma != "No Disponible" else "No Disponible - No Disponible"
            row["Total Pago"] = _numero_emitidos_retencion(
                pago.findtext("total") or info_nd.findtext("valorTotal")
            )
            row["Plazo Pago"] = _texto_emitidos_retencion_na(pago.findtext("plazo"))
            row["Unidad Tiempo Pago"] = _texto_emitidos_retencion_na(pago.findtext("unidadTiempo"))
        else:
            row["Forma Pago"] = "No Disponible - No Disponible"
            row["Total Pago"] = row["Importe Total"]

    base_gravada = 0
    base_no_gravada = 0
    monto_iva = 0
    base_gravada_15 = 0
    monto_iva_15 = 0
    tarifas = []
    for imp in root.findall(".//impuestos/impuesto"):
        codigo = _texto_emitidos_retencion(imp.findtext("codigo"))
        codigo_pct = _texto_emitidos_retencion(imp.findtext("codigoPorcentaje"))
        tarifa = _numero_emitidos_retencion(imp.findtext("tarifa"))
        base = _numero_emitidos_retencion(imp.findtext("baseImponible"))
        valor = _numero_emitidos_retencion(imp.findtext("valor"))
        if codigo == "2":
            if codigo_pct == "0" or not tarifa:
                base_no_gravada += base
            else:
                base_gravada += base
                monto_iva += valor
                if codigo_pct == "4" or abs(tarifa - 15) < 0.001:
                    base_gravada_15 += base
                    monto_iva_15 += valor
            if tarifa:
                etiqueta = f"{int(tarifa) if float(tarifa).is_integer() else tarifa}%"
                if etiqueta not in tarifas:
                    tarifas.append(etiqueta)

    motivos = []
    valor_modificacion = 0
    for motivo in root.findall(".//motivos/motivo"):
        razon = _texto_emitidos_retencion_na(motivo.findtext("razon"))
        valor = _numero_emitidos_retencion(motivo.findtext("valor"))
        if razon and razon != "No Disponible":
            motivos.append(razon)
        valor_modificacion += valor

    row["Motivo"] = " | ".join(motivos) if motivos else row["Motivo"]
    row["Descripciones"] = row["Motivo"] if row["Motivo"] != "No Disponible" else row["Descripciones"]
    row["Valor Modificación"] = valor_modificacion or row["Importe Total"]
    row["Base Gravada"] = base_gravada
    row["Base No Gravada"] = base_no_gravada
    row["Tarifas IVA"] = ", ".join(tarifas)
    row["Monto IVA"] = monto_iva
    row["Base Gravada 15%"] = base_gravada_15
    row["Monto IVA 15%"] = monto_iva_15
    if not row["Total Pago"]:
        row["Total Pago"] = row["Importe Total"]

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)

    return row


# --------------------------------------------------------------------------- #
# Factura
# --------------------------------------------------------------------------- #
def _extraer_datos_xml_factura_emitido(xml_path: Path) -> dict:
    row = _factura_emitidos_default_row()
    root, meta = _extraer_xml_emitidos_autorizacion(xml_path)
    if root is None:
        return row

    info_trib = root.find("infoTributaria")
    info_fact = root.find("infoFactura")
    detalles = root.findall(".//detalles/detalle")

    row["Estado"] = _texto_emitidos_retencion(meta.get("estado"), "AUTORIZADO")
    row["Número de Autorización"] = _texto_emitidos_retencion(meta.get("numero_autorizacion"))
    row["Fecha de Autorización"] = _texto_emitidos_retencion(meta.get("fecha_autorizacion"))

    if info_trib is not None:
        cod_doc = _texto_emitidos_retencion(info_trib.findtext("codDoc"))
        row["Ambiente"] = _label_ambiente_emitidos_retencion(info_trib.findtext("ambiente") or meta.get("ambiente"))
        row["Razón Social Emisor"] = _texto_emitidos_retencion(info_trib.findtext("razonSocial"))
        row["Nombre Comercial"] = _texto_emitidos_retencion_na(
            info_trib.findtext("nombreComercial") or info_trib.findtext("razonSocial")
        )
        row["Tipo Emisión"] = _label_emision_emitidos_retencion(info_trib.findtext("tipoEmision"))
        row["Código del Documento"] = EMITIDOS_RETENCION_DOC_CODE_LABEL.get(cod_doc, "01 - FACTURA")
        row["Establecimiento"] = _texto_emitidos_retencion(info_trib.findtext("estab"))
        row["Punto de Emisión"] = _texto_emitidos_retencion(info_trib.findtext("ptoEmi"))
        row["Secuencial"] = _texto_emitidos_retencion(info_trib.findtext("secuencial"))
        row["Dirección Matriz"] = _texto_emitidos_retencion_na(info_trib.findtext("dirMatriz"))
        row["RUC Emisor"] = _texto_emitidos_retencion(info_trib.findtext("ruc"))
        row["Clave de Acceso"] = _texto_emitidos_retencion(info_trib.findtext("claveAcceso"))

    if info_fact is not None:
        row["Dir. Establecimiento"] = _texto_emitidos_retencion_na(
            info_fact.findtext("dirEstablecimiento") or row["Dirección Matriz"]
        )
        obligado = _texto_emitidos_retencion(info_fact.findtext("obligadoContabilidad"))
        row["Obligado Contabilidad"] = obligado if obligado in {"SI", "NO"} else "No Disponible"
        row["Tipo Identificación Comprador"] = _label_tipo_ident_emitidos_nota_credito(
            info_fact.findtext("tipoIdentificacionComprador")
        )
        row["Identificación Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("identificacionComprador"))
        row["Fecha de Emisión"] = _texto_emitidos_retencion(info_fact.findtext("fechaEmision"))
        row["Razón Social Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("razonSocialComprador"))
        row["Dirección Comprador"] = _texto_emitidos_retencion_na(info_fact.findtext("direccionComprador"))
        row["Moneda"] = _texto_emitidos_retencion_na(info_fact.findtext("moneda") or "DOLAR")
        row["Total Sin Impuestos"] = _numero_emitidos_retencion(info_fact.findtext("totalSinImpuestos"))
        row["Total Descuento"] = _numero_emitidos_retencion(info_fact.findtext("totalDescuento"))
        row["Propina"] = _numero_emitidos_retencion(info_fact.findtext("propina"))
        row["Importe Total"] = _numero_emitidos_retencion(info_fact.findtext("importeTotal"))

        pago = info_fact.find("./pagos/pago")
        if pago is not None:
            forma = _label_forma_pago_emitidos_retencion(pago.findtext("formaPago"))
            row["Forma Pago"] = f"{forma} - {forma}" if forma != "No Disponible" else "No Disponible - No Disponible"
            row["Total Pago"] = _numero_emitidos_retencion(pago.findtext("total"))
            row["Plazo Pago"] = _texto_emitidos_retencion_na(pago.findtext("plazo"))
            row["Unidad Tiempo Pago"] = _texto_emitidos_retencion_na(pago.findtext("unidadTiempo"))
        else:
            row["Forma Pago"] = "No Disponible - No Disponible"

    monto_iva = 0
    for imp in root.findall(".//infoFactura/totalConImpuestos/totalImpuesto"):
        valor = _numero_emitidos_retencion(imp.findtext("valor"))
        if valor:
            monto_iva += valor
    row["Base Gravada"] = 0
    row["Base No Gravada"] = row["Total Sin Impuestos"]
    row["Base No Gravada 0%"] = row["Total Sin Impuestos"]
    row["Tarifas IVA"] = "0%"
    row["Monto IVA"] = monto_iva

    detalle_textos = []
    for detalle in detalles:
        codigo = _texto_emitidos_retencion(detalle.findtext("codigoPrincipal") or detalle.findtext("codigoInterno"))
        descripcion = (detalle.findtext("descripcion") or "").strip()
        cantidad = _texto_emitidos_retencion(detalle.findtext("cantidad"))
        precio_unitario = _texto_emitidos_retencion(detalle.findtext("precioUnitario"))
        partes = [f"Código: {codigo}", f"Desc: {descripcion}"]
        if cantidad:
            partes.append(f"Cant: {cantidad}")
        if precio_unitario:
            partes.append(f"P.Unit: {precio_unitario}")
        detalle_textos.append(", ".join(partes))
    row["Descripciones"] = " ; ".join(detalle_textos)

    adicionales = []
    for campo in root.findall(".//infoAdicional/campoAdicional"):
        nombre = _texto_emitidos_retencion(campo.attrib.get("nombre"))
        valor = _texto_emitidos_retencion(campo.text)
        if nombre or valor:
            adicionales.append(f"{nombre}: {valor}".strip(": "))
    if adicionales:
        row["Campos Adicionales"] = "; ".join(adicionales)
    return row
