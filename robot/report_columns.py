"""Definiciones de columnas para los reportes Excel generados por el bot.

Cada reporte tiene:

- Una **lista de columnas** (`*_REPORT_COLUMNS`) que define el orden en el XLSX.
- Un **set de columnas forzadas a texto** (`*_TEXT_FORCE_COLUMNS`) que evita que
  openpyxl o Excel las interprete como números (ej. RUCs de 13 dígitos que
  empiezan con 0, claves de acceso de 49 dígitos).
- Un **set de columnas numéricas** (`*_NUMERIC_COLUMNS`) que se formatean como
  número con decimales.
- **Mapeos de etiquetas** (`*_LABEL`) que traducen códigos del XML/JSON del
  SRI a textos legibles ("01" → "01 - SIN UTILIZACIÓN DEL SISTEMA FINANCIERO").

Todo lo que está acá es **data, no lógica**: si una columna nueva aparece en
el portal SRI, se agrega aquí sin tocar las funciones de generación.

Originalmente vivía en `robot/downloader.py`; extraído en la Sub-fase 2c-i
del refactor.
"""
from __future__ import annotations


# =========================================================================== #
# Reportes Recibidos — PDF
# =========================================================================== #
PDF_REPORT_COLUMNS = [
    "tipoDocumento",
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "contribuyenteEspecial",
    "agenteRetencion",
    "obligadoContabilidad",
    "tipoContribuyenteRIMPE",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "fechaAutorizacion",
    "razonSocialComprador",
    "identificacionComprador",
    "direccionComprador",
    "placa",
    "guia",
    "comprobanteModificado",
    "fechaEmisionModificado",
    "razonModificacion",
    "valorModificacion",
    "descripcionesProductos",
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
    "formaPago",
    "formaPagoMonto",
    "ambiente",
    "emision",
    "claveAcceso",
    "informacionAdicional",
]


# =========================================================================== #
# Reportes Recibidos — Retención
# =========================================================================== #
RETENCION_REPORT_COLUMNS = [
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "obligadoContabilidad",
    "numeroContribuyenteEspecial",
    "numeroAgenteRetencion",
    "fechaAutorizacion",
    "ambiente",
    "emision",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "razonSocialSujetoRetenido",
    "identificacionSujetoRetenido",
    "claveAcceso",
    "Comprobante_Sustento",
    "Numero_Sustento",
    "Fecha_Emision_Sustento",
    "Ejercicio_Fiscal",
    "Base_Imponible_Ret_IVA",
    "Impuesto_Ret_IVA",
    "Porcentaje_Ret_IVA",
    "Valor_Retenido_IVA",
    "Base_Imponible_Ret_IR",
    "Impuesto_Ret_IR",
    "Porcentaje_Ret_IR",
    "Valor_Retenido_IR",
    "informacionAdicional",
    "Base_Imponible_Ret_IR_1",
    "Impuesto_Ret_IR_1",
    "Porcentaje_Ret_IR_1",
    "Valor_Retenido_IR_1",
    "Base_Imponible_Ret_IVA_1",
    "Impuesto_Ret_IVA_1",
    "Porcentaje_Ret_IVA_1",
    "Valor_Retenido_IVA_1",
    "tipoDocumento",
]


# =========================================================================== #
# Reportes Emitidos — Retención
# =========================================================================== #
EMITIDOS_RETENCION_REPORT_COLUMNS = [
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "obligadoContabilidad",
    "numeroContribuyenteEspecial",
    "numeroAgenteRetencion",
    "fechaAutorizacion",
    "ambiente",
    "emision",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "razonSocialSujetoRetenido",
    "identificacionSujetoRetenido",
    "claveAcceso",
    "Comprobante_Sustento",
    "Numero_Sustento",
    "Fecha_Emision_Sustento",
    "Ejercicio_Fiscal",
    "Base_Imponible_Ret_IVA",
    "Impuesto_Ret_IVA",
    "Porcentaje_Ret_IVA",
    "Valor_Retenido_IVA",
    "Base_Imponible_Ret_IR",
    "Impuesto_Ret_IR",
    "Porcentaje_Ret_IR",
    "Valor_Retenido_IR",
    "informacionAdicional",
    "Base_Imponible_Ret_IR_1",
    "Impuesto_Ret_IR_1",
    "Porcentaje_Ret_IR_1",
    "Valor_Retenido_IR_1",
    "Base_Imponible_Ret_IVA_1",
    "Impuesto_Ret_IVA_1",
    "Porcentaje_Ret_IVA_1",
    "Valor_Retenido_IVA_1",
    "Base_Imponible_Ret_IR_2",
    "Impuesto_Ret_IR_2",
    "Porcentaje_Ret_IR_2",
    "Valor_Retenido_IR_2",
    "Base_Imponible_Ret_IVA_2",
    "Impuesto_Ret_IVA_2",
    "Porcentaje_Ret_IVA_2",
    "Valor_Retenido_IVA_2",
    "tipoDocumento",
]

EMITIDOS_RETENCION_FORMA_PAGO_LABEL = {
    "01": "01 - SIN UTILIZACIÓN DEL SISTEMA FINANCIERO",
    "15": "15 - COMPENSACIÓN DE DEUDAS",
    "16": "16 - TARJETA DE DÉBITO",
    "17": "17 - DINERO ELECTRÓNICO",
    "18": "18 - TARJETA PREPAGO",
    "19": "19 - TARJETA DE CRÉDITO",
    "20": "20 - OTROS CON UTILIZACIÓN DEL SISTEMA FINANCIERO",
    "21": "21 - ENDOSO DE TÍTULOS",
}

EMITIDOS_RETENCION_DOC_CODE_LABEL = {
    "01": "01 - FACTURA",
    "03": "03 - LIQUIDACIÓN DE COMPRA",
    "04": "04 - NOTA DE CRÉDITO",
    "05": "05 - NOTA DE DÉBITO",
    "06": "06 - GUÍA DE REMISIÓN",
    "07": "07 - COMPROBANTE DE RETENCIÓN",
}

EMITIDOS_RETENCION_AMBIENTE_LABEL = {
    "1": "1 - Pruebas",
    "2": "2 - Producción",
}

EMITIDOS_RETENCION_TIPO_EMISION_LABEL = {
    "1": "1 - Emisión normal",
    "2": "2 - Emisión por indisponibilidad del sistema",
}

EMITIDOS_RETENCION_TEXT_FORCE_COLUMNS = {
    "rucEmisor",
    "razonSocialEmisor",
    "nombreComercial",
    "direccionMatrizEmisor",
    "direccionSucursalEmisor",
    "obligadoContabilidad",
    "numeroContribuyenteEspecial",
    "numeroAgenteRetencion",
    "fechaAutorizacion",
    "ambiente",
    "emision",
    "numeroComprobante",
    "establecimiento",
    "puntoEmision",
    "secuencial",
    "fechaEmision",
    "razonSocialSujetoRetenido",
    "identificacionSujetoRetenido",
    "claveAcceso",
    "Comprobante_Sustento",
    "Numero_Sustento",
    "Fecha_Emision_Sustento",
    "Ejercicio_Fiscal",
    "Impuesto_Ret_IVA",
    "Impuesto_Ret_IR",
    "informacionAdicional",
    "Impuesto_Ret_IR_1",
    "Impuesto_Ret_IVA_1",
    "Impuesto_Ret_IR_2",
    "Impuesto_Ret_IVA_2",
    "tipoDocumento",
}

EMITIDOS_RETENCION_NUMERIC_COLUMNS = {
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
    "Base_Imponible_Ret_IR_2",
    "Porcentaje_Ret_IR_2",
    "Valor_Retenido_IR_2",
    "Base_Imponible_Ret_IVA_2",
    "Porcentaje_Ret_IVA_2",
    "Valor_Retenido_IVA_2",
}


# =========================================================================== #
# Reportes Emitidos — Notas de Crédito / Débito
# =========================================================================== #
# (Las notas de débito comparten exactamente las mismas columnas que las notas
# de crédito; los aliases conservan esa relación explícitamente.)
EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS = [
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
    "Descripciones",
    "Forma Pago",
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Tarifas IVA",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Código Documento Modificado",
    "Número Documento Modificado",
    "Fecha Emisión Doc. Sustento",
    "Motivo",
    "Valor Modificación",
    "Campos Adicionales",
    "Base Gravada 15%",
    "Monto IVA 15%",
]

EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS = {
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
    "Descripciones",
    "Forma Pago",
    "Código Documento Modificado",
    "Número Documento Modificado",
    "Fecha Emisión Doc. Sustento",
    "Motivo",
    "Campos Adicionales",
}

EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS = {
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
}

EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS = EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS
EMITIDOS_NOTA_DEBITO_TEXT_FORCE_COLUMNS = EMITIDOS_NOTA_CREDITO_TEXT_FORCE_COLUMNS
EMITIDOS_NOTA_DEBITO_NUMERIC_COLUMNS = EMITIDOS_NOTA_CREDITO_NUMERIC_COLUMNS

EMITIDOS_NOTA_CREDITO_TIPO_IDENT_LABEL = {
    "04": "04 - RUC",
    "05": "05 - CÉDULA",
    "06": "06 - PASAPORTE",
    "07": "07 - CONSUMIDOR FINAL",
    "08": "08 - IDENTIFICACIÓN DEL EXTERIOR",
    "09": "09 - PLACA",
}


# =========================================================================== #
# Reportes Emitidos — Factura
# =========================================================================== #
EMITIDOS_FACTURA_REPORT_COLUMNS = [
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
    "Descripciones",
    "Forma Pago",
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Tarifas IVA",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Campos Adicionales",
    "Base No Gravada 0%",
]

EMITIDOS_FACTURA_TEXT_FORCE_COLUMNS = {
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
    "Descripciones",
    "Forma Pago",
    "Campos Adicionales",
    "Tarifas IVA",
}

EMITIDOS_FACTURA_NUMERIC_COLUMNS = {
    "Total Sin Impuestos",
    "Base Gravada",
    "Base No Gravada",
    "Monto IVA",
    "Total Descuento",
    "Propina",
    "Importe Total",
    "Total Pago",
    "Base No Gravada 0%",
}
