"""Los cuatro dialectos de columnas de los reportes mensuales.

El bug que esto protege
-----------------------
El reporte personalizado de retenciones salia con filas completamente vacias
donde habia valores. Causa: el robot deja dos reportes por mes y NO estan
escritos en el mismo dialecto de columnas.

    .../Marzo/PDF/recibidos_reporte_pdf_comprobante_de_retencion_202603.xlsx
        una hoja, camelCase: `rucEmisor`, `claveAcceso`, `fechaEmision`

    .../Marzo/XML/recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx
        lo escribe `construir_reporte` de robot/parser.py: siete hojas
        (Resumen, Cabecera, Detalle, Impuestos, Pagos, Retenciones, 05) con
        titulos legibles: 'RUC del emisor', 'Clave de acceso'

El lector tomaba cualquier hoja de cualquiera de los dos y daba por hecho que
las columnas eran las de la salida. De ese libro multihoja, la unica hoja cuyas
filas sobrevivian al filtro de fecha era '05' --su encabezado 'Fecha de
Emisión' coincide por casualidad con el que se buscaba-- y ninguna de sus
columnas existe en `RETENCION_REPORT_COLUMNS`. El writer, que hace
`df = df[RETENCION_REPORT_COLUMNS]`, las escribia enteras en blanco. Y encima
esa fila vacia se quedaba con la clave de acceso, asi que la fila buena --la
del reporte PDF, que si esta en camelCase-- se descartaba despues por
duplicada.

Uso:
    python tests_reporte_dialectos.py
"""
from __future__ import annotations

import sys
import tempfile
from collections import defaultdict
from datetime import date
from pathlib import Path

import openpyxl

sys.argv = ["aplicacion.py"]
import aplicacion as app  # noqa: E402
from robot.report_dialects import (  # noqa: E402
    _campo_canonico,
    _norm_col,
    _sufijos_de_retencion,
    columnas_de_salida,
    es_confiable,
    filas_de_libro,
)

RUC = "1715857379001"
CLAVE = "0" * 48 + "1"
OTRA_CLAVE = "0" * 48 + "2"

_fallos: list[str] = []


def chequear(etiqueta: str, obtenido, esperado) -> None:
    if obtenido == esperado:
        print(f"  [ok ] {etiqueta}")
        return
    _fallos.append(etiqueta)
    print(f"  [FALLA] {etiqueta}")
    print(f"          esperado: {esperado!r}")
    print(f"          obtenido: {obtenido!r}")


# --------------------------------------------------------------------------- #
# Los encabezados reales de cada dialecto, tal como los escriben los writers
# --------------------------------------------------------------------------- #
CAB = [
    "Archivo XML", "Ruta del archivo XML", "Codigo del documento",
    "Descripcion del comprobante", "Clave de acceso", "Serie del comprobante",
    "Estado de autorizacion", "Numero de autorizacion", "Fecha de autorizacion",
    "Ambiente", "Establecimiento", "Punto de emision", "Secuencial",
    "RUC del emisor", "Razon social del emisor", "Nombre comercial del emisor",
    "Direccion de la matriz", "Direccion del establecimiento",
    "Tipo de identificacion del receptor", "Identificacion del receptor",
    "Nombre o razon social del receptor", "Direccion del receptor",
    "Fecha de emision", "Total sin impuestos", "Total de descuentos",
    "Importe total", "Moneda", "Motivo", "Valor de la modificacion",
    "Valor total", "Total retenido", "Periodo fiscal", "Ejercicio fiscal",
]

RET = [
    "Clave de acceso", "Archivo XML", "Tipo de impuesto", "Codigo de retencion",
    "Base imponible", "Porcentaje a retener", "Valor retenido",
    "Tipo del documento de sustento", "Establecimiento del documento de sustento",
    "Punto de emision del documento de sustento",
    "Secuencial del documento de sustento", "Fecha del documento de sustento",
    "Periodo fiscal",
]

CINCO = [
    "Estado", "Número de Autorización", "Fecha de Autorización", "Ambiente",
    "Razón Social Emisor", "Dir. Establecimiento", "Obligado Contabilidad",
    "Tipo Identificación Comprador", "Identificación Comprador", "Tipo Emisión",
    "Nombre Comercial", "Código del Documento", "Establecimiento",
    "Punto de Emisión", "Secuencial", "Dirección Matriz", "Contribuyente RIMPE",
    "RUC Emisor", "Clave de Acceso", "Fecha de Emisión", "Razón Social Comprador",
    "Dirección Comprador", "Moneda", "Plazo Pago", "Unidad Tiempo Pago",
    "Descripciones", "Forma Pago", "Total Sin Impuestos", "Base Gravada",
    "Base No Gravada", "Tarifas IVA", "Monto IVA", "Total Descuento", "Propina",
    "Importe Total", "Total Pago", "Código Documento Modificado",
    "Número Documento Modificado", "Fecha Emisión Doc. Sustento", "Motivo",
    "Valor Modificación", "Campos Adicionales", "Base Gravada 12%",
    "Monto IVA 12%",
]


def fila_cabecera(clave=CLAVE, secuencial="000000001", fecha="15/03/2026") -> dict:
    return {
        "Clave de acceso": clave,
        "Codigo del documento": "07",
        "Descripcion del comprobante": "Comprobante de Retencion",
        "Serie del comprobante": f"001-001-{secuencial}",
        "Estado de autorizacion": "AUTORIZADO",
        "Numero de autorizacion": clave,
        "Fecha de autorizacion": f"{fecha} 10:00:00",
        "Ambiente": "PRODUCCION",
        "Establecimiento": "001",
        "Punto de emision": "001",
        "Secuencial": secuencial,
        "RUC del emisor": "1790012345001",
        "Razon social del emisor": "EMISOR SA",
        "Nombre comercial del emisor": "EMISOR",
        "Direccion de la matriz": "AV 1",
        "Direccion del establecimiento": "AV 2",
        "Tipo de identificacion del receptor": "RUC",
        "Identificacion del receptor": RUC,
        "Nombre o razon social del receptor": "YO SA",
        "Direccion del receptor": "CALLE 3",
        "Fecha de emision": fecha,
        "Moneda": "DOLAR",
        "Total retenido": 12.5,
        "Periodo fiscal": "03/2026",
        "Ejercicio fiscal": "2026",
    }


def fila_05(clave=CLAVE, secuencial="000000001", fecha="15/03/2026") -> dict:
    return {
        "Estado": "AUTORIZADO",
        "Número de Autorización": clave,
        "Fecha de Autorización": f"{fecha} 10:00:00",
        "Ambiente": "PRODUCCION",
        "Razón Social Emisor": "EMISOR SA",
        "Dir. Establecimiento": "AV 2",
        "Obligado Contabilidad": "SI",
        "Identificación Comprador": RUC,
        "Tipo Emisión": "NORMAL",
        "Nombre Comercial": "EMISOR",
        "Código del Documento": "07",
        "Establecimiento": "001",
        "Punto de Emisión": "001",
        "Secuencial": secuencial,
        "Dirección Matriz": "AV 1",
        "Contribuyente RIMPE": "NINGUNO",
        "RUC Emisor": "1790012345001",
        "Clave de Acceso": clave,
        "Fecha de Emisión": fecha,
        "Razón Social Comprador": "YO SA",
        "Campos Adicionales": "email: contacto@emisor.ec",
    }


def linea_retencion(tipo, codigo, base, porcentaje, valor, clave=CLAVE) -> dict:
    return {
        "Clave de acceso": clave,
        "Tipo de impuesto": tipo,
        "Codigo de retencion": codigo,
        "Base imponible": base,
        "Porcentaje a retener": porcentaje,
        "Valor retenido": valor,
        "Tipo del documento de sustento": "01",
        "Secuencial del documento de sustento": "000000123",
        "Fecha del documento de sustento": "10/03/2026",
        "Periodo fiscal": "03/2026",
    }


def escribir_libro(ruta: Path, hojas: dict[str, tuple[list[str], list[dict]]]) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, (columnas, filas) in hojas.items():
        ws = wb.create_sheet(nombre)
        ws.append(columnas)
        for fila in filas:
            ws.append([fila.get(c) for c in columnas])
    wb.save(ruta)
    wb.close()


def libro_multihoja(
    ruta: Path,
    *,
    cabeceras: list[dict] | None = None,
    lineas: list[dict] | None = None,
    cincos: list[dict] | None = None,
    con_items: bool = True,
) -> None:
    """El libro de `construir_reporte`: siete hojas, titulos legibles."""
    hojas: dict[str, tuple[list[str], list[dict]]] = {
        "Resumen": (
            ["ID de consulta", "Fecha y hora de ejecucion",
             "Rango de fechas consultado", "Total de documentos"],
            [{"ID de consulta": "abc", "Total de documentos": 1}],
        ),
        "Cabecera": (CAB, cabeceras if cabeceras is not None else [fila_cabecera()]),
    }
    if con_items:
        # Hojas de items: una fila por linea, no por comprobante. Si se las
        # confunde con documentos, aparecen como filas vacias en la salida.
        hojas["Detalle"] = (
            ["Clave de acceso", "Archivo XML", "Numero de linea",
             "Codigo principal", "Descripcion del item", "Cantidad"],
            [{"Clave de acceso": CLAVE, "Numero de linea": 1,
              "Codigo principal": "X1", "Descripcion del item": "SERVICIO",
              "Cantidad": 1}],
        )
        hojas["Impuestos"] = (
            ["Nivel del impuesto", "Clave de acceso", "Codigo del impuesto",
             "Codigo de porcentaje", "Tarifa", "Base imponible",
             "Valor del impuesto"],
            [{"Nivel del impuesto": "DOC", "Clave de acceso": CLAVE,
              "Codigo del impuesto": "2", "Tarifa": 15, "Base imponible": 100,
              "Valor del impuesto": 15}],
        )
        hojas["Pagos"] = (
            ["Clave de acceso", "Numero de pago", "Forma de pago (codigo)",
             "Monto del pago", "Plazo", "Unidad de tiempo"],
            [{"Clave de acceso": CLAVE, "Numero de pago": 1,
              "Forma de pago (codigo)": "20", "Monto del pago": 115}],
        )
    if lineas:
        hojas["Retenciones"] = (RET, lineas)
    hojas["05"] = (CINCO, cincos if cincos is not None else [fila_05()])
    escribir_libro(ruta, hojas)


def arbol_retenciones(raiz: Path) -> Path:
    """<raiz>/<RUC>/Recibidos/Comprobante de Retencion/2026/Marzo/"""
    mes = raiz / RUC / "Recibidos" / "Comprobante de Retencion" / "2026" / "Marzo"
    (mes / "XML").mkdir(parents=True, exist_ok=True)
    (mes / "PDF").mkdir(parents=True, exist_ok=True)
    return mes


def generar(raiz: Path, **kwargs) -> dict:
    base = dict(
        ruc=RUC, origen="Recibidos", tipo="Retenciones",
        fecha_inicio=date(2026, 1, 1), fecha_fin=date(2026, 12, 31),
    )
    base.update(kwargs)
    return app._build_custom_report_from_folder(raiz, **base)


def leer_salida(path) -> tuple[list[str], list[dict]]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    if not filas:
        return [], []
    columnas = [str(c).strip() if c is not None else "" for c in filas[0]]
    datos = [dict(zip(columnas, valores)) for valores in filas[1:]]
    return columnas, datos


def _lleno(valor) -> bool:
    return valor is not None and str(valor).strip() not in ("", "0")


# --------------------------------------------------------------------------- #
def probar_normalizacion() -> None:
    """Los cuatro dialectos tienen que converger en el mismo campo."""
    print("un encabezado, cuatro dialectos")
    grupos = [
        ("RUC del emisor", "RUC Emisor", "RUC_EMISOR", "rucEmisor"),
        ("Clave de acceso", "Clave de Acceso", "CLAVE_ACCESO", "claveAcceso"),
        ("Fecha de emision", "Fecha de Emisión", "FECHA_EMISION", "fechaEmision"),
        ("Razon social del emisor", "Razón Social Emisor", "razonSocialEmisor"),
        ("Serie del comprobante", "SERIE_COMPROBANTE", "numeroComprobante"),
        ("Nombre comercial del emisor", "Nombre Comercial", "nombreComercial"),
        ("Direccion de la matriz", "Dirección Matriz", "direccionMatrizEmisor"),
        ("Dir. Establecimiento", "Direccion del establecimiento",
         "direccionSucursalEmisor"),
        ("Nombre o razon social del receptor", "Razón Social Comprador",
         "razonSocialSujetoRetenido", "razonSocialComprador"),
        ("Identificacion del receptor", "Identificación Comprador",
         "identificacionSujetoRetenido", "identificacionComprador"),
        ("Tipo Emisión", "TIPO_EMISION", "emision"),
        ("Campos Adicionales", "informacionAdicional"),
        ("Fecha del documento de sustento", "Fecha Emisión Doc. Sustento",
         "FECHA_EMISION_DOC_SUSTENTO", "Fecha_Emision_Sustento"),
        ("Tipo del documento de sustento", "Comprobante_Sustento"),
        ("Secuencial del documento de sustento", "Numero_Sustento"),
    ]
    for grupo in grupos:
        campos = {_campo_canonico(n) for n in grupo}
        chequear(f"converge: {grupo[0]!r}", len(campos), 1)

    # Y lo que es distinto tiene que seguir siendo distinto.
    distintos = [
        ("Clave de acceso", "Numero de autorizacion"),
        ("Codigo del documento", "Descripcion del comprobante"),
        ("Fecha de emision", "Fecha de autorizacion"),
        ("Base_Imponible_Ret_IVA", "Base_Imponible_Ret_IVA_1"),
        ("Base_Imponible_Ret_IVA", "Base_Imponible_Ret_IR"),
        ("Tipo del documento de sustento", "Codigo del documento"),
    ]
    for a, b in distintos:
        chequear(f"no se confunden: {a!r} / {b!r}",
                 _campo_canonico(a) == _campo_canonico(b), False)

    chequear("el prefijo se respeta", _norm_col("Base Gravada 12%"), "basegravada12")
    chequear("vacio no explota", _norm_col(None), "")
    print()


def probar_sin_colisiones() -> None:
    """Dos columnas de un mismo esquema no pueden caer en el mismo campo.

    Si cayeran, una se comeria a la otra sin aviso.
    """
    print("ningun esquema de salida se pisa a si mismo")
    tipos = ("retenciones", "facturas", "notas_de_credito", "notas_de_debito",
             "liquidacion_de_compra", "guia_de_remision")
    for origen in ("Emitidos", "Recibidos"):
        for tipo in tipos:
            columnas = columnas_de_salida(tipo, origen)
            por_campo: dict[str, list[str]] = defaultdict(list)
            for col in columnas:
                por_campo[_campo_canonico(col)].append(col)
            choques = {k: v for k, v in por_campo.items() if len(v) > 1}
            chequear(f"{origen} · {tipo}", choques, {})
    print()


def probar_sufijos() -> None:
    """Recibidos llega hasta _1; emitidos hasta _2. No se inventan columnas."""
    print("sufijos de retencion por esquema")
    recibidos = columnas_de_salida("retenciones", "Recibidos")
    emitidos = columnas_de_salida("retenciones", "Emitidos")
    chequear("recibidos IVA", _sufijos_de_retencion(recibidos, "IVA"), ["", "_1"])
    chequear("recibidos IR", _sufijos_de_retencion(recibidos, "IR"), ["", "_1"])
    chequear("emitidos IVA", _sufijos_de_retencion(emitidos, "IVA"), ["", "_1", "_2"])
    chequear("emitidos IR", _sufijos_de_retencion(emitidos, "IR"), ["", "_1", "_2"])
    # Una factura no tiene columnas de retencion.
    chequear("facturas no tiene", _sufijos_de_retencion(
        columnas_de_salida("facturas", "Recibidos"), "IVA"), [])
    print()


def probar_libro_multihoja() -> None:
    """El libro de siete hojas traducido al esquema de retenciones."""
    print("el libro multihoja se traduce entero")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = Path(tmp) / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx"
        libro_multihoja(ruta, lineas=[
            linea_retencion("IVA", "2", 100.0, 10.0, 10.0),
            linea_retencion("Renta", "312", 100.0, 2.5, 2.5),
        ])
        columnas = columnas_de_salida("retenciones", "Recibidos")
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        filas, cobertura = filas_de_libro(wb, columnas)
        wb.close()

        chequear("un comprobante, no una fila por item", len(filas), 1)
        chequear("el libro se considera confiable", es_confiable(cobertura), True)
        fila = filas[0]

        # De la hoja Cabecera.
        chequear("clave de acceso", fila["claveAcceso"], CLAVE)
        chequear("fecha de emision", fila["fechaEmision"], "15/03/2026")
        chequear("RUC del emisor", fila["rucEmisor"], "1790012345001")
        chequear("razon social del emisor", fila["razonSocialEmisor"], "EMISOR SA")
        chequear("numero de comprobante", fila["numeroComprobante"], "001-001-000000001")
        chequear("sujeto retenido", fila["razonSocialSujetoRetenido"], "YO SA")
        chequear("identificacion del sujeto", fila["identificacionSujetoRetenido"], RUC)
        chequear("tipo de documento", fila["tipoDocumento"], "Comprobante de Retencion")
        # Solo esta en la hoja '05': prueba que Cabecera y '05' se fusionan.
        chequear("obligado a contabilidad (hoja 05)", fila["obligadoContabilidad"], "SI")
        chequear("tipo de emision (hoja 05)", fila["emision"], "NORMAL")
        chequear("campos adicionales (hoja 05)",
                 fila["informacionAdicional"], "email: contacto@emisor.ec")
        # De la hoja Retenciones.
        chequear("base imponible IVA", fila["Base_Imponible_Ret_IVA"], 100.0)
        chequear("porcentaje IVA", fila["Porcentaje_Ret_IVA"], 10.0)
        chequear("valor retenido IVA", fila["Valor_Retenido_IVA"], 10.0)
        chequear("base imponible renta", fila["Base_Imponible_Ret_IR"], 100.0)
        chequear("porcentaje renta", fila["Porcentaje_Ret_IR"], 2.5)
        chequear("valor retenido renta", fila["Valor_Retenido_IR"], 2.5)
        chequear("comprobante de sustento", fila["Comprobante_Sustento"], "01")
        chequear("numero de sustento", fila["Numero_Sustento"], "000000123")
        chequear("fecha del sustento", fila["Fecha_Emision_Sustento"], "10/03/2026")
    print()


def probar_varias_lineas_del_mismo_impuesto() -> None:
    """Dos retenciones de IVA van a _IVA y a _IVA_1, en ese orden."""
    print("mas de una retencion por impuesto")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = Path(tmp) / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx"
        libro_multihoja(ruta, lineas=[
            linea_retencion("IVA", "2", 100.0, 10.0, 10.0),
            linea_retencion("IVA", "2", 200.0, 20.0, 40.0),
            linea_retencion("Renta", "312", 100.0, 2.5, 2.5),
            # La tercera de IVA no tiene columna en recibidos: se descarta sin
            # romper nada ni desplazar a las otras.
            linea_retencion("IVA", "2", 300.0, 30.0, 90.0),
        ])
        columnas = columnas_de_salida("retenciones", "Recibidos")
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        filas, _ = filas_de_libro(wb, columnas)
        wb.close()
        fila = filas[0]
        chequear("la 1a de IVA sin sufijo", fila["Valor_Retenido_IVA"], 10.0)
        chequear("la 2a de IVA en _1", fila["Valor_Retenido_IVA_1"], 40.0)
        chequear("base de la 2a de IVA", fila["Base_Imponible_Ret_IVA_1"], 200.0)
        chequear("la de renta no se desplaza", fila["Valor_Retenido_IR"], 2.5)
        chequear("renta _1 queda libre", fila["Valor_Retenido_IR_1"], "")
    print()


def probar_hojas_de_items_no_son_documentos() -> None:
    """Detalle, Impuestos y Pagos traen filas por item, no comprobantes."""
    print("las hojas de items no generan filas")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx",
            lineas=[linea_retencion("IVA", "2", 100.0, 10.0, 10.0)],
            con_items=True,
        )
        r = generar(raiz)
        chequear("genera el reporte", r.get("ok"), True)
        chequear("un comprobante, una fila", r.get("rows"), 1)
        _, filas = leer_salida(r["path"])
        vacias = [i for i, f in enumerate(filas)
                  if not any(_lleno(v) for v in f.values())]
        chequear("ninguna fila vacia", vacias, [])
    print()


def probar_regresion_filas_vacias() -> None:
    """El bug tal cual lo vio el usuario, de punta a punta.

    Los dos reportes del mes conviven en el arbol: el multihoja en XML/ y el
    camelCase en PDF/. Antes, el primero ganaba con una fila vacia.
    """
    print("regresion: el reporte del ano no sale en blanco")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx",
            lineas=[
                linea_retencion("IVA", "2", 100.0, 10.0, 10.0),
                linea_retencion("Renta", "312", 100.0, 2.5, 2.5),
            ],
        )
        # El mismo comprobante, en el dialecto de la salida.
        columnas_pdf = columnas_de_salida("retenciones", "Recibidos")
        escribir_libro(
            mes / "PDF" / "recibidos_reporte_pdf_comprobante_de_retencion_202603.xlsx",
            {"Sheet1": (columnas_pdf, [{
                "claveAcceso": CLAVE, "fechaEmision": "15/03/2026",
                "rucEmisor": "1790012345001", "razonSocialEmisor": "EMISOR SA",
                "Valor_Retenido_IVA": 10.0,
            }])},
        )

        r = generar(raiz)
        chequear("genera el reporte", r.get("ok"), True)
        chequear("el comprobante aparece una sola vez", r.get("rows"), 1)
        columnas, filas = leer_salida(r["path"])
        fila = filas[0]
        chequear("la fila NO esta vacia", any(_lleno(v) for v in fila.values()), True)
        chequear("trae la clave de acceso", str(fila.get("claveAcceso")), CLAVE)
        chequear("trae el valor retenido de IVA",
                 float(fila.get("Valor_Retenido_IVA") or 0), 10.0)
        chequear("trae el valor retenido de renta",
                 float(fila.get("Valor_Retenido_IR") or 0), 2.5)
        # Lo que el usuario mira: cuantas columnas quedaron realmente con dato.
        llenas = [c for c in columnas if _lleno(fila.get(c))]
        chequear("mas de 20 columnas con dato", len(llenas) > 20, True)
    print()


def probar_reporte_sin_prefijo_de_origen() -> None:
    """El reporte que queda al lado de las carpetas PDF/TXT/XML.

    Se llama `reporte_comprobante_de_retencion_2026_03.xlsx`, sin prefijo de
    origen, y es el mismo libro de siete hojas. Cuando la descarga no deja el
    `recibidos_reporte_xml_*`, es el unico lugar donde esta el detalle de
    retenciones; el lector lo ignoraba por el nombre.
    """
    print("el reporte sin prefijo de origen tambien se lee")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "reporte_comprobante_de_retencion_2026_03.xlsx",
            lineas=[
                linea_retencion("IVA", "2", 100.0, 10.0, 10.0),
                linea_retencion("Renta", "312", 100.0, 2.5, 2.5),
            ],
        )
        r = generar(raiz)
        chequear("lo encuentra", r.get("ok"), True)
        chequear("una fila", r.get("rows"), 1)
        chequear("y la toma del reporte", r.get("reporte_count"), 1)
        _, filas = leer_salida(r["path"])
        chequear("con el valor retenido de IVA",
                 float(filas[0].get("Valor_Retenido_IVA") or 0), 10.0)
        chequear("con el valor retenido de renta",
                 float(filas[0].get("Valor_Retenido_IR") or 0), 2.5)

        # La salida del modulo no puede confundirse con una fuente.
        r2 = generar(raiz)
        chequear("una segunda corrida no se duplica", r2.get("rows"), 1)
    print()


def probar_reporte_pobre_no_gana() -> None:
    """Un reporte que no llena el esquema no puede tapar al que si.

    Un libro multihoja SIN hoja Retenciones no tiene los valores retenidos:
    como reporte de retenciones esta incompleto. Sus filas quedan
    provisionales y el reporte PDF, que si los trae, le gana.
    """
    print("el reporte incompleto queda de ultimo recurso")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx",
            lineas=None,
        )
        columnas_pdf = columnas_de_salida("retenciones", "Recibidos")
        escribir_libro(
            mes / "PDF" / "recibidos_reporte_pdf_comprobante_de_retencion_202603.xlsx",
            {"Sheet1": (columnas_pdf, [{
                "claveAcceso": CLAVE, "fechaEmision": "15/03/2026",
                "rucEmisor": "1790012345001", "razonSocialEmisor": "DESDE_PDF",
                "Base_Imponible_Ret_IVA": 100.0, "Valor_Retenido_IVA": 10.0,
            }])},
        )
        r = generar(raiz)
        chequear("una sola fila", r.get("rows"), 1)
        _, filas = leer_salida(r["path"])
        chequear("gana el que trae los valores",
                 filas[0].get("razonSocialEmisor"), "DESDE_PDF")
        chequear("con el valor retenido",
                 float(filas[0].get("Valor_Retenido_IVA") or 0), 10.0)
    print()


def probar_provisional_cuando_es_lo_unico() -> None:
    """Pero si no hay nada mejor, el reporte pobre se usa igual.

    Es el caso del modo rapido: el TXT del portal es lo unico que se descarga.
    """
    print("sin otra fuente, el reporte incompleto sirve")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        (mes / "TXT").mkdir(parents=True, exist_ok=True)
        # Los encabezados crudos del portal.
        escribir_libro(
            mes / "TXT" / "recibidos_reporte_txt_comprobante_de_retencion_202603.xlsx",
            {"Comprobantes": (
                ["CLAVE_ACCESO", "RUC_EMISOR", "RAZON_SOCIAL_EMISOR",
                 "SERIE_COMPROBANTE", "FECHA_EMISION", "FECHA_AUTORIZACION"],
                [{"CLAVE_ACCESO": CLAVE, "RUC_EMISOR": "1790012345001",
                  "RAZON_SOCIAL_EMISOR": "EMISOR SA",
                  "SERIE_COMPROBANTE": "001-001-000000001",
                  "FECHA_EMISION": "15/03/2026",
                  "FECHA_AUTORIZACION": "15/03/2026 10:00:00"}],
            )},
        )
        r = generar(raiz)
        chequear("no se queda sin nada", r.get("ok"), True)
        chequear("aporta el comprobante", r.get("rows"), 1)
        _, filas = leer_salida(r["path"])
        chequear("con la clave traducida", str(filas[0].get("claveAcceso")), CLAVE)
        chequear("y el emisor", filas[0].get("razonSocialEmisor"), "EMISOR SA")
    print()


def probar_emitidos_desde_hoja_05() -> None:
    """La hoja '05' es, casi literalmente, el esquema de los emitidos."""
    print("emitidos facturas desde la hoja 05")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = raiz / RUC / "Emitidos" / "Autorizados" / "Factura" / "2026" / "Marzo"
        (mes / "XML").mkdir(parents=True, exist_ok=True)
        cab = fila_cabecera()
        cab["Codigo del documento"] = "01"
        cab["Descripcion del comprobante"] = "Factura"
        cinco = fila_05()
        cinco["Código del Documento"] = "01"
        cinco["Importe Total"] = 115.0
        cinco["Total Sin Impuestos"] = 100.0
        libro_multihoja(
            mes / "XML" / "emitidos_reporte_xml_factura_202603.xlsx",
            cabeceras=[cab], cincos=[cinco],
        )
        r = app._build_custom_report_from_folder(
            raiz, ruc=RUC, origen="Emitidos", tipo="Facturas",
            fecha_inicio=date(2026, 1, 1), fecha_fin=date(2026, 12, 31),
            estado_emitidos="Autorizados",
        )
        chequear("genera el reporte", r.get("ok"), True)
        chequear("una fila", r.get("rows"), 1)
        _, filas = leer_salida(r["path"])
        fila = filas[0]
        chequear("clave de acceso", str(fila.get("Clave de Acceso")), CLAVE)
        chequear("fecha de emision", fila.get("Fecha de Emisión"), "15/03/2026")
        chequear("RUC emisor", str(fila.get("RUC Emisor")), "1790012345001")
        chequear("importe total", float(fila.get("Importe Total") or 0), 115.0)
        chequear("obligado contabilidad", fila.get("Obligado Contabilidad"), "SI")
    print()


def probar_varios_comprobantes() -> None:
    """Un ano entero: cada comprobante con sus propios valores."""
    print("varios comprobantes en un mismo libro")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx",
            cabeceras=[
                fila_cabecera(CLAVE, "000000001", "15/03/2026"),
                fila_cabecera(OTRA_CLAVE, "000000002", "20/03/2026"),
            ],
            cincos=[
                fila_05(CLAVE, "000000001", "15/03/2026"),
                fila_05(OTRA_CLAVE, "000000002", "20/03/2026"),
            ],
            lineas=[
                linea_retencion("IVA", "2", 100.0, 10.0, 10.0, clave=CLAVE),
                linea_retencion("Renta", "312", 500.0, 2.0, 10.0, clave=OTRA_CLAVE),
            ],
        )
        r = generar(raiz)
        chequear("dos filas", r.get("rows"), 2)
        _, filas = leer_salida(r["path"])
        por_clave = {str(f.get("claveAcceso")): f for f in filas}
        chequear("las dos claves estan", sorted(por_clave), sorted([CLAVE, OTRA_CLAVE]))
        chequear("el IVA va al primero",
                 float(por_clave[CLAVE].get("Valor_Retenido_IVA") or 0), 10.0)
        chequear("el primero no hereda renta",
                 float(por_clave[CLAVE].get("Valor_Retenido_IR") or 0), 0.0)
        chequear("la renta va al segundo",
                 float(por_clave[OTRA_CLAVE].get("Valor_Retenido_IR") or 0), 10.0)
        chequear("la base del segundo",
                 float(por_clave[OTRA_CLAVE].get("Base_Imponible_Ret_IR") or 0), 500.0)
        chequear("el segundo no hereda IVA",
                 float(por_clave[OTRA_CLAVE].get("Valor_Retenido_IVA") or 0), 0.0)
        chequear("cada uno con su secuencial",
                 [por_clave[CLAVE].get("secuencial"),
                  por_clave[OTRA_CLAVE].get("secuencial")],
                 ["000000001", "000000002"])
    print()


def probar_rango_de_fechas() -> None:
    """El filtro de fechas sigue aplicando sobre la fila ya traducida."""
    print("el rango de fechas se respeta")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol_retenciones(raiz)
        libro_multihoja(
            mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx",
            cabeceras=[
                fila_cabecera(CLAVE, "000000001", "15/03/2026"),
                fila_cabecera(OTRA_CLAVE, "000000002", "20/09/2026"),
            ],
            cincos=[
                fila_05(CLAVE, "000000001", "15/03/2026"),
                fila_05(OTRA_CLAVE, "000000002", "20/09/2026"),
            ],
            lineas=[
                linea_retencion("IVA", "2", 100.0, 10.0, 10.0, clave=CLAVE),
                linea_retencion("IVA", "2", 100.0, 10.0, 10.0, clave=OTRA_CLAVE),
            ],
        )
        for etiqueta, ini, fin, esperado in (
            ("un dia exacto", date(2026, 3, 15), date(2026, 3, 15), 1),
            ("el mes", date(2026, 3, 1), date(2026, 3, 31), 1),
            ("el ano", date(2026, 1, 1), date(2026, 12, 31), 2),
            ("un mes sin nada", date(2026, 5, 1), date(2026, 5, 31), 0),
        ):
            r = generar(raiz, fecha_inicio=ini, fecha_fin=fin)
            chequear(etiqueta, r.get("rows", 0) if r.get("ok") else 0, esperado)
    print()


if __name__ == "__main__":
    probar_normalizacion()
    probar_sin_colisiones()
    probar_sufijos()
    probar_libro_multihoja()
    probar_varias_lineas_del_mismo_impuesto()
    probar_hojas_de_items_no_son_documentos()
    probar_regresion_filas_vacias()
    probar_reporte_sin_prefijo_de_origen()
    probar_reporte_pobre_no_gana()
    probar_provisional_cuando_es_lo_unico()
    probar_emitidos_desde_hoja_05()
    probar_varios_comprobantes()
    probar_rango_de_fechas()
    if _fallos:
        print(f"FALLARON {len(_fallos)}:")
        for f in _fallos:
            print(f"  - {f}")
        sys.exit(1)
    print("Todo en orden.")
