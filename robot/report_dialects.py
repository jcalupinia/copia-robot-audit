"""Traduce los reportes mensuales del robot al esquema del reporte por fechas.

El problema que resuelve
-----------------------
Los ``*.xlsx`` que el robot deja al descargar NO estan todos escritos en el
mismo dialecto de columnas. Conviven cuatro:

``recibidos_reporte_pdf_<tipo>_<YYYYMM>.xlsx``
    Una hoja, camelCase: ``rucEmisor``, ``claveAcceso``, ``fechaEmision``.
    Lo escriben los mismos writers de ``robot/reporting.py`` que arman la
    salida del reporte por fechas, asi que coincide columna por columna.

``recibidos_reporte_xml_<tipo>_<YYYYMM>.xlsx``
    Lo escribe ``construir_reporte`` de ``robot/parser.py``: **varias hojas**
    (Resumen, Cabecera, Detalle, Impuestos, Pagos, Retenciones, 05) y titulos
    legibles sin acentos: ``RUC del emisor``, ``Clave de acceso``.

La hoja ``05`` de ese mismo libro
    Titulos con acentos y otra redaccion: ``RUC Emisor``, ``Clave de Acceso``,
    ``Fecha de Emisión``. Es, casi literalmente, el esquema de los emitidos.

``*_reporte_txt_<tipo>_*.xlsx``
    Los encabezados crudos del portal, tal como vienen en el TXT:
    ``CLAVE_ACCESO``, ``FECHA_EMISION``.

Leerlos como si todos hablaran el dialecto de la salida fue el bug que este
modulo arregla. En un reporte de retenciones, las filas de la hoja ``05``
pasaban el filtro de fecha —su encabezado ``Fecha de Emisión`` coincide por
casualidad con el que se buscaba— pero **ninguna** de sus columnas existe en
``RETENCION_REPORT_COLUMNS``, asi que el writer, que hace
``df = df[RETENCION_REPORT_COLUMNS]``, las escribia enteras en blanco. Y
encima se quedaban con la clave de acceso en el set de deduplicacion, de modo
que la fila buena —la del reporte PDF, que si esta en camelCase— se descartaba
despues por duplicada. Resultado: filas completamente vacias donde habia
valores.

Como lo resuelve
----------------
Un unico eje canonico. Cada encabezado, venga del dialecto que venga, se
reduce a un nombre comparable (``_campo_canonico``); cada columna de la salida
se reduce con la misma funcion. Lo que coincide, se conecta. No hay dos mapas
que puedan desincronizarse: el de entrada y el de salida son el mismo.

Ademas:

* Las hojas que no son de nivel documento (Detalle, Impuestos, Pagos, Resumen,
  Errores_Parsing) se descartan: una fila por item no es un comprobante.
* Cabecera y ``05`` se **fusionan** por clave de acceso en vez de competir.
  Se complementan: ``Obligado Contabilidad``, ``Tipo Emisión`` y los campos
  adicionales solo estan en ``05``.
* La hoja Retenciones trae una fila por impuesto retenido; se **pliega** en las
  columnas ``Base_Imponible_Ret_IVA`` / ``_IR`` / ``_1`` / ``_2`` con la misma
  logica que usa el lector de XML.
* Si un libro no alcanza a llenar el esquema (``_COBERTURA_MINIMA``), sus filas
  quedan **provisionales**: solo se usan si ninguna otra fuente aporto ese
  comprobante. Asi un reporte pobre nunca le gana al XML o al PDF, pero en modo
  rapido —donde el TXT es lo unico que hay— sigue sirviendo.
"""
from __future__ import annotations

import re
import unicodedata

from robot.report_columns import (
    EMITIDOS_FACTURA_REPORT_COLUMNS,
    EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS,
    EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS,
    EMITIDOS_RETENCION_REPORT_COLUMNS,
    PDF_REPORT_COLUMNS,
    RETENCION_REPORT_COLUMNS,
)


# --------------------------------------------------------------------------- #
# Normalizacion de encabezados
# --------------------------------------------------------------------------- #
# 'de', 'del', 'la'... sobran para comparar: 'RUC del emisor' y 'RUC Emisor'
# son la misma columna.
_PALABRAS_VACIAS = frozenset({"de", "del", "la", "el", "los", "las", "y", "o", "a", "al"})


def _norm_col(nombre) -> str:
    """Un encabezado reducido a su forma comparable entre dialectos.

    ``RUC del emisor``, ``RUC Emisor``, ``RUC_EMISOR`` y ``rucEmisor`` dan los
    cuatro ``rucemisor``.
    """
    texto = unicodedata.normalize("NFKD", str(nombre or ""))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    # camelCase -> palabras separadas, antes de bajar todo a minusculas.
    texto = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", texto)
    texto = re.sub(r"[^A-Za-z0-9]+", " ", texto).lower()
    return "".join(p for p in texto.split() if p and p not in _PALABRAS_VACIAS)


# Encabezado normalizado -> campo canonico. Solo hace falta anotar lo que los
# dialectos escriben distinto; lo que ya coincide al normalizar se resuelve
# solo, y lo que no esta en esta tabla queda con su propia forma normalizada
# (asi un encabezado desconocido sigue encontrando a su columna homonima).
_ALIAS: dict[str, str] = {
    # --- identificacion del comprobante ---
    "seriecomprobante": "numerocomprobante",
    "descripcioncomprobante": "tipodocumento",
    "estadoautorizacion": "estado",
    "tipoemision": "emision",
    # --- emisor ---
    "nombrecomercialemisor": "nombrecomercial",
    "direccionmatriz": "direccionmatrizemisor",
    "direstablecimiento": "direccionsucursalemisor",
    "direccionestablecimiento": "direccionsucursalemisor",
    "tipocontribuyenterimpe": "contribuyenterimpe",
    # --- la otra parte: comprador en una factura, sujeto retenido en una
    #     retencion, receptor en el reporte de XML. Es el mismo dato. ---
    # 'Nombre o razon social del receptor' pierde la 'o' al normalizar.
    "nombrerazonsocialreceptor": "razonsocialcomprador",
    "razonsocialreceptor": "razonsocialcomprador",
    "razonsocialsujetoretenido": "razonsocialcomprador",
    "identificacionreceptor": "identificacioncomprador",
    "identificacionsujetoretenido": "identificacioncomprador",
    "tipoidentificacionreceptor": "tipoidentificacioncomprador",
    "direccionreceptor": "direccioncomprador",
    # --- importes ---
    "subtotalsinimpuestos": "totalsinimpuestos",
    "totaldescuentos": "totaldescuento",
    "valortotal": "importetotal",
    # --- documento modificado / de sustento ---
    "numerodocumentomodificado": "comprobantemodificado",
    "razonmodificacion": "motivo",
    "motivomodificacion": "motivo",
    "fechaemisionmodificado": "fechaemisiondocsustento",
    "fechaemisionsustento": "fechaemisiondocsustento",
    "fechadocumentosustento": "fechaemisiondocsustento",
    "fechaemisiondocumentosustento": "fechaemisiondocsustento",
    "tipodocumentosustento": "comprobantesustento",
    "secuencialdocumentosustento": "numerosustento",
    "periodofiscal": "ejerciciofiscal",
    # --- varios ---
    "camposadicionales": "informacionadicional",
    "descripcionesproductos": "descripciones",
}


def _campo_canonico(nombre) -> str:
    norm = _norm_col(nombre)
    return _ALIAS.get(norm, norm)


# --------------------------------------------------------------------------- #
# Esquema de salida
# --------------------------------------------------------------------------- #
def columnas_de_salida(target_tipo: str, origen: str) -> list[str]:
    """Las columnas que va a escribir el reporte por fechas.

    Espeja la eleccion de writer de ``_build_custom_report_from_folder``: si
    una cambia, la otra tiene que cambiar con ella.
    """
    if target_tipo == "retenciones":
        return list(
            EMITIDOS_RETENCION_REPORT_COLUMNS
            if origen == "Emitidos"
            else RETENCION_REPORT_COLUMNS
        )
    if origen == "Emitidos":
        if target_tipo == "notas_de_credito":
            return list(EMITIDOS_NOTA_CREDITO_REPORT_COLUMNS)
        if target_tipo == "notas_de_debito":
            return list(EMITIDOS_NOTA_DEBITO_REPORT_COLUMNS)
        if target_tipo == "facturas":
            return list(EMITIDOS_FACTURA_REPORT_COLUMNS)
    return list(PDF_REPORT_COLUMNS)


def _mapa_de_salida(columnas: list[str]) -> dict[str, str]:
    """campo canonico -> columna de salida."""
    mapa: dict[str, str] = {}
    for col in columnas:
        mapa.setdefault(_campo_canonico(col), col)
    return mapa


# --------------------------------------------------------------------------- #
# Lineas de retencion
# --------------------------------------------------------------------------- #
# La hoja Retenciones trae una fila por impuesto retenido. Se reconoce por
# traer los tres campos de una linea, que ninguna hoja de documentos tiene.
_CAMPOS_LINEA_RETENCION = ("baseimponible", "porcentajeretener", "valorretenido")

_PREFIJOS_RETENCION = {
    "IVA": ("Base_Imponible_Ret_IVA", "Impuesto_Ret_IVA", "Porcentaje_Ret_IVA", "Valor_Retenido_IVA"),
    "IR": ("Base_Imponible_Ret_IR", "Impuesto_Ret_IR", "Porcentaje_Ret_IR", "Valor_Retenido_IR"),
}

_CAMPOS_SUSTENTO = (
    ("Comprobante_Sustento", "comprobantesustento"),
    ("Numero_Sustento", "numerosustento"),
    ("Fecha_Emision_Sustento", "fechaemisiondocsustento"),
    ("Ejercicio_Fiscal", "ejerciciofiscal"),
)


def _sufijos_de_retencion(columnas: list[str], grupo: str) -> list[str]:
    """Los sufijos que el esquema admite para un grupo: ``''``, ``'_1'``...

    Recibidos llega hasta ``_1``; emitidos hasta ``_2``. Lo que no entra en el
    esquema no se inventa.
    """
    base = _PREFIJOS_RETENCION[grupo][0]
    patron = re.compile(rf"^{re.escape(base)}(_(\d+))?$")
    encontrados: list[tuple[int, str]] = []
    for col in columnas:
        m = patron.match(col)
        if m:
            encontrados.append((int(m.group(2) or 0), m.group(1) or ""))
    return [suf for _, suf in sorted(encontrados)]


def _columnas_de_retencion(columnas: list[str]) -> set[str]:
    """Las columnas que una hoja de lineas de retencion puede llenar."""
    alcanzables: set[str] = set()
    for grupo, prefijos in _PREFIJOS_RETENCION.items():
        for suf in _sufijos_de_retencion(columnas, grupo):
            alcanzables.update(f"{p}{suf}" for p in prefijos)
    alcanzables.update(col for col, _ in _CAMPOS_SUSTENTO)
    return alcanzables & set(columnas)


def _vacio(valor) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _plegar_lineas(fila: dict, lineas: list[dict], columnas: list[str]) -> None:
    """Pliega las lineas de retencion de un documento en su fila.

    Misma reparticion que el lector de XML: la primera retencion de IVA va a
    las columnas sin sufijo, la segunda a las ``_1``, y lo mismo para renta.
    """
    if not lineas:
        return
    usados = {"IVA": 0, "IR": 0}
    sufijos = {grupo: _sufijos_de_retencion(columnas, grupo) for grupo in _PREFIJOS_RETENCION}
    presentes = set(columnas)
    for linea in lineas:
        tipo = _norm_col(linea.get("tipoimpuesto", ""))
        grupo = "IVA" if "iva" in tipo else "IR"
        disponibles = sufijos[grupo]
        if usados[grupo] >= len(disponibles):
            continue  # el esquema no tiene mas columnas para este grupo
        suf = disponibles[usados[grupo]]
        usados[grupo] += 1
        base, impuesto, porcentaje, valor = _PREFIJOS_RETENCION[grupo]
        for col, dato in (
            (f"{base}{suf}", linea.get("baseimponible")),
            (f"{impuesto}{suf}", linea.get("tipoimpuesto")),
            (f"{porcentaje}{suf}", linea.get("porcentajeretener")),
            (f"{valor}{suf}", linea.get("valorretenido")),
        ):
            if col in presentes and not _vacio(dato):
                fila[col] = dato
        # El documento de sustento es el mismo para todas las lineas: lo pone
        # la primera que lo traiga.
        for col, campo in _CAMPOS_SUSTENTO:
            if col in presentes and _vacio(fila.get(col)) and not _vacio(linea.get(campo)):
                fila[col] = linea[campo]


# --------------------------------------------------------------------------- #
# Lectura de un libro
# --------------------------------------------------------------------------- #
# Cuanto del esquema tiene que poder llenar un libro para que sus filas valgan
# por si mismas. Por debajo quedan provisionales: se usan solo si ninguna otra
# fuente aporto ese comprobante.
#
# El 0.7 separa con margen los casos medidos:
#
#   1.00  reportes PDF/TXT en camelCase: son el mismo esquema de la salida
#   0.97  libro multihoja -> emitidos factura (la hoja '05' ES ese esquema)
#   0.95  libro multihoja -> retenciones, con su hoja Retenciones
#   ----
#   0.60  libro multihoja -> recibidos factura: le falta todo el desglose de
#         IVA (subtotal15, iva15...), que la Cabecera trae en columnas por
#         tarifa todavia sin mapear
#   0.51  libro multihoja -> retenciones SIN hoja Retenciones: sin los valores
#         retenidos, un reporte de retenciones no esta completo
#
# Quedar provisional no pierde el comprobante: lo aporta el XML o el PDF, que
# son mas fieles. Solo cuesta tiempo de lectura.
_COBERTURA_MINIMA = 0.7

# Una hoja de documentos tiene que reconocer al menos esto, ademas de la
# clave, para no confundirla con una hoja de items.
_MINIMO_COLUMNAS_RECONOCIDAS = 3


def _clave_de_fila(fila: dict) -> str:
    clave = str(fila.get("claveacceso") or "").strip()
    if clave:
        return clave
    partes = [
        str(fila.get(campo) or "").strip()
        for campo in ("numerocomprobante", "fechaemision", "rucemisor")
    ]
    return "|".join(p for p in partes if p)


def filas_de_libro(libro, columnas: list[str]) -> tuple[list[dict], float]:
    """Las filas de un reporte mensual, traducidas al esquema de ``columnas``.

    Devuelve ``(filas, cobertura)``. La cobertura es estructural —que fraccion
    del esquema los encabezados del libro podrian llenar— y no depende de que
    un mes en particular traiga o no retenciones de renta.
    """
    mapa = _mapa_de_salida(columnas)
    documentos: dict[str, dict] = {}
    orden: list[str] = []
    lineas: dict[str, list[dict]] = {}
    alcanzables: set[str] = set()

    for hoja in libro.worksheets:
        filas = hoja.iter_rows(values_only=True)
        try:
            encabezado = next(filas)
        except StopIteration:
            continue
        titulos = [str(c).strip() if c is not None else "" for c in encabezado]
        if not any(titulos):
            continue
        campos = [_campo_canonico(t) if t else "" for t in titulos]
        presentes = {c for c in campos if c}

        es_linea = (
            all(c in presentes for c in _CAMPOS_LINEA_RETENCION)
            and "claveacceso" in presentes
        )
        if es_linea:
            alcanzables |= _columnas_de_retencion(columnas)
        else:
            # Hoja de nivel documento: tiene que identificar al comprobante y
            # aportar algo mas que la clave. Asi quedan afuera Detalle,
            # Impuestos, Pagos, Resumen y Errores_Parsing, cuyas filas no son
            # comprobantes.
            if not presentes & {"claveacceso", "numerocomprobante"}:
                continue
            reconocidas = {mapa[c] for c in presentes if c in mapa}
            if len(reconocidas) < _MINIMO_COLUMNAS_RECONOCIDAS:
                continue
            alcanzables |= reconocidas

        for valores in filas:
            fila: dict = {}
            for campo, valor in zip(campos, valores):
                # Si dos encabezados caen en el mismo campo, gana el primero:
                # en Cabecera, 'Importe total' antes que 'Valor total'.
                if campo and campo not in fila and not _vacio(valor):
                    fila[campo] = valor
            if not fila:
                continue
            clave = _clave_de_fila(fila)
            if not clave:
                continue
            if es_linea:
                lineas.setdefault(clave, []).append(fila)
                continue
            anterior = documentos.get(clave)
            if anterior is None:
                documentos[clave] = fila
                orden.append(clave)
            else:
                # Cabecera y '05' se complementan, no compiten.
                for campo, valor in fila.items():
                    anterior.setdefault(campo, valor)

    resultado: list[dict] = []
    for clave in orden:
        canonico = documentos[clave]
        fila = {col: "" for col in columnas}
        for campo, valor in canonico.items():
            col = mapa.get(campo)
            if col:
                fila[col] = valor
        _plegar_lineas(fila, lineas.get(clave) or [], columnas)
        if all(_vacio(v) for v in fila.values()):
            # Nada reconocido: antes esto se escribia como una fila en blanco.
            continue
        resultado.append(fila)

    cobertura = len(alcanzables) / len(columnas) if columnas else 0.0
    return resultado, cobertura


def es_confiable(cobertura: float) -> bool:
    return cobertura >= _COBERTURA_MINIMA
