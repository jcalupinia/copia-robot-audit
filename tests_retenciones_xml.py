"""Las retenciones leidas del XML: dos esquemas, y que no se dupliquen.

Lo que esto protege, todo encontrado en una muestra real de 100 retenciones
recibidas de 2026:

1. **Dos esquemas en uso a la vez.** 65 de las 100 eran version 1.0.0, con las
   lineas en `impuestos/impuesto`; las otras 35 eran 2.0.0, con las lineas en
   `docsSustento/docSustento/retenciones/retencion`. El parser solo leia el
   1.0.0, asi que 35 comprobantes salian **sin un solo valor retenido**.

2. **El documento de sustento se llama `codDocSustento` / `numDocSustento`.**
   El parser buscaba `tipoDocumento` / `numeroDocumento`, que no existen en
   ninguna version del esquema, asi que las columnas Comprobante_Sustento y
   Numero_Sustento salian **siempre vacias**, en los 100.

3. **`infoAdicional` cuelga del nodo raiz**, no de infoFactura /
   infoCompRetencion. Buscarlo dentro del bloque infoXxx lo dejaba siempre
   vacio, para todos los tipos de documento.

4. **Duplicacion.** El bucle que lee los XML sumaba la clave de acceso al set
   de deduplicacion pero nunca la consultaba antes. Cuando los reportes
   mensuales empezaron a aportar filas, cada comprobante salia dos veces: el
   reporte del ano daba 199 filas para 100 comprobantes.

5. **El XML le gana al reporte derivado del XML.** `construir_reporte` escribe
   etiquetas en vez de codigos ("1 - Emision normal" por "1") y rellena los
   huecos con el literal "No Disponible", que despues tapa el valor real.

Uso:
    python tests_retenciones_xml.py
"""
from __future__ import annotations

import sys
import tempfile
from datetime import date
from pathlib import Path

import openpyxl

sys.argv = ["aplicacion.py"]
import aplicacion as app  # noqa: E402
from robot.parser import _parse_recibido_xml  # noqa: E402
from robot.report_dialects import columnas_de_salida  # noqa: E402

RUC = "1715857379001"
CLAVE = "0" * 48 + "1"

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
# Fixtures: retenciones en los dos esquemas que el SRI tiene en circulacion
# --------------------------------------------------------------------------- #
_INFO_TRIBUTARIA = f"""  <infoTributaria>
    <ambiente>2</ambiente>
    <tipoEmision>1</tipoEmision>
    <razonSocial>EMISOR SA</razonSocial>
    <nombreComercial>EMISOR</nombreComercial>
    <ruc>1790012345001</ruc>
    <claveAcceso>{CLAVE}</claveAcceso>
    <codDoc>07</codDoc>
    <estab>001</estab>
    <ptoEmi>019</ptoEmi>
    <secuencial>000000001</secuencial>
    <dirMatriz>AV 1</dirMatriz>
    <agenteRetencion>1</agenteRetencion>
  </infoTributaria>
"""

_INFO_COMP_RETENCION = f"""  <infoCompRetencion>
    <fechaEmision>15/03/2026</fechaEmision>
    <dirEstablecimiento>AV 2</dirEstablecimiento>
    <contribuyenteEspecial>5368</contribuyenteEspecial>
    <obligadoContabilidad>SI</obligadoContabilidad>
    <tipoIdentificacionSujetoRetenido>04</tipoIdentificacionSujetoRetenido>
    <razonSocialSujetoRetenido>YO SA</razonSocialSujetoRetenido>
    <identificacionSujetoRetenido>{RUC}</identificacionSujetoRetenido>
    <periodoFiscal>03/2026</periodoFiscal>
  </infoCompRetencion>
"""

# `infoAdicional` va aca: hermano del bloque infoXxx, hijo del nodo raiz.
_INFO_ADICIONAL = """  <infoAdicional>
    <campoAdicional nombre="DIRECCION">CALLE 3 Y AV 4</campoAdicional>
    <campoAdicional nombre="EMAIL">contacto@emisor.ec</campoAdicional>
  </infoAdicional>
"""

RETENCION_V1 = f"""<?xml version="1.0" encoding="UTF-8"?>
<comprobanteRetencion id="comprobante" version="1.0.0">
{_INFO_TRIBUTARIA}{_INFO_COMP_RETENCION}  <impuestos>
    <impuesto>
      <codigo>2</codigo>
      <codigoRetencion>1</codigoRetencion>
      <baseImponible>100.00</baseImponible>
      <porcentajeRetener>30.00</porcentajeRetener>
      <valorRetenido>10.00</valorRetenido>
      <codDocSustento>22</codDocSustento>
      <numDocSustento>000000000260418</numDocSustento>
      <fechaEmisionDocSustento>10/03/2026</fechaEmisionDocSustento>
    </impuesto>
    <impuesto>
      <codigo>1</codigo>
      <codigoRetencion>344A</codigoRetencion>
      <baseImponible>200.00</baseImponible>
      <porcentajeRetener>2.00</porcentajeRetener>
      <valorRetenido>4.00</valorRetenido>
      <codDocSustento>22</codDocSustento>
      <numDocSustento>000000000260418</numDocSustento>
      <fechaEmisionDocSustento>10/03/2026</fechaEmisionDocSustento>
    </impuesto>
  </impuestos>
{_INFO_ADICIONAL}</comprobanteRetencion>
"""

RETENCION_V2 = f"""<?xml version="1.0" encoding="UTF-8"?>
<comprobanteRetencion id="comprobante" version="2.0.0">
{_INFO_TRIBUTARIA}{_INFO_COMP_RETENCION}  <docsSustento>
    <docSustento>
      <codSustento>01</codSustento>
      <codDocSustento>01</codDocSustento>
      <numDocSustento>001100000000142</numDocSustento>
      <fechaEmisionDocSustento>01/03/2026</fechaEmisionDocSustento>
      <totalSinImpuestos>570.90</totalSinImpuestos>
      <importeTotal>759.00</importeTotal>
      <impuestosDocSustento>
        <impuestoDocSustento>
          <codImpuestoDocSustento>2</codImpuestoDocSustento>
          <codigoPorcentaje>4</codigoPorcentaje>
          <baseImponible>660.00</baseImponible>
          <tarifa>15.00</tarifa>
          <valorImpuesto>99.00</valorImpuesto>
        </impuestoDocSustento>
      </impuestosDocSustento>
      <retenciones>
        <retencion>
          <codigo>2</codigo>
          <codigoRetencion>1</codigoRetencion>
          <baseImponible>99.00</baseImponible>
          <porcentajeRetener>30.00</porcentajeRetener>
          <valorRetenido>29.70</valorRetenido>
        </retencion>
        <retencion>
          <codigo>1</codigo>
          <codigoRetencion>309</codigoRetencion>
          <baseImponible>660.00</baseImponible>
          <porcentajeRetener>3.00</porcentajeRetener>
          <valorRetenido>19.80</valorRetenido>
        </retencion>
      </retenciones>
      <pagos>
        <pago>
          <formaPago>20</formaPago>
          <total>759.00</total>
        </pago>
      </pagos>
    </docSustento>
  </docsSustento>
{_INFO_ADICIONAL}</comprobanteRetencion>
"""


def escribir_xml(ruta: Path, contenido: str) -> Path:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(contenido, encoding="utf-8")
    return ruta


def arbol(raiz: Path, mes: str = "Marzo") -> Path:
    """<raiz>/<RUC>/Recibidos/Comprobante de Retencion/2026/<Mes>/"""
    d = raiz / RUC / "Recibidos" / "Comprobante de Retencion" / "2026" / mes
    (d / "XML").mkdir(parents=True, exist_ok=True)
    (d / "PDF").mkdir(parents=True, exist_ok=True)
    return d


def generar(raiz: Path, **kwargs) -> dict:
    base = dict(
        ruc=RUC, origen="Recibidos", tipo="Retenciones",
        fecha_inicio=date(2026, 1, 1), fecha_fin=date(2026, 12, 31),
    )
    base.update(kwargs)
    return app._build_custom_report_from_folder(raiz, **base)


def leer_salida(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    cols = [str(c).strip() if c is not None else "" for c in filas[0]]
    return cols, [dict(zip(cols, v)) for v in filas[1:]]


# --------------------------------------------------------------------------- #
def probar_esquema_1_0_0() -> None:
    """impuestos/impuesto, con el sustento repetido en cada linea."""
    print("retencion esquema 1.0.0")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = escribir_xml(Path(tmp) / "ret_v1.xml", RETENCION_V1)
        cab, _, _, _, _, ret, err, _ = _parse_recibido_xml(ruta)
        chequear("sin error de parseo", err, None)
        chequear("dos lineas de retencion", len(ret), 2)
        chequear("el total retenido suma", round(cab["TOTAL_RETENIDO"], 2), 14.0)
        iva = next(r for r in ret if "iva" in r["TIPO_IMPUESTO"].lower())
        renta = next(r for r in ret if "iva" not in r["TIPO_IMPUESTO"].lower())
        chequear("base del IVA", iva["BASE_IMPONIBLE"], 100.0)
        chequear("valor retenido de IVA", iva["VALOR_RETENIDO"], 10.0)
        chequear("valor retenido de renta", renta["VALOR_RETENIDO"], 4.0)
        # Lo que estaba roto: codDocSustento / numDocSustento.
        chequear("tipo del documento de sustento", iva["DOC_SUSTENTO_TIPO"], "22")
        chequear("numero del documento de sustento",
                 iva["DOC_SUSTENTO_SECUENCIAL"], "000000000260418")
        chequear("fecha del documento de sustento",
                 iva["FECHA_EMISION_DOC_SUSTENTO"], "10/03/2026")
    print()


def probar_esquema_2_0_0() -> None:
    """docsSustento/docSustento/retenciones/retencion.

    35 de 100 retenciones reales usan este esquema y salian todas en cero.
    """
    print("retencion esquema 2.0.0")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = escribir_xml(Path(tmp) / "ret_v2.xml", RETENCION_V2)
        cab, _, _, _, _, ret, err, _ = _parse_recibido_xml(ruta)
        chequear("sin error de parseo", err, None)
        chequear("dos lineas de retencion", len(ret), 2)
        chequear("el total retenido suma", round(cab["TOTAL_RETENIDO"], 2), 49.5)
        iva = next(r for r in ret if "iva" in r["TIPO_IMPUESTO"].lower())
        renta = next(r for r in ret if "iva" not in r["TIPO_IMPUESTO"].lower())
        chequear("base del IVA", iva["BASE_IMPONIBLE"], 99.0)
        chequear("valor retenido de IVA", iva["VALOR_RETENIDO"], 29.7)
        chequear("codigo de retencion de renta", renta["CODIGO_RETENCION"], "309")
        chequear("valor retenido de renta", renta["VALOR_RETENIDO"], 19.8)
        # El sustento es del docSustento que las agrupa, no de cada linea.
        chequear("tipo del documento de sustento", iva["DOC_SUSTENTO_TIPO"], "01")
        chequear("numero del documento de sustento",
                 iva["DOC_SUSTENTO_SECUENCIAL"], "001100000000142")
        chequear("fecha del documento de sustento",
                 iva["FECHA_EMISION_DOC_SUSTENTO"], "01/03/2026")
        chequear("la renta hereda el mismo sustento",
                 renta["DOC_SUSTENTO_SECUENCIAL"], "001100000000142")
        # Las lineas del impuesto del documento de sustento no son retenciones.
        chequear("impuestoDocSustento no se cuenta como retencion", len(ret), 2)
    print()


def probar_varios_doc_sustento() -> None:
    """Un 2.0.0 con dos documentos de sustento: cada uno con su numero."""
    print("esquema 2.0.0 con varios documentos de sustento")
    segundo = """    <docSustento>
      <codSustento>01</codSustento>
      <codDocSustento>18</codDocSustento>
      <numDocSustento>002200000000999</numDocSustento>
      <fechaEmisionDocSustento>05/03/2026</fechaEmisionDocSustento>
      <retenciones>
        <retencion>
          <codigo>1</codigo>
          <codigoRetencion>312</codigoRetencion>
          <baseImponible>50.00</baseImponible>
          <porcentajeRetener>1.75</porcentajeRetener>
          <valorRetenido>0.88</valorRetenido>
        </retencion>
      </retenciones>
    </docSustento>
  </docsSustento>"""
    contenido = RETENCION_V2.replace("  </docsSustento>", segundo)
    with tempfile.TemporaryDirectory() as tmp:
        ruta = escribir_xml(Path(tmp) / "ret_v2_doble.xml", contenido)
        cab, _, _, _, _, ret, err, _ = _parse_recibido_xml(ruta)
        chequear("sin error de parseo", err, None)
        chequear("tres lineas en total", len(ret), 3)
        chequear("el total retenido suma las tres",
                 round(cab["TOTAL_RETENIDO"], 2), 50.38)
        numeros = sorted({r["DOC_SUSTENTO_SECUENCIAL"] for r in ret})
        chequear("cada linea con el sustento de su documento",
                 numeros, ["001100000000142", "002200000000999"])
    print()


def probar_info_adicional() -> None:
    """`infoAdicional` cuelga del nodo raiz, no del bloque infoXxx."""
    print("campos adicionales")
    with tempfile.TemporaryDirectory() as tmp:
        for nombre, contenido in (("v1", RETENCION_V1), ("v2", RETENCION_V2)):
            ruta = escribir_xml(Path(tmp) / f"ret_{nombre}.xml", contenido)
            cab, _, _, _, adic, _, _, _ = _parse_recibido_xml(ruta)
            chequear(f"{nombre}: los lee", len(adic), 2)
            chequear(f"{nombre}: el JSON no queda vacio",
                     bool((cab.get("INFO_ADICIONAL_JSON") or "").strip()), True)
            texto = app._texto_info_adicional(cab)
            chequear(f"{nombre}: legible, no JSON crudo",
                     texto, "DIRECCION: CALLE 3 Y AV 4\nEMAIL: contacto@emisor.ec")
        # Entradas raras no pueden tumbar el reporte.
        chequear("sin campos adicionales", app._texto_info_adicional({}), "")
        chequear("JSON invalido se devuelve tal cual",
                 app._texto_info_adicional({"INFO_ADICIONAL_JSON": "{roto"}), "{roto")
    print()


def probar_contribuyente_especial_y_agente() -> None:
    """Dos columnas que el XML traia y el mapeo ignoraba."""
    print("contribuyente especial y agente de retencion")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = escribir_xml(Path(tmp) / "ret_v1.xml", RETENCION_V1)
        cab, _, _, _, _, ret, _, _ = _parse_recibido_xml(ruta)
        chequear("el parser lee contribuyenteEspecial",
                 cab.get("CONTRIBUYENTE_ESPECIAL"), "5368")
        chequear("el parser lee agenteRetencion", cab.get("AGENTE_RETENCION"), "1")
        fila = app._map_xml_to_retencion_row(cab, ret)
        chequear("llega a la columna de retenciones",
                 fila.get("numeroContribuyenteEspecial"), "5368")
        chequear("y el agente de retencion",
                 fila.get("numeroAgenteRetencion"), "1")
    print()


def probar_sin_duplicados() -> None:
    """El bug de las 199 filas para 100 comprobantes.

    El reporte mensual y el XML hablan del mismo comprobante; tiene que salir
    una sola fila.
    """
    print("el mismo comprobante no sale dos veces")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol(raiz)
        escribir_xml(mes / "XML" / "comprobante_de_retencion__20260315__EMISOR.xml",
                     RETENCION_V1)
        # El reporte del mes, en el dialecto de la salida, con el mismo
        # comprobante.
        columnas = columnas_de_salida("retenciones", "Recibidos")
        ruta_rep = mes / "PDF" / "recibidos_reporte_pdf_comprobante_de_retencion_202603.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(columnas)
        ws.append([{"claveAcceso": CLAVE, "fechaEmision": "15/03/2026",
                    "rucEmisor": "1790012345001",
                    "razonSocialEmisor": "DESDE_REPORTE"}.get(c) for c in columnas])
        wb.save(ruta_rep)
        wb.close()

        r = generar(raiz)
        chequear("genera el reporte", r.get("ok"), True)
        chequear("una sola fila", r.get("rows"), 1)
        chequear("la aporta el XML", r.get("xml_count"), 1)
        chequear("no la aporta el reporte", r.get("reporte_count"), 0)
        _, filas = leer_salida(r["path"])
        chequear("gana el XML", filas[0].get("razonSocialEmisor"), "EMISOR SA")
    print()


def probar_xml_le_gana_al_reporte_derivado() -> None:
    """El reporte derivado del XML trae etiquetas y 'No Disponible'.

    `construir_reporte` escribe '1 - Emision normal' donde el XML dice '1', y
    rellena los huecos con el literal 'No Disponible'. Mientras el XML este en
    disco, es mas fiel leerlo.
    """
    print("el XML le gana al reporte derivado del XML")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol(raiz)
        escribir_xml(mes / "XML" / "comprobante_de_retencion__20260315__EMISOR.xml",
                     RETENCION_V1)
        # Un libro multihoja como el de construir_reporte, con sus etiquetas.
        ruta = mes / "XML" / "recibidos_reporte_xml_comprobante_de_retencion_202603.xlsx"
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet("Cabecera")
        ws.append(["Clave de acceso", "Fecha de emision", "RUC del emisor",
                   "Razon social del emisor", "Serie del comprobante"])
        ws.append([CLAVE, "15/03/2026", "1790012345001", "EMISOR SA",
                   "001-019-000000001"])
        ws = wb.create_sheet("05")
        ws.append(["Clave de Acceso", "Fecha de Emisión", "Tipo Emisión",
                   "Obligado Contabilidad", "Campos Adicionales"])
        ws.append([CLAVE, "15/03/2026", "1 - Emision normal", "SI", "No Disponible"])
        wb.save(ruta)
        wb.close()

        r = generar(raiz)
        chequear("una sola fila", r.get("rows"), 1)
        chequear("la aporta el XML", r.get("xml_count"), 1)
        _, filas = leer_salida(r["path"])
        fila = filas[0]
        chequear("el codigo de emision, no la etiqueta", fila.get("emision"), "1")
        chequear("sin el literal 'No Disponible'",
                 "No Disponible" in str(fila.get("informacionAdicional") or ""), False)
        chequear("con los campos adicionales de verdad",
                 "contacto@emisor.ec" in str(fila.get("informacionAdicional") or ""),
                 True)
        chequear("y con el valor retenido",
                 float(fila.get("Valor_Retenido_IVA") or 0), 10.0)
    print()


def probar_reporte_cuando_no_hay_xml() -> None:
    """Pero si el XML no esta en disco, el reporte sigue siendo la fuente."""
    print("sin XML en disco, manda el reporte")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol(raiz, "Abril")
        columnas = columnas_de_salida("retenciones", "Recibidos")
        ruta_rep = mes / "PDF" / "recibidos_reporte_pdf_comprobante_de_retencion_202604.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(columnas)
        ws.append([{"claveAcceso": CLAVE, "fechaEmision": "15/04/2026",
                    "rucEmisor": "1790012345001",
                    "razonSocialEmisor": "DESDE_REPORTE",
                    "Valor_Retenido_IVA": 7.5}.get(c) for c in columnas])
        wb.save(ruta_rep)
        wb.close()
        r = generar(raiz)
        chequear("genera el reporte", r.get("ok"), True)
        chequear("una fila", r.get("rows"), 1)
        chequear("la aporta el reporte", r.get("reporte_count"), 1)
        _, filas = leer_salida(r["path"])
        chequear("con su valor", float(filas[0].get("Valor_Retenido_IVA") or 0), 7.5)
    print()


def probar_los_dos_esquemas_en_un_reporte() -> None:
    """Un ano real mezcla las dos versiones; las dos tienen que traer montos."""
    print("los dos esquemas en el mismo reporte")
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        mes = arbol(raiz)
        escribir_xml(mes / "XML" / "comprobante_de_retencion__20260315__V1.xml",
                     RETENCION_V1)
        otra = "0" * 48 + "2"
        escribir_xml(mes / "XML" / "comprobante_de_retencion__20260315__V2.xml",
                     RETENCION_V2.replace(CLAVE, otra).replace(
                         "<secuencial>000000001</secuencial>",
                         "<secuencial>000000002</secuencial>"))
        r = generar(raiz)
        chequear("dos filas", r.get("rows"), 2)
        _, filas = leer_salida(r["path"])
        por_clave = {str(f.get("claveAcceso")): f for f in filas}
        chequear("las dos claves", sorted(por_clave), sorted([CLAVE, otra]))
        for clave, esperado_iva in ((CLAVE, 10.0), (otra, 29.7)):
            chequear(f"...{clave[-1]}: valor retenido de IVA",
                     float(por_clave[clave].get("Valor_Retenido_IVA") or 0),
                     esperado_iva)
            chequear(f"...{clave[-1]}: trae numero de sustento",
                     bool(str(por_clave[clave].get("Numero_Sustento") or "").strip()),
                     True)
    print()


if __name__ == "__main__":
    probar_esquema_1_0_0()
    probar_esquema_2_0_0()
    probar_varios_doc_sustento()
    probar_info_adicional()
    probar_contribuyente_especial_y_agente()
    probar_sin_duplicados()
    probar_xml_le_gana_al_reporte_derivado()
    probar_reporte_cuando_no_hay_xml()
    probar_los_dos_esquemas_en_un_reporte()
    if _fallos:
        print(f"FALLARON {len(_fallos)}:")
        for f in _fallos:
            print(f"  - {f}")
        sys.exit(1)
    print("Todo en orden.")
