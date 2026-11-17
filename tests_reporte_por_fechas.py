"""Desde dónde busca documentos el "Reporte por fechas".

El bug: el usuario abría la carpeta del año que ya tenía a la vista
(`G:\\<RUC>\\Emitidos\\Autorizados\\Factura\\2026`), ponía el RUC, y la app
respondía que no encontraba una carpeta con ese nombre adentro. Claro: el RUC
estaba en la ruta, pero como ANCESTRO, y el guard solo aceptaba que fuera la
carpeta misma o una hija directa.

Lo que se protege acá:

  1. Las tres formas de apuntar funcionan: la carpeta del RUC, su padre, y
     cualquier punto más adentro de su árbol.
  2. Una carpeta ajena al RUC se sigue rechazando. Esa restricción existe por
     algo: sin ella, un `base_dir` con varios contribuyentes los mezcla en un
     mismo reporte.
  3. El Excel sale dentro de la carpeta que eligió el usuario, apunte donde
     apunte.

Uso:
    python tests_reporte_por_fechas.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.argv = ["aplicacion.py"]
import aplicacion as app  # noqa: E402

RUC = "1715857379001"
OTRO_RUC = "1792442206001"

_fallos: list[str] = []


def chequear(etiqueta: str, obtenido, esperado) -> None:
    if obtenido == esperado:
        print(f"  [ok ] {etiqueta}")
        return
    _fallos.append(etiqueta)
    print(f"  [FALLA] {etiqueta}")
    print(f"          esperado: {esperado!r}")
    print(f"          obtenido: {obtenido!r}")


def armar_arbol(raiz: Path) -> dict[str, Path]:
    """La estructura que deja el robot: <base>/<RUC>/<origen>/<estado>/..."""
    anio = raiz / RUC / "Emitidos" / "Autorizados" / "Factura" / "2026"
    for mes in ("Enero", "Junio", "Diciembre"):
        (anio / mes / "XML").mkdir(parents=True, exist_ok=True)
        (anio / mes / "PDF").mkdir(parents=True, exist_ok=True)
    # Un segundo contribuyente, para la garantía de no mezclar.
    (raiz / OTRO_RUC / "Emitidos" / "Autorizados").mkdir(parents=True, exist_ok=True)
    (raiz / "sin_relacion").mkdir(parents=True, exist_ok=True)
    return {
        "padre": raiz,
        "ruc": raiz / RUC,
        "origen": raiz / RUC / "Emitidos",
        "tipo": raiz / RUC / "Emitidos" / "Autorizados" / "Factura",
        "anio": anio,
        "mes": anio / "Enero",
        "formato": anio / "Enero" / "XML",
        "otro_ruc": raiz / OTRO_RUC,
        "ajena": raiz / "sin_relacion",
    }


def probar_raiz_de_busqueda() -> None:
    print("desde dónde busca, según a dónde apuntes")
    with tempfile.TemporaryDirectory() as tmp:
        t = armar_arbol(Path(tmp))

        # Las tres formas válidas.
        chequear("la carpeta del RUC", app._raiz_de_busqueda(t["ruc"], RUC), t["ruc"])
        chequear("el padre que la contiene", app._raiz_de_busqueda(t["padre"], RUC), t["ruc"])
        # El bug original: todas estas daban None.
        for nombre in ("origen", "tipo", "anio", "mes", "formato"):
            chequear(
                f"más adentro del árbol: .../{t[nombre].name}",
                app._raiz_de_busqueda(t[nombre], RUC),
                t[nombre],
            )

        # La garantía que el guard protege: no mezclar contribuyentes.
        chequear("una carpeta ajena se rechaza", app._raiz_de_busqueda(t["ajena"], RUC), None)
        chequear(
            "el árbol de OTRO RUC se rechaza",
            app._raiz_de_busqueda(t["otro_ruc"] / "Emitidos", RUC),
            None,
        )
        chequear(
            "apuntando al padre, se elige el RUC pedido y no el otro",
            app._raiz_de_busqueda(t["padre"], OTRO_RUC),
            t["otro_ruc"],
        )

        # Bordes.
        chequear("RUC vacío", app._raiz_de_busqueda(t["ruc"], ""), None)
        chequear("RUC con espacios alrededor", app._raiz_de_busqueda(t["anio"], f"  {RUC} "), t["anio"])
        chequear(
            "una ruta inexistente dentro del árbol del RUC igual se acepta",
            app._raiz_de_busqueda(t["anio"] / "Marzo", RUC),
            t["anio"] / "Marzo",
        )
    print()


def probar_mensaje_de_error() -> None:
    """Rechazar está bien; dejar al usuario sin saber qué hacer, no."""
    print("el mensaje cuando la carpeta no corresponde")
    with tempfile.TemporaryDirectory() as tmp:
        t = armar_arbol(Path(tmp))
        from datetime import date

        r = app._build_custom_report_from_folder(
            t["ajena"], ruc=RUC, origen="Emitidos", tipo="facturas",
            fecha_inicio=date(2026, 1, 1), fecha_fin=date(2026, 12, 31),
            estado_emitidos="Autorizados",
        )
        chequear("rechaza la carpeta ajena", r.get("ok"), False)
        msg = r.get("message") or ""
        chequear("nombra el RUC", RUC in msg, True)
        chequear("dice que se puede apuntar adentro", "dentro de ella" in msg, True)
    print()


def probar_que_pasa_el_guard() -> None:
    """Apuntando adentro, tiene que llegar a barrer, no morir en el guard."""
    print("apuntando al año, llega a buscar documentos")
    with tempfile.TemporaryDirectory() as tmp:
        t = armar_arbol(Path(tmp))
        from datetime import date

        for nombre in ("ruc", "padre", "anio", "mes"):
            r = app._build_custom_report_from_folder(
                t[nombre], ruc=RUC, origen="Emitidos", tipo="facturas",
                fecha_inicio=date(2026, 1, 1), fecha_fin=date(2026, 12, 31),
                estado_emitidos="Autorizados",
            )
            msg = r.get("message") or ""
            # El árbol está vacío de documentos, así que el resultado correcto
            # es "no encontré documentos" -- nunca "no encontré la carpeta".
            chequear(
                f"apuntando a .../{t[nombre].name}: pasa el guard",
                "no pertenece al RUC" in msg,
                False,
            )
    print()


def escribir_reporte(ruta: Path, filas: list[dict]) -> None:
    """Un reporte mensual como los que deja el robot."""
    import openpyxl

    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    columnas = ["Clave de Acceso", "Fecha de Emisión", "Secuencial", "Razón Social Emisor"]
    ws.append(columnas)
    for f in filas:
        ws.append([f.get(c) for c in columnas])
    wb.save(ruta)
    wb.close()


def probar_lectura_de_reportes() -> None:
    """Reusar el reporte mensual en vez de releer los documentos."""
    print("lectura de los reportes mensuales")
    from datetime import date

    with tempfile.TemporaryDirectory() as tmp:
        t = armar_arbol(Path(tmp))
        pdf_dir = t["anio"] / "Junio" / "PDF"

        escribir_reporte(
            pdf_dir / "emitidos_reporte_pdf_factura_202606.xlsx",
            [
                {"Clave de Acceso": "A" * 49, "Fecha de Emisión": "15/06/2026", "Secuencial": "001"},
                {"Clave de Acceso": "B" * 49, "Fecha de Emisión": "20/06/2026", "Secuencial": "002"},
                # Fuera de rango: no debe entrar.
                {"Clave de Acceso": "C" * 49, "Fecha de Emisión": "15/09/2026", "Secuencial": "003"},
                # Fila en blanco, como las que deja el extractor roto.
                {},
            ],
        )
        # Archivo de bloqueo de Excel: tiene que ignorarse.
        (pdf_dir / "~$emitidos_reporte_pdf_factura_202606.xlsx").write_text("x", encoding="utf-8")
        # Un reporte de OTRO tipo no puede colarse en el de facturas.
        escribir_reporte(
            t["ruc"] / "Emitidos" / "Autorizados" / "Comprobante de Retencion" / "2026" / "Junio" / "PDF"
            / "emitidos_reporte_pdf_comprobante_de_retencion_202606.xlsx",
            [{"Clave de Acceso": "D" * 49, "Fecha de Emisión": "16/06/2026", "Secuencial": "004"}],
        )

        r = app._build_custom_report_from_folder(
            t["anio"], ruc=RUC, origen="Emitidos", tipo="facturas",
            fecha_inicio=date(2026, 6, 1), fecha_fin=date(2026, 6, 30),
            estado_emitidos="Autorizados",
        )
        chequear("genera el reporte", r.get("ok"), True)
        chequear("toma las 2 filas en rango", r.get("rows"), 2)
        chequear("y las toma del reporte, no del PDF", r.get("reporte_count"), 2)
        chequear("no abrió ningún PDF", r.get("pdf_count"), 0)
        salida = Path(r["path"])
        chequear(
            "el Excel sale dentro de la carpeta elegida",
            str(salida).startswith(str(t["anio"])),
            True,
        )

        # La salida no se puede comer a sí misma en la corrida siguiente.
        r2 = app._build_custom_report_from_folder(
            t["anio"], ruc=RUC, origen="Emitidos", tipo="facturas",
            fecha_inicio=date(2026, 6, 1), fecha_fin=date(2026, 6, 30),
            estado_emitidos="Autorizados",
        )
        chequear("una segunda corrida no se duplica", r2.get("rows"), 2)
    print()


def probar_prioridad_de_fuentes() -> None:
    """El reporte derivado del XML le gana al derivado del PDF."""
    print("prioridad entre reportes")
    from datetime import date

    with tempfile.TemporaryDirectory() as tmp:
        t = armar_arbol(Path(tmp))
        clave = "E" * 49
        base = t["anio"] / "Diciembre"
        escribir_reporte(
            base / "XML" / "emitidos_reporte_xml_factura_202612.xlsx",
            [{"Clave de Acceso": clave, "Fecha de Emisión": "10/12/2026",
              "Razón Social Emisor": "DESDE_XML"}],
        )
        escribir_reporte(
            base / "PDF" / "emitidos_reporte_pdf_factura_202612.xlsx",
            [{"Clave de Acceso": clave, "Fecha de Emisión": "10/12/2026",
              "Razón Social Emisor": "DESDE_PDF"}],
        )
        r = app._build_custom_report_from_folder(
            t["anio"], ruc=RUC, origen="Emitidos", tipo="facturas",
            fecha_inicio=date(2026, 12, 1), fecha_fin=date(2026, 12, 31),
            estado_emitidos="Autorizados",
        )
        chequear("el comprobante aparece una sola vez", r.get("rows"), 1)

        import openpyxl

        wb = openpyxl.load_workbook(r["path"], read_only=True, data_only=True)
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
        cab = [str(c).strip() if c else "" for c in filas[0]]
        valor = filas[1][cab.index("Razón Social Emisor")] if "Razón Social Emisor" in cab else None
        wb.close()
        chequear("gana el derivado del XML", valor, "DESDE_XML")
    print()


if __name__ == "__main__":
    probar_raiz_de_busqueda()
    probar_mensaje_de_error()
    probar_que_pasa_el_guard()
    probar_lectura_de_reportes()
    probar_prioridad_de_fuentes()
    if _fallos:
        print(f"FALLARON {len(_fallos)}:")
        for f in _fallos:
            print(f"  - {f}")
        sys.exit(1)
    print("Todo en orden.")
