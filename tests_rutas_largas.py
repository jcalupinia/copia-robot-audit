"""Las rutas largas de Windows, que es como se perdian las descargas.

El sintoma era el peor posible: el proceso "termina bien" y no baja un solo
archivo. La causa es que Windows corta las rutas en 260 caracteres salvo que el
equipo tenga LongPathsEnabled en 1, y el robot llega ahi sin que la ruta sea
rara -- el usuario elige una carpeta descriptiva, el robot le agrega 75
caracteres de estructura y el nombre lleva la clave de acceso de 49 digitos.

Lo que se protege aca:

  1. `normalizar_carpeta_destino` devuelve la ruta con el prefijo `\\?\`, que
     es lo que levanta el limite, y lo heredan todas las rutas que se cuelgan
     de ella.
  2. Lo que se le guarda y se le muestra al usuario NO lleva el prefijo.
  3. Con el prefijo no se acorta ningun nombre: el archivo conserva la fecha y
     la serie.
  4. Sin el prefijo el recorte sigue siendo la red de seguridad, y lo que
     recorta nunca es el prefijo del tipo, porque es por ahi que el reporte
     reconoce los archivos.

Uso:
    python tests_rutas_largas.py
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

from robot.browser import _acortar_si_no_entra, _resolver_destino_unico
from robot.file_utils import (
    PREFIJO_RUTA_LARGA,
    mostrar_ruta,
    normalizar_carpeta_destino,
    ruta_larga,
)
from robot.workflows import _nombre_documento_mes, _xml_files_por_tipo

TIPO = "Comprobante de Retencion"
RUC = "1792442206001"
CLAVE = "0401202607179244220600120010010010000890000008900189"
NOMBRE = _nombre_documento_mes(TIPO, "20260104", f"001-050-{CLAVE}")

RAIZ = Path(os.environ.get("TEMP", "c:/tmp")) / "tests_rutas_largas"

_fallos: list[str] = []


def chequear(etiqueta: str, obtenido, esperado) -> None:
    if obtenido == esperado:
        print(f"  [ok ] {etiqueta}")
        return
    _fallos.append(etiqueta)
    print(f"  [FALLA] {etiqueta}")
    print(f"          esperado: {esperado!r}")
    print(f"          obtenido: {obtenido!r}")


def afirmar(etiqueta: str, condicion: bool) -> None:
    chequear(etiqueta, bool(condicion), True)


def carpeta_de(largo: int) -> str:
    """Una ruta real bajo TEMP del largo pedido, para no depender del equipo."""
    p = RAIZ
    tramo = "CARPETA CON NOMBRE DESCRIPTIVO"
    while len(str(p)) < largo - len(tramo) - 1:
        p = p / tramo
    relleno = max(1, largo - len(str(p)) - 1)
    return str(p / ("R" * relleno))


# --------------------------------------------------------------------------- #
# 1. El contrato de los tres helpers
# --------------------------------------------------------------------------- #
def probar_helpers() -> None:
    print("ruta_larga / mostrar_ruta / normalizar_carpeta_destino")
    solo_windows = os.name == "nt"

    if solo_windows:
        afirmar(
            "normalizar_carpeta_destino aplica el prefijo",
            str(normalizar_carpeta_destino("C:/SRI")).startswith(PREFIJO_RUTA_LARGA),
        )
        afirmar("ruta_larga es idempotente", ruta_larga(ruta_larga("C:/x")) == ruta_larga("C:/x"))
        afirmar(
            "ruta_larga resuelve '..' (el prefijo ya no lo interpreta)",
            ".." not in str(ruta_larga("C:/a/b/../c")),
        )
        afirmar(
            "ruta_larga marca los recursos de red como UNC",
            "UNC" in str(ruta_larga(f"{os.sep*2}servidor{os.sep}compartida")),
        )
        chequear(
            "mostrar_ruta deshace el prefijo",
            mostrar_ruta(ruta_larga("C:/SRI/Enero")),
            str(Path("C:/SRI/Enero")),
        )
        chequear(
            "mostrar_ruta deshace el UNC",
            mostrar_ruta(ruta_larga(f"{os.sep*2}srv{os.sep}comp{os.sep}Clientes")),
            f"{os.sep*2}srv{os.sep}comp{os.sep}Clientes",
        )
    else:
        afirmar(
            "fuera de Windows no hay prefijo que aplicar",
            not str(ruta_larga("/tmp/x")).startswith(PREFIJO_RUTA_LARGA),
        )

    # Lo de antes tiene que seguir andando: Windows no admite que un nombre de
    # carpeta termine en espacio o punto.
    afirmar(
        "espacio al final se sigue recortando",
        not mostrar_ruta(normalizar_carpeta_destino("C:/Cliente S.A. ")).endswith(" "),
    )
    afirmar(
        "punto al final se sigue recortando",
        not mostrar_ruta(normalizar_carpeta_destino("C:/Datos.")).endswith("."),
    )
    for vacio in ("", None):
        try:
            normalizar_carpeta_destino(vacio)
            print(f"  [ok ] {vacio!r} no explota")
        except Exception as err:  # pragma: no cover
            _fallos.append(f"{vacio!r} explota")
            print(f"  [FALLA] {vacio!r} explota: {err}")
    chequear("mostrar_ruta(None)", mostrar_ruta(None), "")
    print()


# --------------------------------------------------------------------------- #
# 2. El flujo real: elegir -> guardar -> usar -> bajar -> reporte
# --------------------------------------------------------------------------- #
def probar_flujo() -> None:
    print("flujo completo con rutas que se pasan de 260")
    shutil.rmtree(str(ruta_larga(RAIZ)), ignore_errors=True)
    try:
        for etiqueta, elegida in (
            ("carpeta base de 129 caracteres", carpeta_de(129)),
            ("carpeta base de 204 caracteres", carpeta_de(204)),
            ("carpeta con un espacio al final", carpeta_de(120) + " "),
        ):
            # El usuario elige la carpeta y la app la guarda.
            nueva = normalizar_carpeta_destino(elegida)
            nueva.mkdir(parents=True, exist_ok=True)
            guardado = mostrar_ruta(nueva)
            afirmar(
                f"{etiqueta}: se guarda sin el prefijo",
                not guardado.startswith(PREFIJO_RUTA_LARGA),
            )

            # La app la vuelve a leer y baja un comprobante.
            base = normalizar_carpeta_destino(guardado)
            carpeta = (
                base / RUC / "Emitidos" / "Autorizados" / TIPO / "2026" / "Enero" / "XML"
            )
            carpeta.mkdir(parents=True, exist_ok=True)
            destino = _resolver_destino_unico(carpeta / NOMBRE, ".xml")
            destino.write_bytes(b"<autorizacion/>")

            afirmar(f"{etiqueta}: el archivo queda en disco", destino.exists())
            chequear(
                f"{etiqueta}: conserva el nombre completo", destino.name, f"{NOMBRE}.xml"
            )
            chequear(
                f"{etiqueta}: el reporte lo encuentra",
                len(_xml_files_por_tipo(base, "")),
                1,
            )
            if os.name == "nt":
                afirmar(
                    f"{etiqueta}: la ruta se pasaba de 260",
                    len(mostrar_ruta(destino)) > 260,
                )
    finally:
        shutil.rmtree(str(ruta_larga(RAIZ)), ignore_errors=True)
    print()


# --------------------------------------------------------------------------- #
# 3. La red de seguridad: recortar sin romper el prefijo del tipo
# --------------------------------------------------------------------------- #
def probar_recorte() -> None:
    print("recorte de nombres (para rutas sin prefijo)")
    tipo_prefijo = "comprobante_de_retencion__"

    def recortar(largo_carpeta: int) -> str:
        carpeta = Path("C:/" + "x" * (largo_carpeta - 3))
        return _acortar_si_no_entra(carpeta / NOMBRE, ".xml").name

    afirmar("con lugar de sobra no toca el nombre", recortar(80) == NOMBRE)

    nombre = recortar(204)
    afirmar("cuando no entra, recorta", nombre != NOMBRE)
    afirmar(
        "lo recortado conserva el prefijo del tipo, que usa el reporte",
        nombre.startswith(tipo_prefijo),
    )
    afirmar("lo recortado conserva la cola de la clave", nombre.endswith(CLAVE[-8:]))

    # Si ni el prefijo entra, avisa y no entrega un nombre que el reporte no
    # reconoceria: el problema es la carpeta, no el nombre.
    afirmar(
        "con una carpeta imposible deja el nombre intacto y avisa",
        recortar(250 - len(tipo_prefijo) - 1 - 7 - 4 - 3 - 1) == NOMBRE,
    )

    if os.name == "nt":
        prefijada = ruta_larga("C:/" + "x" * 240) / NOMBRE
        afirmar(
            "con el prefijo de ruta larga no recorta nada",
            _acortar_si_no_entra(prefijada, ".xml").name == NOMBRE,
        )
    print()


if __name__ == "__main__":
    print(f"Sistema: {os.name}\n")
    probar_helpers()
    probar_flujo()
    probar_recorte()
    if _fallos:
        print(f"FALLARON {len(_fallos)}:")
        for f in _fallos:
            print(f"  - {f}")
        sys.exit(1)
    print("Todo en orden.")
