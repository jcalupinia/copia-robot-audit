"""Utilidades de archivos, paths y parsing de texto.

Funciones puras (sin lógica de negocio del SRI) que manipulan el sistema
de archivos, normalizan nombres y parsean archivos TXT/CSV de claves.

Originalmente vivían en `robot/downloader.py`; se extrajeron como parte
del refactor para reducir el tamaño de ese módulo monolítico. El downloader
las re-importa para mantener la API previa intacta.
"""
from __future__ import annotations

import csv
import os
import re
import unicodedata
from pathlib import Path

from robot._logging import get_logger


logger = get_logger(__name__)


# --------------------------------------------------------------------------- #
# Reportes Excel intermedios — colección y limpieza
# --------------------------------------------------------------------------- #
def _collect_existing_reports(
    base_dir: Path, prefix: str, tipo_slug: str, suffixes
) -> list[str]:
    """Lista los archivos .xlsx de reporte que coinciden con un patrón.

    Busca dentro de `base_dir` archivos con nombre
    `{prefix}_{tipo_slug}_{suffix}*.xlsx`, tolerando el sufijo numérico
    `_N` que se añade cuando un reporte ya existe.
    """
    if not base_dir.exists():
        return []
    encontrados: dict[str, Path] = {}
    for suffix in suffixes or []:
        suffix_str = str(suffix or "").strip()
        if not suffix_str:
            continue
        patron = f"{prefix}_{tipo_slug}_{suffix_str}*.xlsx"
        for ruta in sorted(base_dir.glob(patron)):
            if not ruta.is_file():
                continue
            stem = ruta.stem
            esperado = f"{prefix}_{tipo_slug}_{suffix_str}"
            if stem != esperado and not re.fullmatch(rf"{re.escape(esperado)}_\d+", stem):
                continue
            encontrados[str(ruta.resolve())] = ruta
    return [str(ruta) for ruta in sorted(encontrados.values())]


def _delete_report_files(reportes: list[str]) -> None:
    """Elimina archivos de reporte intermedios. Tolera errores por archivo."""
    for ruta in reportes or []:
        try:
            Path(ruta).unlink(missing_ok=True)
        except Exception as err:
            logger.warning(f"No se pudo eliminar reporte intermedio '{ruta}': {err}")


# --------------------------------------------------------------------------- #
# Helpers de fecha / nombres
# --------------------------------------------------------------------------- #
def _mes_a_texto(mes: int) -> str:
    """Convierte 1-12 al nombre del mes en español (1 → 'Enero')."""
    return [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ][mes - 1]


# --------------------------------------------------------------------------- #
# Rutas largas de Windows
# --------------------------------------------------------------------------- #
# Windows corta las rutas en 260 caracteres salvo que el equipo tenga la clave
# LongPathsEnabled en 1, que por defecto NO la tiene. El robot llega facil a ese
# limite sin que la ruta sea rara: el usuario elige una carpeta descriptiva de
# ~130 caracteres, el robot le agrega 75 de estructura
# (RUC/Emitidos/Autorizados/Tipo/anio/mes/PDF) y el nombre del archivo lleva la
# clave de acceso de 49 digitos. El sintoma era el peor posible: el proceso
# "termina bien" y no baja un solo archivo.
#
# El prefijo `\\?\` le dice a la API de Windows que no aplique ese limite (pasa
# a 32767 caracteres). Sirve sin permisos de administrador y sin tocar el
# registro, que es lo unico viable cuando la app corre en la maquina de un
# cliente. Verificado a 393 caracteres con mkdir, write, exists, rglob,
# openpyxl.save y el save_as de Playwright.
#
# A cambio exige una ruta absoluta, con barras invertidas y sin tramos "." ni
# "..": el prefijo le saca a Windows el trabajo de normalizarla, asi que hay que
# entregarsela ya normalizada.
PREFIJO_RUTA_LARGA = "\\\\?\\"
_PREFIJO_RUTA_LARGA_UNC = PREFIJO_RUTA_LARGA + "UNC" + "\\"


def ruta_larga(ruta) -> Path:
    """Devuelve la ruta en la forma que ignora el limite de 260 de Windows.

    Es idempotente, asi que se puede aplicar sin averiguar si ya se aplico.
    Fuera de Windows no hay limite que levantar y la ruta se devuelve igual.
    """
    texto = str(ruta or "")
    if not texto or os.name != "nt" or texto.startswith(PREFIJO_RUTA_LARGA):
        return Path(texto)
    # `abspath` resuelve lo relativo y los tramos "." y "..", que con el
    # prefijo Windows ya no interpreta, y pasa todo a barras invertidas.
    absoluta = os.path.abspath(texto)
    if absoluta.startswith("\\\\"):
        # Recurso de red: \\servidor\carpeta -> \\?\UNC\servidor\carpeta
        return Path(_PREFIJO_RUTA_LARGA_UNC + absoluta[2:])
    return Path(PREFIJO_RUTA_LARGA + absoluta)


def mostrar_ruta(ruta) -> str:
    """La ruta como la reconoce el usuario, sin el prefijo tecnico.

    `ruta_larga` deja rutas que empiezan con `\\?\`, que es correcto para el
    sistema de archivos pero desconcertante en un mensaje o en el log. Todo lo
    que el usuario lee pasa por aca.
    """
    texto = str(ruta or "")
    if texto.startswith(_PREFIJO_RUTA_LARGA_UNC):
        return "\\\\" + texto[len(_PREFIJO_RUTA_LARGA_UNC):]
    if texto.startswith(PREFIJO_RUTA_LARGA):
        return texto[len(PREFIJO_RUTA_LARGA):]
    return texto


def normalizar_carpeta_destino(ruta) -> Path:
    """Quita espacios y puntos del final de cada tramo de una ruta.

    Windows no admite que un nombre de carpeta termine en espacio o punto: al
    crearla los recorta en silencio, pero despues la ruta CON el espacio no
    resuelve y escribir ahi falla con FileNotFoundError. El usuario escribe
    "F:\Cliente S.A. " en el campo manual sin darse cuenta, la carpeta se crea
    como "Cliente S.A", y la descarga no guarda un solo archivo.

    Se respeta el ancla ("F:\\", "/") y se dejan intactos los tramos que
    quedarian vacios, como "." y "..".

    Al final se aplica `ruta_larga`, y con eso la longitud de la carpeta que
    elige el usuario deja de importar: como todo el resto del robot construye
    sus rutas colgando de la que devuelve esta funcion (`destino / RUC /
    "Emitidos" / ...`), el prefijo se hereda solo y no hay que repetirlo en cada
    mkdir ni en cada escritura.
    """
    original = Path(str(ruta or "").strip()).expanduser()
    partes = list(original.parts)
    if not partes:
        return original
    ancla = original.anchor
    resto = partes[1:] if ancla else partes
    limpias = []
    for tramo in resto:
        recortado = tramo.rstrip(" .")
        limpias.append(recortado or tramo)
    destino = Path(ancla, *limpias) if ancla else Path(*limpias)
    if destino != original:
        logger.info(
            f"Carpeta destino normalizada: {original} -> {destino} "
            "(Windows no admite espacios ni puntos al final)."
        )
    return ruta_larga(destino)


def _sanear_nombre_archivo(texto: str, sufijo: str = "") -> str:
    """Devuelve `texto` apto para usar como nombre de archivo.

    Normaliza acentos a ASCII, reemplaza caracteres no alfanuméricos por
    guion bajo y opcionalmente añade un sufijo. Si el resultado queda
    vacío, devuelve 'documento'.
    """
    base = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_")
    if not base:
        base = "documento"
    if sufijo:
        base = f"{base}_{sufijo}"
    return base


# --------------------------------------------------------------------------- #
# Parsing de archivos TXT/CSV de claves del SRI
# --------------------------------------------------------------------------- #
def _es_clave(valor: str) -> bool:
    """True si `valor` es una clave de acceso del SRI (49 dígitos exactos)."""
    return bool(re.fullmatch(r"\d{49}", (valor or "").strip()))


def _detectar_delimitador(sample: str) -> str:
    """Adivina el delimitador de un CSV a partir de una muestra de texto.

    Cuenta ocurrencias de `;`, `,` y `\\t` y devuelve el más frecuente.
    Si ninguno aparece, asume `;` (delimitador habitual del SRI).
    """
    counts = {";": sample.count(";"), ",": sample.count(","), "\t": sample.count("\t")}
    return max(counts, key=counts.get) if any(counts.values()) else ";"


def _extraer_claves_desde_txt(txt_path: Path):
    """Parsea un archivo TXT/CSV del SRI y devuelve lista de dicts con
    `clave`, `tipo` y `fecha` por cada fila que contenga una clave válida.
    """
    claves = []
    sample = txt_path.read_text(encoding="utf-8", errors="ignore")[:4096]
    sep = _detectar_delimitador(sample)
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f, delimiter=sep)
        for row in reader:
            if not row:
                continue
            clave = next((c.strip() for c in row if _es_clave(c)), None)
            if not clave:
                continue
            tipo = next(
                (
                    c.strip()
                    for c in row
                    if c.lower().startswith(("factura", "comprobante", "nota", "liquidacion"))
                ),
                "",
            )
            fecha = next(
                (c.strip() for c in row if re.fullmatch(r"\d{2}/\d{2}/\d{4}", c.strip())),
                "",
            )
            claves.append({"clave": clave, "tipo": tipo, "fecha": fecha})
    return claves
