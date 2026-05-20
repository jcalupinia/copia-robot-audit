"""Utilidades de archivos, paths y parsing de texto.

Funciones puras (sin lógica de negocio del SRI) que manipulan el sistema
de archivos, normalizan nombres y parsean archivos TXT/CSV de claves.

Originalmente vivían en `robot/downloader.py`; se extrajeron como parte
del refactor para reducir el tamaño de ese módulo monolítico. El downloader
las re-importa para mantener la API previa intacta.
"""
from __future__ import annotations

import csv
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
