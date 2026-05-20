"""Configuración centralizada de logging para el paquete `robot`.

Provee un setup idempotente (puede llamarse varias veces sin duplicar
handlers) y un helper `get_logger(name)` para uso desde los módulos.

Por defecto:
- Nivel controlado por la env var `ROBOT_LOG_LEVEL` (default: INFO).
- Salida a stderr con timestamp, nivel y nombre del logger.
- No toca el root logger: aísla la config al árbol `robot.*` para no
  interferir con Streamlit ni con FastAPI cuando se importa desde ahí.

Uso típico desde otro módulo del paquete:

    from robot._logging import get_logger
    logger = get_logger(__name__)
    logger.info("Descarga iniciada para RUC %s", ruc)
"""
from __future__ import annotations

import logging
import os
import sys

_PACKAGE_LOGGER_NAME = "robot"
_DEFAULT_FORMAT = "%(asctime)s %(levelname)-7s %(name)s | %(message)s"
_configured = False


def _resolve_level() -> int:
    raw = os.getenv("ROBOT_LOG_LEVEL", "INFO").strip().upper()
    # Soportar tanto nombres ("INFO") como valores numéricos ("20").
    if raw.isdigit():
        return int(raw)
    return getattr(logging, raw, logging.INFO)


def configure(level: int | None = None, *, force: bool = False) -> logging.Logger:
    """Configura el logger del paquete `robot`. Idempotente."""
    global _configured
    pkg_logger = logging.getLogger(_PACKAGE_LOGGER_NAME)

    if _configured and not force:
        return pkg_logger

    pkg_logger.setLevel(level if level is not None else _resolve_level())

    # Evitar duplicar handlers si configure() corre dos veces (ej. al
    # reimportar bajo Streamlit o en tests).
    if not any(
        getattr(h, "_robot_logging", False) for h in pkg_logger.handlers
    ):
        handler = logging.StreamHandler(stream=sys.stderr)
        handler.setFormatter(logging.Formatter(_DEFAULT_FORMAT))
        handler._robot_logging = True  # type: ignore[attr-defined]
        pkg_logger.addHandler(handler)

    # No propagar al root para evitar duplicados si el caller ya configuró
    # su propio handler global.
    pkg_logger.propagate = False
    _configured = True
    return pkg_logger


def get_logger(name: str) -> logging.Logger:
    """Devuelve un logger configurado bajo el árbol `robot.*`."""
    configure()
    return logging.getLogger(name)
