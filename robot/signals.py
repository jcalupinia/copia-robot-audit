"""Señales y notificaciones compartidas del paquete `robot`.

Encapsula dos piezas de estado global que el bot intercambia con la UI:

- `CANCEL_EVENT`: bandera de cancelación cooperativa. Cualquier código de
  larga duración llama a `_check_cancel()` periódicamente para abortar
  limpiamente si el usuario pidió cancelar desde la UI.
- `USER_NOTIFICATION_CALLBACK`: callable opcional al que se le pasan
  mensajes para mostrar al usuario (Streamlit lo configura con
  `set_user_notifier`). Se accede SIEMPRE a través de `notify()`; nunca
  importes la variable directamente porque capturarías el valor inicial
  (None) en tiempo de import y nunca verías actualizaciones posteriores.
"""
from __future__ import annotations

import threading
from typing import Callable, Optional

from robot._logging import get_logger


logger = get_logger(__name__)


# Estado mutable: se accede mediante las funciones públicas de este módulo.
USER_NOTIFICATION_CALLBACK: Optional[Callable[[str], None]] = None
CANCEL_EVENT = threading.Event()


# --------------------------------------------------------------------------- #
# Notificación al usuario
# --------------------------------------------------------------------------- #
def set_user_notifier(callback: Optional[Callable[[str], None]]) -> None:
    """Registra (o limpia, pasando None) el callback de notificación al UI."""
    global USER_NOTIFICATION_CALLBACK
    USER_NOTIFICATION_CALLBACK = callback


def notify(mensaje: str) -> None:
    """Envía `mensaje` al callback configurado, si lo hay.

    Esta función centraliza el manejo de errores del callback: si la UI
    falla por cualquier motivo (Streamlit reiniciando, callback rota, etc.)
    se loguea el problema pero no se rompe el flujo del bot.
    """
    if not USER_NOTIFICATION_CALLBACK:
        return
    try:
        USER_NOTIFICATION_CALLBACK(mensaje)
    except Exception as err:
        logger.warning(f"No se pudo enviar notificacion al UI: {err}")


# --------------------------------------------------------------------------- #
# Cancelación cooperativa
# --------------------------------------------------------------------------- #
def request_cancel() -> None:
    """Marca que se debe abortar el proceso en curso (llamado desde la UI)."""
    CANCEL_EVENT.set()


def clear_cancel() -> None:
    """Limpia la señal de cancelación (al iniciar un nuevo proceso)."""
    CANCEL_EVENT.clear()


def cancel_requested() -> bool:
    """Indica si la UI pidió cancelar."""
    return CANCEL_EVENT.is_set()


def _check_cancel(paso: str = "") -> None:
    """Lanza RuntimeError si la cancelación fue solicitada.

    `paso` es informativo y permite identificar dónde se cortó el flujo.
    """
    if CANCEL_EVENT.is_set():
        raise RuntimeError("Proceso cancelado por el usuario.")
