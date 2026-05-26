"""Detección y manejo de captchas del portal del SRI.

Política actual (2026-05): **NO usamos servicios externos** (2Captcha u otros)
para resolver captchas. Todo se resuelve manualmente por el usuario.

Cubre dos tipos:

- **Captcha de imagen** (input numérico): se notifica al usuario y se espera
  a que escriba el código en el input visible.
- **reCAPTCHA** (Google v2/Enterprise/invisible): se notifica al usuario y se
  espera a que Google emita un token válido o el desafío visual desaparezca.

En ambos casos, si la UI tiene callback configurado (`robot.signals.notify`),
se le manda un mensaje pidiendo intervención.

Originalmente vivía dentro de `robot/downloader.py`; se extrajo en la
Sub-fase 2b del refactor. La integración con 2Captcha se eliminó por
completo en 2026-05 (ya no se usa el servicio).
"""
from __future__ import annotations

import time

from robot._logging import get_logger
from robot.signals import _notificar_usuario_captcha


logger = get_logger(__name__)


# --------------------------------------------------------------------------- #
# Reintentos
# --------------------------------------------------------------------------- #
# Cantidad de veces que el flujo de login puede reintentar tras detectar un
# captcha residual (cuando el usuario ingresó el código pero el portal lo
# rechaza). Sirve como tope global; no implica resolución automática.
CAPTCHA_MAX_ATTEMPTS = 3


# --------------------------------------------------------------------------- #
# Selectores comunes
# --------------------------------------------------------------------------- #
CAPTCHA_INPUT_SELECTORS = [
    "input[name*='captcha' i]",
    "input[id*='captcha' i]",
    "input[name='captcha']",
    "input[id='captcha']",
    "input#captchaIngresar",
    "input#captchaTxt",
]
CAPTCHA_INPUT_QUERY = ",".join(CAPTCHA_INPUT_SELECTORS)


# --------------------------------------------------------------------------- #
# Captcha de imagen
# --------------------------------------------------------------------------- #
def _espera_captcha(page, timeout: int = 1000):
    """Espera a que la imagen de captcha desaparezca del DOM, hasta `timeout` ms."""
    try:
        loc = page.locator("img[alt='captcha']")
        if loc.is_visible(timeout=1000):
            page.wait_for_selector("img[alt='captcha']", state="detached", timeout=timeout)
    except Exception:
        pass


def _captcha_visible(page, timeout: int = 0) -> bool:
    """True si la imagen de captcha está visible. `timeout` opcional en ms."""
    try:
        loc = page.locator("img[alt='captcha']")
        if timeout:
            return loc.is_visible(timeout=timeout)
        return loc.is_visible()
    except Exception:
        return False


def _localizar_input_captcha(page):
    """Devuelve el primer Locator de input de captcha que exista en la página."""
    for selector in CAPTCHA_INPUT_SELECTORS:
        try:
            locator = page.locator(selector)
            if locator.count():
                return locator.first
        except Exception:
            continue
    return None


def _esperar_captcha_manual_input(page, timeout: int = 300000) -> bool:
    """Espera hasta que el usuario ingrese un valor de captcha de forma manual.

    Se considera resuelto cuando cualquiera de los inputs registrados tiene
    al menos 4 caracteres.
    """
    try:
        page.wait_for_function(
            """(selectorCadena) => {
                const inputs = document.querySelectorAll(selectorCadena);
                for (const input of inputs) {
                    const valor = (input.value || "").trim();
                    if (valor.length >= 4) {
                        return true;
                    }
                }
                return false;
            }""",
            arg=CAPTCHA_INPUT_QUERY,
            timeout=timeout,
        )
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------- #
# reCAPTCHA (Google v2 / Enterprise / invisible)
# --------------------------------------------------------------------------- #
def _recaptcha_presente(page) -> bool:
    """True si hay un widget de reCAPTCHA en la página."""
    try:
        if page.locator("iframe[src*='recaptcha']").count():
            return True
    except Exception:
        pass
    try:
        if page.locator("[data-sitekey]").count():
            return True
    except Exception:
        pass
    return False


def _recaptcha_challenge_activo(page) -> bool:
    """True si el desafío visual de reCAPTCHA está actualmente desplegado."""
    try:
        frame = page.locator("iframe[src*='recaptcha/api2/bframe']")
        if frame.count():
            try:
                return frame.first.is_visible()
            except Exception:
                pass
    except Exception:
        pass
    return False


def _esperar_recaptcha_resuelto(page, timeout: int = 300000) -> bool:
    """Espera a que el desafío de reCAPTCHA desaparezca u obtenga respuesta."""
    fin = time.time() + timeout / 1000
    while time.time() < fin:
        challenge_activo = False
        try:
            challenge_activo = _recaptcha_challenge_activo(page)
        except Exception:
            challenge_activo = False
        if not challenge_activo:
            return True
        try:
            page.wait_for_timeout(500)
        except Exception:
            pass
    return False


# --------------------------------------------------------------------------- #
# Resolución orquestada — siempre manual
# --------------------------------------------------------------------------- #
def _resolver_captcha(page, contexto: str) -> bool:
    """Notifica al usuario y espera resolución manual del captcha.

    Estrategia (sin resolvers externos):
      1. Si hay reCAPTCHA, notifica al usuario y espera a que el desafío
         desaparezca (Google emite token o el usuario lo resuelve a mano).
      2. Si hay captcha de imagen, notifica al usuario y espera a que la
         imagen sea reemplazada (cuando el portal recibe la respuesta).
      3. Si no hay captcha visible, devuelve False.

    Devuelve True si se considera resuelto, False si no había captcha
    detectable.
    """
    try:
        recaptcha_detectado = _recaptcha_presente(page)
    except Exception:
        recaptcha_detectado = False

    if recaptcha_detectado:
        _notificar_usuario_captcha("reCAPTCHA", contexto)
        logger.info(f"reCAPTCHA detectado ({contexto}); esperando resolucion manual.")
        _esperar_recaptcha_resuelto(page, timeout=300000)
        return True

    try:
        if not _captcha_visible(page, timeout=1000):
            return False
    except Exception:
        return False

    _notificar_usuario_captcha("captcha de imagen", contexto)
    logger.info(f"Captcha de imagen detectado ({contexto}); esperando resolucion manual.")
    _esperar_captcha_manual_input(page, timeout=300000)
    _espera_captcha(page)
    return True


__all__ = [
    "CAPTCHA_MAX_ATTEMPTS",
    "CAPTCHA_INPUT_SELECTORS",
    "CAPTCHA_INPUT_QUERY",
    "_espera_captcha",
    "_captcha_visible",
    "_localizar_input_captcha",
    "_esperar_captcha_manual_input",
    "_recaptcha_presente",
    "_recaptcha_challenge_activo",
    "_esperar_recaptcha_resuelto",
    "_resolver_captcha",
]
