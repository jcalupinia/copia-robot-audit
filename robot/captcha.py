"""Detección y resolución de captchas del portal del SRI.

Cubre dos tipos:

- **Captcha de imagen** (input numérico): se resuelve con `robot.captcha_solver`
  (2Captcha) si está activado, o se le pide al usuario que lo ingrese a mano.
- **reCAPTCHA** (Google v2): siempre lo resuelve el usuario; el bot solo
  espera a que el desafío desaparezca.

En ambos casos, si la UI tiene callback configurado (`robot.signals.notify`),
se le manda un mensaje pidiendo intervención.

Originalmente vivía dentro de `robot/downloader.py`; se extrajo en la
Sub-fase 2b del refactor.
"""
from __future__ import annotations

import time

from robot._logging import get_logger
from robot.captcha_solver import (
    CaptchaSolverError,
    MAX_ATTEMPTS as CAPTCHA_MAX_ATTEMPTS,
    is_enabled as captcha_solver_enabled,
    solve_image as solve_captcha_image,
)
from robot.signals import _notificar_usuario_captcha


logger = get_logger(__name__)


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
# reCAPTCHA (Google v2)
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
# Resolución orquestada
# --------------------------------------------------------------------------- #
def _resolver_captcha(page, contexto: str) -> bool:
    """Intenta resolver el captcha de la página actual.

    Estrategia:
    1. Si hay reCAPTCHA, notifica al usuario y espera que lo resuelva.
    2. Si hay captcha de imagen y 2Captcha está deshabilitado, espera manual.
    3. Si 2Captcha está habilitado, intenta hasta `CAPTCHA_MAX_ATTEMPTS` veces.
    4. Si todos los intentos automáticos fallan, cae a espera manual.

    Devuelve True si el captcha fue resuelto (automática o manualmente),
    False si no había captcha o no se pudo resolver.
    """
    recaptcha_detectado = False
    try:
        recaptcha_detectado = _recaptcha_presente(page)
    except Exception:
        recaptcha_detectado = False

    if recaptcha_detectado:
        _notificar_usuario_captcha("reCAPTCHA", contexto)
        _esperar_recaptcha_resuelto(page, timeout=300000)
        return True

    try:
        if not _captcha_visible(page, timeout=1000):
            return False
    except Exception:
        return False

    if not captcha_solver_enabled():
        _notificar_usuario_captcha("captcha de imagen", contexto)
        _espera_captcha(page)
        return False

    for intento in range(1, CAPTCHA_MAX_ATTEMPTS + 1):
        try:
            if not _captcha_visible(page, timeout=1000):
                return False
        except Exception:
            return False

        input_captcha = _localizar_input_captcha(page)
        if input_captcha is None:
            logger.warning(
                f"Campo de texto para captcha no encontrado ({contexto}); "
                "esperando resolucion manual."
            )
            _notificar_usuario_captcha("captcha de imagen", contexto)
            _espera_captcha(page)
            return False

        try:
            imagen = page.locator("img[alt='captcha']").screenshot(type="png")
        except Exception as err:
            logger.warning(
                f"No se pudo capturar la imagen del captcha "
                f"(intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}"
            )
            break

        try:
            codigo = solve_captcha_image(imagen)
        except CaptchaSolverError as err:
            logger.warning(
                f"Fallo al resolver captcha con 2Captcha "
                f"(intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}"
            )
            continue

        try:
            input_captcha.fill("")
            input_captcha.fill(codigo)
            return True
        except Exception as err:
            logger.warning(
                f"No se pudo escribir el captcha resuelto "
                f"(intento {intento}/{CAPTCHA_MAX_ATTEMPTS}): {err}"
            )

    logger.warning(
        "Se agotaron los intentos automaticos de captcha; "
        "esperando resolucion manual."
    )
    _notificar_usuario_captcha("captcha de imagen", contexto)
    _espera_captcha(page)
    return False
