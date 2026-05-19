# -*- coding: ascii -*-
import os
import tempfile
import time
from typing import Optional

try:
    from twocaptcha import TwoCaptcha  # type: ignore
except Exception:  # pragma: no cover - dependency opcional
    TwoCaptcha = None  # type: ignore


def _env_flag(valor: Optional[str]) -> Optional[bool]:
    if valor is None:
        return None
    texto = valor.strip().lower()
    if not texto:
        return None
    return texto in {"1", "true", "yes", "on", "si"}



class CaptchaSolverError(RuntimeError):
    """Fallo al resolver un captcha mediante 2Captcha."""


API_KEY = (os.getenv("TWOCAPTCHA_API_KEY") or "").strip()
USE_2CAPTCHA = _env_flag(os.getenv("USE_2CAPTCHA"))
FORCE_DISABLE_2CAPTCHA = True
DEFAULT_TIMEOUT = int(os.getenv("TWOCAPTCHA_TIMEOUT", "120"))
DEFAULT_POLLING = int(os.getenv("TWOCAPTCHA_POLLING", "5"))
MAX_ATTEMPTS = max(1, int(os.getenv("TWOCAPTCHA_ATTEMPTS", "3")))
ATTEMPT_DELAY = max(0, float(os.getenv("TWOCAPTCHA_ATTEMPT_DELAY", "3")))

_solver_instance: Optional["TwoCaptcha"] = None


def is_enabled() -> bool:
    if FORCE_DISABLE_2CAPTCHA:
        return False
    if USE_2CAPTCHA is not None:
        return USE_2CAPTCHA and bool(API_KEY)
    return bool(API_KEY)


def _ensure_solver() -> "TwoCaptcha":
    if not is_enabled():
        raise CaptchaSolverError("2Captcha no configurado; faltan TWOCAPTCHA_API_KEY o USE_2CAPTCHA desactivado.")
    if TwoCaptcha is None:
        raise CaptchaSolverError("La libreria '2captcha-python' no esta disponible. Instalala e intentalo de nuevo.")

    global _solver_instance
    if _solver_instance is None:
        _solver_instance = TwoCaptcha(
            API_KEY,
            defaultTimeout=DEFAULT_TIMEOUT,
            pollingInterval=DEFAULT_POLLING,
        )
    return _solver_instance


def solve_image(image_bytes: bytes) -> str:
    solver = _ensure_solver()
    ultimo_error: Optional[Exception] = None

    for intento in range(1, MAX_ATTEMPTS + 1):
        tmp_path: Optional[str] = None
        try:
            tmp = tempfile.NamedTemporaryFile(prefix="captcha_", suffix=".png", delete=False)
            tmp.write(image_bytes)
            tmp.flush()
            tmp_path = tmp.name
            tmp.close()

            resultado = solver.normal(tmp_path)
            if isinstance(resultado, dict):
                codigo = (resultado.get("code") or resultado.get("text") or "").strip()
            else:
                codigo = str(resultado or "").strip()
            if codigo:
                return codigo
            ultimo_error = CaptchaSolverError("2Captcha no retorno un codigo util.")
        except Exception as exc:  # pragma: no cover - depende del servicio externo
            ultimo_error = exc
        finally:
            if tmp_path:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
        if intento < MAX_ATTEMPTS and ATTEMPT_DELAY:
            time.sleep(ATTEMPT_DELAY)

    raise CaptchaSolverError(str(ultimo_error) if ultimo_error else "Fallo desconocido resolviendo captcha.")


def solve_recaptcha(
    site_key: str,
    page_url: str,
    data_s: Optional[str] = None,
    action: Optional[str] = None,
    enterprise: bool = False,
) -> str:
    solver = _ensure_solver()
    ultimo_error: Optional[Exception] = None

    for intento in range(1, MAX_ATTEMPTS + 1):
        try:
            params = {
                "sitekey": site_key,
                "url": page_url,
            }
            if data_s:
                params["data_s"] = data_s
            if action:
                params["action"] = action
            if enterprise:
                params["enterprise"] = 1

            resultado = solver.recaptcha(**params)
            if isinstance(resultado, dict):
                codigo = (resultado.get("code") or resultado.get("text") or "").strip()
            else:
                codigo = str(resultado or "").strip()
            if codigo:
                return codigo
            ultimo_error = CaptchaSolverError("2Captcha no retorno token para reCAPTCHA.")
        except Exception as exc:  # pragma: no cover - depende del servicio externo
            ultimo_error = exc
        if intento < MAX_ATTEMPTS and ATTEMPT_DELAY:
            time.sleep(ATTEMPT_DELAY)

    raise CaptchaSolverError(str(ultimo_error) if ultimo_error else "Fallo desconocido resolviendo reCAPTCHA.")


__all__ = [
    "CaptchaSolverError",
    "is_enabled",
    "solve_image",
    "solve_recaptcha",
    "MAX_ATTEMPTS",
]
