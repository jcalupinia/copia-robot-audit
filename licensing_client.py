"""Cliente de la API de licencias.

La distincion que sostiene todo este modulo: **una cosa es que el servidor
conteste que no, y otra muy distinta es no poder llegar al servidor**. Antes las
dos salian como `ValueError` y quien llamaba no podia separarlas, asi que un
corte de red de tres segundos -- o Render despertando de su suspension -- le
cerraba la sesion al usuario y le pedia de nuevo el codigo de licencia.

Por eso las excepciones estan separadas:

- `RespuestaDelServidor`: el servidor contesto y rechazo. Es autoritativo y
  quien llama puede actuar en consecuencia (cerrar sesion, pedir activacion).
- `ServidorInalcanzable`: no hubo respuesta utilizable -- timeout, DNS, 5xx,
  429. NO dice nada sobre la licencia ni sobre la sesion, y nunca debe usarse
  para echar a nadie.
"""
from __future__ import annotations

import os
from typing import Optional

import requests


class LicenciaError(RuntimeError):
    """Base de los errores del cliente de licencias."""


class RespuestaDelServidor(LicenciaError):
    """El servidor contesto y rechazo la peticion. Es una respuesta valida."""

    def __init__(self, mensaje: str, status: int):
        super().__init__(mensaje)
        self.status = status

    @property
    def es_sesion_invalida(self) -> bool:
        """401: el token no sirve. Hay que renovarlo o volver a entrar."""
        return self.status == 401


class ServidorInalcanzable(LicenciaError):
    """No se pudo obtener una respuesta. No concluye nada sobre la cuenta."""


# Render suspende los servicios gratuitos por inactividad y despertarlos puede
# tomar mas de medio minuto. Con el timeout de 10 s que habia, la primera
# consulta del dia fallaba casi siempre y el usuario terminaba afuera.
TIMEOUT_POR_DEFECTO = int(os.getenv("LICENSE_API_TIMEOUT", "40"))


class LicensingClient:
    def __init__(self, base_url: Optional[str] = None, timeout: Optional[int] = None):
        self.base_url = (
            base_url or os.getenv("LICENSE_API_URL") or "http://localhost:8000"
        ).rstrip("/")
        self.timeout = timeout or TIMEOUT_POR_DEFECTO

    # ------------------------------------------------------------------ #
    # Transporte
    # ------------------------------------------------------------------ #
    def _detalle(self, response) -> str:
        try:
            if response.headers.get("content-type", "").startswith("application/json"):
                return str(response.json().get("detail") or "")
        except Exception:
            pass
        return (response.text or "")[:300]

    def _revisar(self, response) -> dict:
        if response.status_code < 400:
            return response.json() if response.content else {}
        detalle = self._detalle(response) or f"Error {response.status_code}"
        # 5xx y 429 son problemas del servidor, no un veredicto sobre el
        # usuario: tratarlos como rechazo es lo que dejaba gente afuera.
        if response.status_code >= 500 or response.status_code == 429:
            raise ServidorInalcanzable(detalle)
        raise RespuestaDelServidor(detalle, response.status_code)

    def _pedir(self, metodo: str, path: str, token: Optional[str] = None, json: Optional[dict] = None) -> dict:
        headers = {"Content-Type": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        try:
            response = requests.request(
                metodo,
                f"{self.base_url}{path}",
                json=json,
                headers=headers,
                timeout=self.timeout,
            )
        except requests.RequestException as err:
            # Timeout, DNS, conexion rechazada, sin internet.
            raise ServidorInalcanzable(str(err)) from err
        return self._revisar(response)

    def _post(self, path: str, token: Optional[str] = None, json: Optional[dict] = None) -> dict:
        return self._pedir("POST", path, token=token, json=json)

    # ------------------------------------------------------------------ #
    # Sesion
    # ------------------------------------------------------------------ #
    def login(self, email: str, password: str) -> dict:
        """Devuelve {'access_token', 'refresh_token'}.

        El refresh token puede venir vacio si la API todavia no se desplego con
        el endpoint nuevo; quien llama debe tolerarlo.
        """
        data = self._post("/auth/login", json={"email": email, "password": password})
        return {
            "access_token": data["access_token"],
            "refresh_token": data.get("refresh_token") or "",
        }

    def refresh(self, refresh_token: str) -> dict:
        """Canjea el refresh token por un access token nuevo."""
        data = self._post("/auth/refresh", json={"refresh_token": refresh_token})
        return {
            "access_token": data["access_token"],
            "refresh_token": data.get("refresh_token") or refresh_token,
        }

    def get_profile(self, token: str) -> dict:
        return self._pedir("GET", "/me", token=token)

    # ------------------------------------------------------------------ #
    # Licencia
    # ------------------------------------------------------------------ #
    def activate_license(self, token: str, code: str, fingerprint: str) -> dict:
        return self._post(
            "/license/activate",
            token=token,
            json={"code": code, "fingerprint": fingerprint},
        )

    def validate_license(self, token: str, fingerprint: str) -> dict:
        return self._post(
            "/license/validate",
            token=token,
            json={"fingerprint": fingerprint},
        )

    # ------------------------------------------------------------------ #
    # Contrasena
    # ------------------------------------------------------------------ #
    def request_password_reset(self, email: str) -> dict:
        return self._post("/auth/password-reset/request", json={"email": email})

    def confirm_password_reset(self, token: str, new_password: str) -> dict:
        return self._post(
            "/auth/password-reset/confirm",
            json={"token": token, "new_password": new_password},
        )

    def preview_password_reset(self, token: str) -> dict:
        return self._post("/auth/password-reset/preview", json={"token": token})
