from __future__ import annotations

import os
from typing import Optional

import requests


class LicensingClient:
    def __init__(self, base_url: Optional[str] = None, timeout: int = 10):
        self.base_url = (base_url or os.getenv("LICENSE_API_URL") or "http://localhost:8000").rstrip("/")
        self.timeout = timeout

    def _post(self, path: str, token: Optional[str] = None, json: Optional[dict] = None):
        headers = {"Content-Type": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        response = requests.post(
            f"{self.base_url}{path}",
            json=json,
            headers=headers,
            timeout=self.timeout,
        )
        if response.status_code >= 400:
            detail = response.json().get("detail") if response.headers.get("content-type", "").startswith("application/json") else response.text
            raise ValueError(detail or f"Error {response.status_code}")
        return response.json()

    def login(self, email: str, password: str) -> str:
        data = self._post("/auth/login", json={"email": email, "password": password})
        return data["access_token"]

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

    def get_profile(self, token: str) -> dict:
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{self.base_url}/me", headers=headers, timeout=self.timeout)
        if response.status_code >= 400:
            detail = response.json().get("detail") if response.headers.get("content-type", "").startswith("application/json") else response.text
            raise ValueError(detail or f"Error {response.status_code}")
        return response.json()

    def request_password_reset(self, email: str) -> dict:
        return self._post(
            "/auth/password-reset/request",
            json={"email": email},
        )

    def confirm_password_reset(self, token: str, new_password: str) -> dict:
        return self._post(
            "/auth/password-reset/confirm",
            json={"token": token, "new_password": new_password},
        )

    def preview_password_reset(self, token: str) -> dict:
        return self._post(
            "/auth/password-reset/preview",
            json={"token": token},
        )
