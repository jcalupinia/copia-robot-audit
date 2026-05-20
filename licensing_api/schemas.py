from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, EmailStr, constr


# --------------------------------------------------------------------------- #
# Patrones reutilizables
# --------------------------------------------------------------------------- #
# Alfabeto base64url-safe: lo que genera secrets.token_urlsafe(...). Se usa
# para license codes y password-reset tokens. Si en el futuro hace falta
# admitir códigos con guiones bajos por bloques (ej: AAAA-BBBB-CCCC), basta
# con incluir aquí el carácter extra.
_BASE64URL_PATTERN = r"^[A-Za-z0-9_\-]+$"

# Fingerprint: hex hash por defecto, pero el usuario puede ingresarlo
# manualmente desde la UI ([aplicacion.py]). Aceptamos cualquier ASCII
# imprimible para no romper input legítimo; el límite real lo pone max_length
# y el alfabeto bloquea control chars y bytes no-ASCII.
_PRINTABLE_ASCII_PATTERN = r"^[\x20-\x7e]+$"


# --------------------------------------------------------------------------- #
# Auth
# --------------------------------------------------------------------------- #
class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class LoginRequest(BaseModel):
    # EmailStr ya valida formato (RFC 5322). Limitamos para evitar emails
    # absurdamente largos que igual no caben en la columna de DB (255 chars).
    email: EmailStr
    # min_length=6 se conserva para no romper a usuarios con passwords cortas
    # ya registradas. max_length=128 evita DoS por hashing de inputs gigantes.
    password: constr(min_length=6, max_length=128)


# --------------------------------------------------------------------------- #
# Licencias
# --------------------------------------------------------------------------- #
class LicenseActivationRequest(BaseModel):
    # `code` se genera con secrets.token_urlsafe(24) → 32 chars base64url.
    # La columna DB acepta hasta 64. Restringimos alfabeto al base64url-safe.
    code: constr(min_length=8, max_length=64, pattern=_BASE64URL_PATTERN)
    fingerprint: constr(
        min_length=6, max_length=255, pattern=_PRINTABLE_ASCII_PATTERN
    )


class LicenseValidationRequest(BaseModel):
    fingerprint: constr(
        min_length=6, max_length=255, pattern=_PRINTABLE_ASCII_PATTERN
    )


class LicenseDeviceInfo(BaseModel):
    fingerprint: str
    created_at: Optional[datetime]

    class Config:
        orm_mode = True


class LicenseInfo(BaseModel):
    code: str
    is_active: bool
    activated_at: Optional[datetime]
    expires_at: Optional[datetime]
    device_fingerprint: Optional[str]
    devices: list[LicenseDeviceInfo] = []

    class Config:
        orm_mode = True


class UserInfo(BaseModel):
    email: EmailStr
    is_active: bool
    licenses: list[LicenseInfo] = []

    class Config:
        orm_mode = True


# --------------------------------------------------------------------------- #
# Password reset
# --------------------------------------------------------------------------- #
class PasswordResetRequest(BaseModel):
    email: EmailStr


class PasswordResetConfirm(BaseModel):
    # Tokens generados con secrets.token_urlsafe(32) → 43 chars base64url.
    # Aceptamos 16–128 para tolerar variaciones futuras.
    token: constr(min_length=16, max_length=128, pattern=_BASE64URL_PATTERN)
    # Para passwords NUEVAS exigimos mínimo 8 (mejor higiene). Esto no afecta
    # logins de cuentas existentes con passwords cortas: solo aplica al setear
    # una contraseña nueva vía reset.
    new_password: constr(min_length=8, max_length=128)


class PasswordResetPreviewRequest(BaseModel):
    token: constr(min_length=16, max_length=128, pattern=_BASE64URL_PATTERN)


class PasswordResetPreviewResponse(BaseModel):
    email: EmailStr


# --------------------------------------------------------------------------- #
# Genéricos
# --------------------------------------------------------------------------- #
class MessageResponse(BaseModel):
    detail: str
