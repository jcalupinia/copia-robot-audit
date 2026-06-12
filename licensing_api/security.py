from __future__ import annotations

import os
from datetime import datetime, timedelta
from typing import Optional

from jose import jwt, JWTError
from passlib.context import CryptContext


pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv("LICENSE_TOKEN_EXPIRE_MINUTES", "60"))


def _load_jwt_secret() -> str:
    """Carga y valida la clave secreta para firmar JWTs.

    Falla rápido en el arranque si la variable de entorno LICENSE_API_SECRET
    no está configurada, conserva el placeholder original o es demasiado corta.
    Esto evita arrancar la API con un secret predecible (cualquiera podría
    forjar tokens y suplantar usuarios).
    """
    secret = os.getenv("LICENSE_API_SECRET")
    hint = (
        'Generá uno con: python -c "import secrets; '
        'print(secrets.token_urlsafe(48))" y configurá la variable de '
        "entorno LICENSE_API_SECRET (en Render: Settings → Environment) "
        "antes de arrancar la API."
    )
    if not secret or secret == "CHANGE_THIS_SECRET":
        raise RuntimeError(
            "LICENSE_API_SECRET no está configurado (o sigue usando el "
            f"placeholder por defecto). {hint}"
        )
    if len(secret) < 32:
        raise RuntimeError(
            "LICENSE_API_SECRET es demasiado corto: se requieren al menos "
            f"32 caracteres para HS256 (longitud actual: {len(secret)}). {hint}"
        )
    return secret


JWT_SECRET_KEY = _load_jwt_secret()
JWT_ALGORITHM = os.getenv("LICENSE_API_ALGORITHM", "HS256")


def verify_password(plain_password: str, password_hash: str) -> bool:
    return pwd_context.verify(plain_password, password_hash)


def get_password_hash(password: str) -> str:
    return pwd_context.hash(password)


def create_access_token(subject: str, expires_delta: Optional[timedelta] = None) -> str:
    expire = datetime.utcnow() + (
        expires_delta or timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    )
    payload = {"sub": subject, "exp": expire}
    return jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)


def decode_access_token(token: str) -> Optional[str]:
    try:
        payload = jwt.decode(token, JWT_SECRET_KEY, algorithms=[JWT_ALGORITHM])
        return payload.get("sub")
    except JWTError:
        return None
