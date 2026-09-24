from __future__ import annotations

import os
from datetime import datetime, timedelta
from typing import Optional

from jose import jwt, JWTError
from passlib.context import CryptContext


pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv("LICENSE_TOKEN_EXPIRE_MINUTES", "60"))

# El access token dura poco a proposito: si se filtra, se vence solo. Lo que
# mantiene viva la sesion es el refresh token, que el cliente guarda y canjea
# por uno nuevo sin volver a pedir la contrasena. Sin esto la sesion se caia a
# los 60 minutos y al usuario le parecia que se cerraba "de la nada".
REFRESH_TOKEN_EXPIRE_DAYS = int(os.getenv("LICENSE_REFRESH_EXPIRE_DAYS", "180"))


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
    payload = {"sub": subject, "exp": expire, "typ": "access"}
    return jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)


def create_refresh_token(subject: str, expires_delta: Optional[timedelta] = None) -> str:
    """Token de larga duracion, util SOLO para pedir un access token nuevo."""
    expire = datetime.utcnow() + (
        expires_delta or timedelta(days=REFRESH_TOKEN_EXPIRE_DAYS)
    )
    payload = {"sub": subject, "exp": expire, "typ": "refresh"}
    return jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)


def _decode(token: str, tipo_esperado: str) -> Optional[str]:
    """Decodifica exigiendo el tipo, para que uno no sirva por el otro.

    Sin el `typ` el refresh token -- que vive meses -- se podria mandar como
    Authorization y daria acceso todo ese tiempo, que es justo lo que su
    duracion larga hace peligroso. Los tokens viejos no traen `typ`: se
    aceptan como access para no cerrarle la sesion a quien ya estaba dentro
    cuando se despliegue esto.
    """
    try:
        payload = jwt.decode(token, JWT_SECRET_KEY, algorithms=[JWT_ALGORITHM])
    except JWTError:
        return None
    tipo = payload.get("typ", "access")
    if tipo != tipo_esperado:
        return None
    return payload.get("sub")


def decode_access_token(token: str) -> Optional[str]:
    return _decode(token, "access")


def decode_refresh_token(token: str) -> Optional[str]:
    return _decode(token, "refresh")
