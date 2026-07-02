from __future__ import annotations


from datetime import datetime, timedelta
from email.message import EmailMessage
import hashlib
import os
from pathlib import Path
import secrets
import smtplib
import ssl

from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

import boto3
from botocore.config import Config

from sqlalchemy.orm import Session


from licensing_api import schemas, models, crud

from licensing_api.database import Base, engine, get_db

from licensing_api.security import (

    verify_password,

    create_access_token,

    decode_access_token,
    get_password_hash,

)

from licensing_api.admin import router as admin_router


oauth2_scheme = OAuth2PasswordBearer(tokenUrl="auth/login")


app = FastAPI(

    title="Licensing API",

    version="1.0.0",

    description="API para autenticación y activación de licencias del robot.",

)


# CORS: por defecto cerrado (solo localhost para dev). En producción, configurar
# la env var ALLOWED_ORIGINS con la lista separada por comas de orígenes
# permitidos (ej: "https://sri-robot-audit-ik01.onrender.com,https://admin.miempresa.com").
# Usar "*" solo si se entiende el riesgo; no es compatible con allow_credentials=True.
_default_origins = "http://localhost:8501,http://127.0.0.1:8501"
_allowed_origins_raw = os.getenv("ALLOWED_ORIGINS", _default_origins).strip()
_allowed_origins = [o.strip() for o in _allowed_origins_raw.split(",") if o.strip()]
_allow_credentials = "*" not in _allowed_origins  # incompatibles según CORS spec

app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins or [],
    allow_credentials=_allow_credentials,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type", "X-Update-Token"],
)


Base.metadata.create_all(bind=engine)

# Migraciones idempotentes de esquema y bootstrap de admins.
# Se corre DESPUES de create_all para poder hacer ALTER TABLE ADD COLUMN sobre
# tablas ya creadas. Ver licensing_api/migrations.py.
from licensing_api.database import SessionLocal as _SessionLocal  # noqa: E402
from licensing_api.migrations import run_migrations as _run_migrations  # noqa: E402

try:
    _run_migrations(engine, _SessionLocal)
except Exception as _mig_err:
    # No abortamos el arranque si la migracion falla — preferimos que la
    # API levante y el admin arregle a mano vs quedar caido.
    import logging as _logging
    _logging.getLogger(__name__).error(
        f"Migraciones fallaron: {_mig_err}. La API arranca igual."
    )


# Panel administrativo web (HTML) bajo /admin.
# Permite gestionar usuarios y licencias desde una interfaz HTML simple,
# sin pasar por el Shell de Render. Protegido por ADMIN_EMAIL/ADMIN_PASSWORD
# (env vars) + cookie httponly samesite=lax. Ver licensing_api/admin.py.
app.include_router(admin_router)


# === Static mount para los assets del landing nuevo ===
# El landing vive en `licensing_api/landing/` (HTML + assets/{css, js, fonts, img}).
# Se sirven los assets bajo la ruta /landing-assets/ para que paths relativos
# como /landing-assets/styles.css o /landing-assets/img/robot-audit.svg funcionen
# desde el HTML. El index.html lo lee y sirve el endpoint `landing_page()` con
# substitucion de placeholders (version + URL de descarga).
_LANDING_DIR = Path(__file__).resolve().parent / "landing"
_LANDING_ASSETS_DIR = _LANDING_DIR / "assets"
if _LANDING_ASSETS_DIR.is_dir():
    app.mount(
        "/landing-assets",
        StaticFiles(directory=str(_LANDING_ASSETS_DIR)),
        name="landing-assets",
    )


def _require_update_token(request: Request) -> None:
    expected = os.getenv("UPDATE_TOKEN", "").strip()
    if not expected:
        return
    provided = request.headers.get("X-Update-Token", "").strip()
    query_token = request.query_params.get("token", "").strip()
    if provided != expected and query_token != expected:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token de actualizacion invalido.")


def _r2_client():
    account_id = os.getenv("R2_ACCOUNT_ID", "").strip()
    access_key = os.getenv("R2_ACCESS_KEY_ID", "").strip()
    secret_key = os.getenv("R2_SECRET_ACCESS_KEY", "").strip()
    endpoint = os.getenv("R2_ENDPOINT", "").strip()
    region = os.getenv("R2_REGION", "auto").strip() or "auto"
    if not endpoint and account_id:
        endpoint = f"https://{account_id}.r2.cloudflarestorage.com"
    if not (endpoint and access_key and secret_key):
        return None
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=region,
        config=Config(signature_version="s3v4"),
    )


def _r2_presigned_url() -> str | None:
    bucket = os.getenv("R2_BUCKET", "").strip()
    object_key = os.getenv("R2_OBJECT_KEY", "").strip() or os.getenv("UPDATE_OBJECT_KEY", "").strip()
    if not object_key:
        object_key = "ROBOT_AUDIT_SRI.exe"
    expires = os.getenv("R2_URL_EXPIRES", "900").strip()
    try:
        expires_in = int(expires)
    except ValueError:
        expires_in = 900
    if not bucket:
        return None
    client = _r2_client()
    if client is None:
        return None
    try:
        return client.generate_presigned_url(
            "get_object",
            Params={
                "Bucket": bucket,
                "Key": object_key,
                "ResponseContentDisposition": f'attachment; filename="{object_key}"',
                "ResponseContentType": "application/x-msdownload",
            },
            ExpiresIn=expires_in,
        )
    except Exception:
        return None


def _r2_download_source() -> tuple[object, str, str] | tuple[None, None, None]:
    bucket = os.getenv("R2_BUCKET", "").strip()
    object_key = os.getenv("R2_OBJECT_KEY", "").strip() or os.getenv("UPDATE_OBJECT_KEY", "").strip()
    if not object_key:
        object_key = "ROBOT_AUDIT_SRI.exe"
    if not bucket:
        return None, None, None
    client = _r2_client()
    if client is None:
        return None, None, None
    return client, bucket, object_key


def _iter_r2_body(stream_body, chunk_size: int = 1024 * 512):
    try:
        for chunk in stream_body.iter_chunks(chunk_size=chunk_size):
            if chunk:
                yield chunk
    finally:
        try:
            stream_body.close()
        except Exception:
            pass


def _public_base_url(request: Request) -> str:
    explicit = os.getenv("APP_BASE_URL", "").strip()
    if explicit:
        return explicit.rstrip("/")
    proto = request.headers.get("x-forwarded-proto", "").strip() or request.url.scheme
    host = request.headers.get("x-forwarded-host", "").strip() or request.headers.get("host", "").strip() or request.url.netloc
    return f"{proto}://{host}".rstrip("/")


def _password_reset_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def _build_reset_link(request: Request, token: str) -> str:
    base_url = os.getenv("RESET_LINK_BASE_URL", "").strip() or os.getenv("APP_BASE_URL", "").strip()
    if not base_url:
        base_url = "http://127.0.0.1:8501"
    base_url = base_url.rstrip("/")
    return f"{base_url}/?reset_token={token}"


def _send_reset_email(target_email: str, reset_link: str) -> None:
    sender = os.getenv("SMTP_FROM") or os.getenv("SMTP_USER") or "no-reply@example.com"
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    if not all([host, user, password]):
        raise RuntimeError("SMTP no configurado en el servidor.")

    msg = EmailMessage()
    msg["Subject"] = "Recupera tu contraseña - SRI Robot"
    msg["From"] = sender
    msg["To"] = target_email
    msg.set_content(
        f"""Hola,

Hemos recibido una solicitud para restablecer tu contraseña.

Enlace de recuperación:
{reset_link}

Si no solicitaste este cambio, ignora este mensaje.
"""
    )

    use_tls = os.getenv("SMTP_USE_TLS", "1").lower() not in {"0", "false", "no"}
    context = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=context, timeout=10) as server:
            server.login(user, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=10) as server:
            if use_tls:
                server.starttls(context=context)
            server.login(user, password)
            server.send_message(msg)


def _create_password_reset_token(db: Session, user: models.User) -> str:
    ttl_raw = os.getenv("RESET_TOKEN_TTL", "3600").strip()
    try:
        ttl_seconds = max(300, int(ttl_raw))
    except ValueError:
        ttl_seconds = 3600

    now = datetime.utcnow()
    token = secrets.token_urlsafe(32)
    token_hash = _password_reset_hash(token)
    expires_at = now + timedelta(seconds=ttl_seconds)

    db.query(models.PasswordResetToken).filter(
        models.PasswordResetToken.user_id == user.id
    ).delete(synchronize_session=False)
    db.add(
        models.PasswordResetToken(
            user_id=user.id,
            token_hash=token_hash,
            expires_at=expires_at,
            used_at=None,
        )
    )
    db.commit()
    return token


def _get_valid_password_reset(db: Session, raw_token: str) -> tuple[models.PasswordResetToken, models.User]:
    token_hash = _password_reset_hash(raw_token.strip())
    now = datetime.utcnow()
    reset_req = (
        db.query(models.PasswordResetToken)
        .filter(
            models.PasswordResetToken.token_hash == token_hash,
            models.PasswordResetToken.used_at.is_(None),
        )
        .first()
    )
    if not reset_req:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Token de recuperación inválido.")
    if reset_req.expires_at < now:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="El token de recuperación expiró.")

    user = db.query(models.User).filter(models.User.id == reset_req.user_id).first()
    if not user or not user.is_active:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Usuario no válido para recuperación.")
    return reset_req, user

def get_current_user(

    token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)

) -> models.User:

    email = decode_access_token(token)

    if not email:

        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token inválido.")

    user = crud.get_user_by_email(db, email=email)

    if not user or not user.is_active:

        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Usuario inactivo.")

    return user



@app.post("/auth/login", response_model=schemas.TokenResponse)

def login(request: schemas.LoginRequest, db: Session = Depends(get_db)):

    user = crud.get_user_by_email(db, email=request.email)

    if not user or not verify_password(request.password, user.password_hash):

        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Credenciales incorrectas.")

    if not user.is_active:

        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Usuario inactivo.")

    token = create_access_token(user.email)

    return schemas.TokenResponse(access_token=token)



@app.get("/me", response_model=schemas.UserInfo)

def read_current_user(user: models.User = Depends(get_current_user)):

    return user


@app.post("/auth/password-reset/request", response_model=schemas.MessageResponse)
def request_password_reset(
    payload: schemas.PasswordResetRequest,
    request: Request,
    db: Session = Depends(get_db),
):
    # Respuesta uniforme para no revelar si el correo existe o no.
    generic_message = "Si el correo existe, enviaremos un enlace de recuperación."
    user = crud.get_user_by_email(db, email=payload.email.strip().lower())
    if not user or not user.is_active:
        return schemas.MessageResponse(detail=generic_message)

    try:
        token = _create_password_reset_token(db, user)
        reset_link = _build_reset_link(request, token)
        _send_reset_email(user.email, reset_link)
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"No se pudo enviar el correo de recuperación: {exc}",
        )
    return schemas.MessageResponse(detail=generic_message)


@app.post("/auth/password-reset/preview", response_model=schemas.PasswordResetPreviewResponse)
def preview_password_reset(
    payload: schemas.PasswordResetPreviewRequest,
    db: Session = Depends(get_db),
):
    _, user = _get_valid_password_reset(db, payload.token)
    return schemas.PasswordResetPreviewResponse(email=user.email)


@app.post("/auth/password-reset/confirm", response_model=schemas.MessageResponse)
def confirm_password_reset(
    payload: schemas.PasswordResetConfirm,
    db: Session = Depends(get_db),
):
    now = datetime.utcnow()
    reset_req, user = _get_valid_password_reset(db, payload.token)

    user.password_hash = get_password_hash(payload.new_password)
    reset_req.used_at = now
    db.query(models.PasswordResetToken).filter(
        models.PasswordResetToken.user_id == user.id,
        models.PasswordResetToken.id != reset_req.id,
    ).delete(synchronize_session=False)
    db.add(user)
    db.add(reset_req)
    db.commit()
    return schemas.MessageResponse(detail="Contraseña actualizada correctamente.")



@app.post("/license/activate", response_model=schemas.LicenseInfo)

def activate_license(

    request: schemas.LicenseActivationRequest,

    user: models.User = Depends(get_current_user),

    db: Session = Depends(get_db),

):

    license_obj = (

        db.query(models.License)

        .filter(models.License.code == request.code, models.License.user_id == user.id)

        .first()

    )

    if not license_obj:

        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Licencia no encontrada.")

    if not license_obj.is_active:

        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Licencia desactivada.")

    # Bloquear activación si la licencia ya expiró (2026-06-21)
    if license_obj.expires_at and license_obj.expires_at < datetime.utcnow():
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Licencia expirada. Contacta al administrador para renovarla.",
        )

    # Bloquear activación en otro equipo si ya fue vinculada
    if license_obj.device_fingerprint and license_obj.device_fingerprint != request.fingerprint:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Esta licencia ya está activada en otro equipo.",
        )

    device = (

        db.query(models.LicenseDevice)

        .filter(

            models.LicenseDevice.license_id == license_obj.id,

            models.LicenseDevice.fingerprint == request.fingerprint,

        )

        .first()

    )

    if not device and license_obj.device_fingerprint:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Esta licencia ya está activada en otro equipo.",
        )

    if not device:

        device = models.LicenseDevice(license=license_obj, fingerprint=request.fingerprint)

        db.add(device)

        if not license_obj.device_fingerprint:

            license_obj.device_fingerprint = request.fingerprint

    if not license_obj.activated_at:

        license_obj.activated_at = datetime.utcnow()

    db.add(license_obj)

    db.commit()

    db.refresh(license_obj)

    return license_obj



@app.post("/license/validate", response_model=schemas.LicenseInfo)

def validate_license(
    request: schemas.LicenseValidationRequest,

    user: models.User = Depends(get_current_user),

    db: Session = Depends(get_db),

):

    license_obj = (
        db.query(models.License)
        .join(models.LicenseDevice, models.License.id == models.LicenseDevice.license_id)
        .filter(
            models.License.user_id == user.id,
            models.License.is_active.is_(True),
            models.LicenseDevice.fingerprint == request.fingerprint,
        )
        .first()
    )
    if not license_obj:
        legacy_license = (
            db.query(models.License)
            .filter(
                models.License.user_id == user.id,
                models.License.device_fingerprint == request.fingerprint,
                models.License.is_active.is_(True),
            )
            .first()
        )
        if legacy_license:
            device = models.LicenseDevice(license=legacy_license, fingerprint=request.fingerprint)
            db.add(device)
            db.commit()
            db.refresh(legacy_license)
            license_obj = legacy_license
    if not license_obj:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Licencia no válida o no encontrada.")
    if license_obj.expires_at and license_obj.expires_at < datetime.utcnow():

        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Licencia expirada.")
    return license_obj

@app.get("/updates/latest")
def updates_latest(request: Request):
    _require_update_token(request)
    version = os.getenv("UPDATE_VERSION", "").strip()
    if not version:
        try:
            version = Path("version.txt").read_text(encoding="utf-8-sig").strip()
        except Exception:
            version = ""
    if not version:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="No hay actualizaciones disponibles.")
    file_url = os.getenv("UPDATE_FILE_URL", "").strip()
    file_path = os.getenv("UPDATE_FILE_PATH", "").strip()
    if not file_url and file_path:
        base = _public_base_url(request)
        file_url = f"{base}/updates/download"
    if not file_url and not file_path:
        r2_url = _r2_presigned_url()
        if r2_url:
            file_url = r2_url
    payload = {"version": version, "url": file_url}
    sha = os.getenv("UPDATE_SHA256", "").strip()
    if sha:
        payload["sha256"] = sha
    size_env = os.getenv("UPDATE_SIZE", "").strip()
    if size_env.isdigit():
        payload["size"] = int(size_env)
    elif file_path:
        try:
            payload["size"] = Path(file_path).stat().st_size
        except OSError:
            pass
    return payload


@app.get("/updates/download")
def updates_download(request: Request):
    _require_update_token(request)
    # Prioridad de fuentes:
    #   1) UPDATE_FILE_URL  → redirect 302 (e.g., GitHub Releases CDN).
    #      Recomendado: el cliente baja directo del CDN del proveedor sin
    #      consumir ancho de banda de Render.
    #   2) UPDATE_FILE_PATH → FileResponse desde el disco persistente.
    #   3) Cloudflare R2    → StreamingResponse (mantenido por compat).
    file_url = os.getenv("UPDATE_FILE_URL", "").strip()
    if file_url:
        return RedirectResponse(url=file_url, status_code=status.HTTP_302_FOUND)
    file_path = os.getenv("UPDATE_FILE_PATH", "").strip()
    if file_path:
        return FileResponse(
            file_path,
            filename="ROBOT_AUDIT_SRI.exe",
            media_type="application/x-msdownload",
            headers={"Content-Disposition": 'attachment; filename="ROBOT_AUDIT_SRI.exe"'},
        )
    client, bucket, object_key = _r2_download_source()
    if client and bucket and object_key:
        try:
            r2_obj = client.get_object(Bucket=bucket, Key=object_key)
        except Exception:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Archivo de actualizacion no encontrado en R2.")

        body = r2_obj.get("Body")
        if body is None:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="No se pudo leer el archivo de actualizacion en R2.")
        filename = Path(object_key).name or "ROBOT_AUDIT_SRI.exe"
        media_type = r2_obj.get("ContentType") or "application/x-msdownload"
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
        }
        content_length = r2_obj.get("ContentLength")
        if content_length is not None:
            headers["Content-Length"] = str(content_length)
        return StreamingResponse(
            _iter_r2_body(body),
            media_type=media_type,
            headers=headers,
        )
    raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Archivo de actualizacion no configurado.")


@app.get("/manual")
def download_user_manual():
    """Sirve el manual de usuario en PDF.

    Prioridad de fuentes (espejo de /updates/download):
      1) UPDATE_MANUAL_URL (env var) → redirect 302.
         Nombre consistente con UPDATE_FILE_URL del .exe. RECOMENDADO:
         apuntar a un asset PDF en GitHub Releases para que el cliente
         baje del CDN de GitHub sin consumir bandwidth de Render.
      2) MANUAL_URL (env var legacy) → redirect 302.
         Soportado por backward compat con deploys que ya tienen
         configurada esa variable.
      3) MANUAL_USUARIO.pdf en la raiz del proyecto → FileResponse.
         Fallback ultimo: el archivo baked-in en la imagen Docker.
      4) Ninguno → 404.
    """
    manual_url = (
        os.getenv("UPDATE_MANUAL_URL", "").strip()
        or os.getenv("MANUAL_URL", "").strip()
    )
    if manual_url:
        return RedirectResponse(url=manual_url, status_code=status.HTTP_302_FOUND)
    candidates = [
        Path("MANUAL_USUARIO.pdf"),
        Path(__file__).resolve().parent.parent / "MANUAL_USUARIO.pdf",
    ]
    for path in candidates:
        if path.is_file():
            return FileResponse(
                path=str(path),
                media_type="application/pdf",
                filename="MANUAL_USUARIO.pdf",
            )
    raise HTTPException(
        status_code=status.HTTP_404_NOT_FOUND,
        detail="Manual de usuario no disponible en el servidor.",
    )


@app.get("/", response_class=HTMLResponse)
def landing_page(request: Request):
    """Sirve la landing nueva (licensing_api/landing/index.html) con la
    misma funcionalidad de antes: version dinamica + URL de descarga
    con token + boton "Elegir donde guardar" (showSaveFilePicker).
    Solo cambia el ENVOLTORIO visual (Tailwind + secciones extendidas
    del mockup en Proyectos-de-Claude-main/landing/).

    El HTML vive en disco y solo se sustituyen 2 placeholders:
      - __VERSION__       → version actual (env UPDATE_VERSION o version.txt)
      - __DOWNLOAD_URL__  → /updates/download[?token=…]
    """
    version = os.getenv("UPDATE_VERSION", "").strip()
    if not version:
        try:
            version = Path("version.txt").read_text(encoding="utf-8-sig").strip()
        except Exception:
            version = ""
    if not version:
        version = "desconocida"
    # URL RELATIVA (sin slash inicial) para que funcione tanto cuando la
    # landing se sirve desde Render directo (sri-robot-audit-ik01.onrender.com/)
    # como cuando se sirve via proxy desde audit-ia.ec/sri_robot_audit/landing.
    # Una URL absoluta al dominio (/updates/download) rompe el segundo caso
    # porque el browser la resolveria a audit-ia.ec/updates/download (404).
    download_url = "updates/download"
    token = os.getenv("UPDATE_TOKEN", "").strip()
    if token:
        download_url = f"{download_url}?token={token}"

    index_path = _LANDING_DIR / "index.html"
    try:
        html = index_path.read_text(encoding="utf-8")
    except Exception:
        # Fallback minimo si el archivo falta (p.ej. deploy mal copiado).
        return HTMLResponse(
            content=(
                "<html><body style='font-family:sans-serif;padding:2rem'>"
                "<h1>ROBOT AUDIT SRI</h1>"
                f"<p>Version actual: {version}</p>"
                f"<p><a href='{download_url}'>Descargar ROBOT_AUDIT_SRI.exe</a></p>"
                "</body></html>"
            )
        )
    html = html.replace("__VERSION__", version).replace("__DOWNLOAD_URL__", download_url)
    return HTMLResponse(content=html)


