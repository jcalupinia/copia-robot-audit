from __future__ import annotations


import base64
from datetime import datetime, timedelta
from email.message import EmailMessage
import hashlib
import os
from pathlib import Path
import secrets
import smtplib
import ssl

from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.security import OAuth2PasswordBearer
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse

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


oauth2_scheme = OAuth2PasswordBearer(tokenUrl="auth/login")


app = FastAPI(

    title="Licensing API",

    version="1.0.0",

    description="API para autenticación y activación de licencias del robot.",

)


Base.metadata.create_all(bind=engine)


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

    # Bloquear activaci?n en otro equipo si ya fue vinculada
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
        base = str(request.base_url).rstrip("/")
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


def _landing_logo_data_uri() -> str:
    logo_env = os.getenv("LANDING_LOGO_PATH", "").strip()
    candidates = []
    if logo_env:
        candidates.append(Path(logo_env))
    candidates.extend(
        [
            Path("AUDIT_IA_sin_fondo_transparente_FINAL.png"),
            Path("LogoAUDIT.png"),
            Path("logo.png"),
        ]
    )
    for path in candidates:
        try:
            if not path.exists():
                continue
            ext = path.suffix.lower()
            if ext == ".svg":
                mime = "image/svg+xml"
            elif ext == ".jpg" or ext == ".jpeg":
                mime = "image/jpeg"
            else:
                mime = "image/png"
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            return f"data:{mime};base64,{encoded}"
        except Exception:
            continue
    return ""


@app.get("/", response_class=HTMLResponse)
def landing_page(request: Request):
    version = os.getenv("UPDATE_VERSION", "").strip()
    if not version:
        try:
            version = Path("version.txt").read_text(encoding="utf-8-sig").strip()
        except Exception:
            version = ""
    if not version:
        version = "desconocida"
    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/updates/download"
    token = os.getenv("UPDATE_TOKEN", "").strip()
    if token:
        download_url = f"{download_url}?token={token}"
    logo_uri = _landing_logo_data_uri()
    logo_html = (
        f"<img src='{logo_uri}' alt='Audit IA' class='logo-img'/>"
        if logo_uri
        else "<div class='logo-fallback'>AUDIT IA</div>"
    )
    html = f"""
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ROBOT AUDIT SRI</title>
  <style>
    :root {{
      --brand-navy: #0b1c54;
      --brand-blue: #2563eb;
      --brand-cyan: #16c7d7;
      --brand-soft: #eaf4ff;
      --ink: #0f172a;
      --muted: #475569;
      --white: #ffffff;
      --shadow: rgba(15, 23, 42, 0.16);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Calibri", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(1200px 620px at 12% 12%, rgba(37,99,235,0.20), transparent 60%),
        radial-gradient(900px 500px at 88% 14%, rgba(22,199,215,0.24), transparent 60%),
        linear-gradient(160deg, #f4f9ff 0%, #eaf5ff 46%, #f7fbff 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 28px;
    }}
    .shell {{
      max-width: 1080px;
      width: 100%;
      display: grid;
      grid-template-columns: 1.2fr 0.8fr;
      gap: 24px;
      background: rgba(255,255,255,0.88);
      border-radius: 26px;
      box-shadow: 0 22px 54px var(--shadow);
      padding: 30px;
      border: 1px solid rgba(37, 99, 235, 0.16);
      backdrop-filter: blur(6px);
    }}
    .hero {{
      padding: 8px 8px 8px 10px;
    }}
    .brand {{
      display: flex;
      align-items: center;
      gap: 14px;
      margin-bottom: 16px;
    }}
    .logo-img {{
      width: 64px;
      height: 64px;
      object-fit: contain;
      border-radius: 12px;
      background: var(--white);
      border: 1px solid rgba(37,99,235,0.16);
      padding: 6px;
    }}
    .logo-fallback {{
      width: 64px;
      height: 64px;
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: linear-gradient(145deg, var(--brand-blue), var(--brand-cyan));
      color: var(--white);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.08em;
    }}
    .brand-title {{
      margin: 0;
      font-size: 24px;
      font-weight: 800;
      letter-spacing: 0.02em;
      color: var(--brand-navy);
    }}
    .tag {{
      margin: 2px 0 0;
      color: var(--muted);
      font-size: 14px;
      font-weight: 500;
    }}
    .hero h1 {{
      margin: 10px 0 10px;
      font-size: clamp(32px, 4.2vw, 46px);
      line-height: 1.08;
      color: #0f1f57;
      letter-spacing: -0.02em;
    }}
    .lead {{
      margin: 0 0 20px;
      font-size: 18px;
      line-height: 1.6;
      color: var(--muted);
      max-width: 90%;
    }}
    .benefits {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }}
    .benefit {{
      background: var(--brand-soft);
      border: 1px solid rgba(37,99,235,0.16);
      border-radius: 14px;
      padding: 14px 14px;
    }}
    .benefit b {{
      display: block;
      margin-bottom: 4px;
      color: #123375;
      font-size: 15px;
    }}
    .benefit span {{
      color: #46607f;
      font-size: 14px;
      line-height: 1.4;
    }}
    .download-card {{
      background: linear-gradient(170deg, #0f2d70 0%, #0b1f52 62%, #0a1b48 100%);
      color: var(--white);
      border-radius: 22px;
      padding: 26px 24px;
      box-shadow: 0 18px 40px rgba(11, 28, 84, 0.35);
      border: 1px solid rgba(115, 169, 255, 0.26);
    }}
    .download-card h2 {{
      margin: 0 0 8px;
      font-size: 30px;
      line-height: 1.05;
      letter-spacing: -0.02em;
    }}
    .download-card p {{
      margin: 0;
      color: rgba(230,240,255,0.92);
      font-size: 15px;
      line-height: 1.5;
    }}
    .cta {{
      margin-top: 20px;
      width: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      padding: 18px 18px;
      border-radius: 16px;
      background: linear-gradient(140deg, var(--brand-cyan), #42d9e6 45%, #6be4ee 100%);
      color: #08283d;
      font-weight: 800;
      font-size: 18px;
      letter-spacing: 0.01em;
      border: 0;
      box-shadow: 0 10px 24px rgba(13, 216, 235, 0.32);
      transition: transform 0.18s ease, filter 0.18s ease;
    }}
    .cta:hover {{
      transform: translateY(-1px);
      filter: brightness(1.04);
    }}
    .cta-secondary {{
      margin-top: 10px;
      width: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      padding: 13px 14px;
      border-radius: 14px;
      background: rgba(122, 177, 255, 0.12);
      color: #d6e9ff;
      font-weight: 700;
      font-size: 14px;
      border: 1px solid rgba(122, 177, 255, 0.26);
      transition: transform 0.18s ease, filter 0.18s ease, background 0.18s ease;
    }}
    .cta-secondary:hover {{
      transform: translateY(-1px);
      filter: brightness(1.04);
      background: rgba(122, 177, 255, 0.18);
    }}
    .version {{
      margin-top: 14px;
      padding: 11px 12px;
      border-radius: 12px;
      background: rgba(122, 177, 255, 0.14);
      border: 1px solid rgba(122, 177, 255, 0.26);
      font-weight: 600;
      color: #d6e9ff;
      font-size: 14px;
    }}
    .note {{
      margin-top: 12px;
      font-size: 13px;
      color: rgba(230,240,255,0.85);
    }}
    .trust {{
      margin-top: 16px;
      padding-top: 14px;
      border-top: 1px solid rgba(122, 177, 255, 0.24);
      font-size: 14px;
      color: rgba(215, 232, 255, 0.9);
    }}
    @media (max-width: 860px) {{
      .shell {{
        grid-template-columns: 1fr;
        padding: 22px;
      }}
      .lead {{
        max-width: 100%;
      }}
      .benefits {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <div class="brand">
        {logo_html}
        <div>
          <h3 class="brand-title">ROBOT AUDIT SRI</h3>
          <p class="tag">Audit Consulting</p>
        </div>
      </div>
      <h1>Controla tus comprobantes en minutos.</h1>
      <p class="lead">
        Una sola aplicacion para ordenar tu gestion, ahorrar tiempo y mantener tus reportes siempre al dia.
      </p>
      <div class="benefits">
        <div class="benefit">
          <b>Rapido y simple</b>
          <span>Reduce tareas repetitivas y gana tiempo desde el primer dia.</span>
        </div>
        <div class="benefit">
          <b>Mas claridad</b>
          <span>Revisa tu informacion organizada para tomar decisiones con confianza.</span>
        </div>
        <div class="benefit">
          <b>Mejor control</b>
          <span>Ten todo centralizado para que tu seguimiento sea mas facil.</span>
        </div>
        <div class="benefit">
          <b>Siempre al dia</b>
          <span>Recibe mejoras continuas sin complicaciones para tu equipo.</span>
        </div>
      </div>
    </section>
    <aside class="download-card">
      <h2>Descarga ahora</h2>
      <p>Instala tu software de forma inmediata y empieza a usarlo hoy mismo.</p>
      <a class="cta" id="download-direct" href="{download_url}" download="ROBOT_AUDIT_SRI.exe">Descargar ROBOT_AUDIT_SRI.exe</a>
      <button class="cta-secondary" id="save-btn" type="button">Elegir donde guardar</button>
      <div class="version">Version actual: {version}</div>
      <p class="note">Usa el boton principal para descargar normalmente. Si tu navegador lo permite, tambien puedes elegir la ubicacion de guardado.</p>
      <p class="note" id="download-status" aria-live="polite"></p>
      <div class="trust">Solucion profesional para equipos que buscan orden, velocidad y confianza.</div>
    </aside>
  </div>
  <script>
    (function () {{
      const saveBtn = document.getElementById("save-btn");
      const directLink = document.getElementById("download-direct");
      const status = document.getElementById("download-status");
      const fileName = "ROBOT_AUDIT_SRI.exe";
      const downloadUrl = directLink ? directLink.href : "";

      function setStatus(msg) {{
        if (status) status.textContent = msg || "";
      }}

      async function saveWithPicker() {{
        const handle = await window.showSaveFilePicker({{
          suggestedName: fileName,
          excludeAcceptAllOption: true,
          types: [{{
            description: "Application",
            accept: {{
              "application/x-msdownload": [".exe"]
            }}
          }}]
        }});

        const response = await fetch(downloadUrl, {{ credentials: "same-origin" }});
        if (!response.ok) {{
          throw new Error("No se pudo iniciar la descarga.");
        }}
        if (!response.body) {{
          throw new Error("El navegador no devolvio el flujo del archivo.");
        }}
        const writable = await handle.createWritable();
        let closed = false;
        try {{
          const reader = response.body.getReader();
          while (true) {{
            const {{ done, value }} = await reader.read();
            if (done) break;
            if (value) {{
              await writable.write(value);
            }}
          }}
          await writable.close();
          closed = true;
        }} catch (err) {{
          if (!closed) {{
            try {{
              await writable.abort();
            }} catch (_abortErr) {{
              // Ignorado a proposito.
            }}
          }}
          throw err;
        }}
      }}

      if (!saveBtn) {{
        return;
      }}

      saveBtn.addEventListener("click", async function () {{
        if (!downloadUrl) {{
          setStatus("No se encontro el archivo de descarga.");
          return;
        }}
        saveBtn.disabled = true;
        setStatus("Preparando descarga...");
        try {{
          if (!(window.isSecureContext && "showSaveFilePicker" in window)) {{
            setStatus("Tu navegador no permite elegir ubicacion desde este boton. Usa la descarga principal.");
            return;
          }}
          await saveWithPicker();
          setStatus("Descarga completada.");
        }} catch (err) {{
          if (err && err.name === "AbortError") {{
            setStatus("Descarga cancelada.");
          }} else {{
            setStatus("No se pudo completar la descarga desde este boton. Usa la descarga principal.");
          }}
        }} finally {{
          saveBtn.disabled = false;
        }}
      }});
    }})();
  </script>
</body>
</html>
    """
    return HTMLResponse(content=html)
