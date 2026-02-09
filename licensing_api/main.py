from __future__ import annotations


from datetime import datetime
import os
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.security import OAuth2PasswordBearer
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

import boto3
from botocore.config import Config

from sqlalchemy.orm import Session


from licensing_api import schemas, models, crud

from licensing_api.database import Base, engine, get_db

from licensing_api.security import (

    verify_password,

    create_access_token,

    decode_access_token,

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
                "ResponseContentType": "application/octet-stream",
            },
            ExpiresIn=expires_in,
        )
    except Exception:
        return None

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
        return FileResponse(file_path, filename="ROBOT_AUDIT_SRI.exe", media_type="application/octet-stream")
    r2_url = _r2_presigned_url()
    if r2_url:
        return RedirectResponse(r2_url)
    raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Archivo de actualizacion no configurado.")

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
    html = f"""
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ROBOT AUDIT SRI</title>
  <style>
    :root {{
      --bg: #fdfaf6;
      --panel: #ffffff;
      --ink: #1b1b1b;
      --muted: #836f60;
      --accent: #e38a53;
      --accent-dark: #cf7a45;
      --shadow: rgba(17, 13, 9, 0.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Georgia", "Times New Roman", serif;
      color: var(--ink);
      background: radial-gradient(circle at top left, #fffdf7 0%, #faf6f1 45%, #f4ede4 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 32px;
    }}
    .shell {{
      max-width: 980px;
      width: 100%;
      display: grid;
      grid-template-columns: 1.1fr 0.9fr;
      gap: 28px;
      background: var(--panel);
      border-radius: 28px;
      box-shadow: 0 16px 36px var(--shadow);
      padding: 40px;
      border: 1px solid rgba(17, 13, 9, 0.08);
    }}
    .badge {{
      font-size: 13px;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--muted);
      margin-bottom: 16px;
    }}
    h1 {{
      font-size: 46px;
      margin: 0 0 12px;
    }}
    h2 {{
      margin: 0;
      font-size: 22px;
    }}
    p {{
      margin: 0 0 18px;
      color: var(--muted);
      line-height: 1.7;
      font-size: 17px;
    }}
    ul {{
      margin: 0 0 24px;
      padding-left: 18px;
      color: var(--muted);
      line-height: 1.7;
      font-size: 16px;
    }}
    .card {{
      background: #fff;
      border-radius: 22px;
      padding: 28px;
      box-shadow: inset 0 0 0 1px rgba(17, 13, 9, 0.08);
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 20px;
    }}
    .cta {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      padding: 16px 26px;
      border-radius: 16px;
      background: var(--accent);
      color: #fff;
      font-weight: 700;
      font-size: 16px;
      transition: transform 0.2s ease, background 0.2s ease;
    }}
    .cta:hover {{
      background: var(--accent-dark);
      transform: translateY(-2px);
    }}
    .info {{
      background: #fff2e5;
      border-radius: 14px;
      padding: 14px 16px;
      font-size: 14px;
      color: var(--muted);
    }}
    .meta {{
      font-size: 14px;
      color: var(--muted);
    }}
    @media (max-width: 860px) {{
      .shell {{
        grid-template-columns: 1fr;
        padding: 28px;
      }}
      h1 {{
        font-size: 36px;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section>
      <div class="badge">Descarga segura</div>
      <h1>ROBOT AUDIT SRI</h1>
      <p>Descarga el ejecutable oficial y mantente actualizado con la ultima version aprobada.</p>
      <ul>
        <li>Actualizacion silenciosa desde el propio ejecutable.</li>
        <li>Instalacion automatica en AppData sin permisos de admin.</li>
        <li>Compatible con Windows 10/11.</li>
      </ul>
    </section>
    <aside class="card">
      <div>
        <h2>Descarga directa</h2>
        <p class="meta">Haz clic para bajar el ejecutable.</p>
        <p class="meta">Version actual: {version}</p>
      </div>
      <a class="cta" href="{download_url}">Descargar ROBOT_AUDIT_SRI.exe</a>
      <div class="info">Si ya tienes la app instalada, solo abre el exe y se actualizara automaticamente.</div>
    </aside>
  </div>
</body>
</html>
    """
    return HTMLResponse(content=html)