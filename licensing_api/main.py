from __future__ import annotations

from datetime import datetime
import os

from fastapi import Depends, FastAPI, HTTPException, status, Header
from fastapi.security import OAuth2PasswordBearer
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


def require_admin_token(x_admin_token: str | None = Header(default=None)) -> None:
    expected = os.getenv("LICENSE_ADMIN_TOKEN", "")
    if not expected or not x_admin_token or x_admin_token != expected:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Admin token inválido.")


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
    device = (
        db.query(models.LicenseDevice)
        .filter(
            models.LicenseDevice.license_id == license_obj.id,
            models.LicenseDevice.fingerprint == request.fingerprint,
        )
        .first()
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


@app.post("/admin/create-user-license", response_model=schemas.UserInfo)
def admin_create_user_license(
    request: schemas.AdminCreateUserLicenseRequest,
    db: Session = Depends(get_db),
    _: None = Depends(require_admin_token),
):
    user = crud.get_user_by_email(db, email=request.email)
    if not user:
        user = crud.create_user(db, email=request.email, password=request.password)
    else:
        user.password_hash = lic_security.get_password_hash(request.password)
        db.add(user)
        db.commit()
        db.refresh(user)
    license_obj = (
        db.query(models.License)
        .filter(models.License.code == request.code, models.License.user_id == user.id)
        .first()
    )
    if not license_obj:
        crud.create_license(db, user=user, code=request.code, expires_at=request.expires_at)
    return user
