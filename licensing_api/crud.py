from __future__ import annotations

from typing import Optional
from sqlalchemy.orm import Session

from licensing_api import models
from licensing_api.security import get_password_hash


def get_user_by_email(db: Session, email: str) -> Optional[models.User]:
    return db.query(models.User).filter(models.User.email == email).first()


def create_user(db: Session, email: str, password: str) -> models.User:
    user = models.User(email=email, password_hash=get_password_hash(password))
    db.add(user)
    db.commit()
    db.refresh(user)
    return user


def create_license(
    db: Session,
    user: models.User,
    code: str,
    expires_at=None,
) -> models.License:
    license_obj = models.License(code=code, user=user, expires_at=expires_at)
    db.add(license_obj)
    db.commit()
    db.refresh(license_obj)
    return license_obj
