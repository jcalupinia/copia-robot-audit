from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

from sqlalchemy.orm import Session

from licensing_api import models
from licensing_api.security import get_password_hash


VALID_ROLES = ("admin", "operador", "cliente")


def get_user_by_email(db: Session, email: str) -> Optional[models.User]:
    return db.query(models.User).filter(models.User.email == email).first()


def get_user_by_id(db: Session, user_id: int) -> Optional[models.User]:
    return db.query(models.User).filter(models.User.id == user_id).first()


def create_user(
    db: Session,
    email: str,
    password: str,
    role: str = "operador",
) -> models.User:
    if role not in VALID_ROLES:
        raise ValueError(f"Rol inválido: {role!r}. Debe ser uno de {VALID_ROLES}.")
    user = models.User(
        email=email,
        password_hash=get_password_hash(password),
        role=role,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return user


def set_user_role(db: Session, user: models.User, role: str) -> models.User:
    """Cambia el rol de un usuario. No hace verificaciones de invariante
    ('no bajar el ultimo admin'); esa logica vive en admin.py."""
    if role not in VALID_ROLES:
        raise ValueError(f"Rol inválido: {role!r}. Debe ser uno de {VALID_ROLES}.")
    user.role = role
    db.add(user)
    db.commit()
    db.refresh(user)
    return user


def count_active_admins(db: Session) -> int:
    """Cuenta admins activos. Se usa para bloquear la baja del ultimo admin."""
    return (
        db.query(models.User)
        .filter(
            models.User.role == "admin",
            models.User.is_active.is_(True),
        )
        .count()
    )


def create_license(
    db: Session,
    user: models.User,
    code: str,
    expires_at: Optional[datetime] = None,
    validity_days: Optional[int] = None,
) -> models.License:
    """Crea una licencia. Prioridad de expiracion:
      1) expires_at explicito (fecha manual)
      2) validity_days (calculado desde AHORA — fecha de asignacion)
      3) None (sin vencimiento)
    """
    if expires_at is None and validity_days:
        expires_at = datetime.utcnow() + timedelta(days=int(validity_days))
    license_obj = models.License(code=code, user=user, expires_at=expires_at)
    db.add(license_obj)
    db.commit()
    db.refresh(license_obj)
    return license_obj


def set_license_expires_at(
    db: Session,
    license_obj: models.License,
    expires_at: Optional[datetime],
) -> models.License:
    """Setea expires_at directamente. `None` deja la licencia sin vencimiento."""
    license_obj.expires_at = expires_at
    db.add(license_obj)
    db.commit()
    db.refresh(license_obj)
    return license_obj


def renew_license(
    db: Session,
    license_obj: models.License,
    additional_days: int,
    from_now: bool = True,
) -> models.License:
    """Extiende expires_at por N dias adicionales.

    Args:
      additional_days: cuantos dias sumar.
      from_now: si True, base = ahora. Si False, base = expires_at actual
                (util para renovar antes de que venza sin perder los dias que
                quedan).
    """
    if from_now:
        base = datetime.utcnow()
    else:
        base = license_obj.expires_at or datetime.utcnow()
    license_obj.expires_at = base + timedelta(days=int(additional_days))
    db.add(license_obj)
    db.commit()
    db.refresh(license_obj)
    return license_obj


def is_license_expired(license_obj: models.License) -> bool:
    """True si la licencia tiene fecha de expiracion y ya paso."""
    return bool(
        license_obj.expires_at
        and license_obj.expires_at < datetime.utcnow()
    )


def get_license_status(license_obj: models.License) -> str:
    """'activa' | 'vencida' | 'inactiva'.

    - inactiva: fue desactivada explicitamente (is_active=False).
    - vencida: activa pero paso su expires_at.
    - activa: is_active y no vencida.

    Uso en admin UI y en /license/activate y /license/validate.
    """
    if not license_obj.is_active:
        return "inactiva"
    if is_license_expired(license_obj):
        return "vencida"
    return "activa"


def days_remaining(license_obj: models.License) -> Optional[int]:
    """Dias restantes hasta el vencimiento, o None si no vence.
    Negativo si ya vencio.
    """
    if not license_obj.expires_at:
        return None
    delta = license_obj.expires_at - datetime.utcnow()
    return delta.days
