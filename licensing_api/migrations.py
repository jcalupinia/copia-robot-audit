"""Migraciones idempotentes de esquema y bootstrap de datos.

Se corre en el startup de main.py DESPUES de `Base.metadata.create_all(...)`.
SQLAlchemy's create_all solo crea tablas nuevas; NO agrega columnas a tablas
existentes. Para eso hacemos ALTER TABLE manual con detección por dialecto.

Bootstrap:
  - Crea los 3 usuarios admin iniciales si no existen (o cambia su rol si
    ya existen con otro rol).
  - Password inicial: leida de ADMIN_BOOTSTRAP_PASSWORD env var. Si no
    esta seteada, salta la creacion y solo cambia rol de los existentes.

Idempotente: correr esto varias veces es seguro. No duplica users ni rompe
data existente.
"""
from __future__ import annotations

import logging
import os
from typing import Optional

from sqlalchemy import inspect, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session

from licensing_api import models
from licensing_api.security import get_password_hash


logger = logging.getLogger(__name__)


# Los 3 emails que deben quedar como admin. Fijos en el codigo porque son
# datos de negocio, no configuracion. Si mañana hay que sumar otro, se agrega
# a esta lista y se redeploya.
ADMIN_EMAILS = (
    "adminjcalupinia1@auditconsulting.ec",
    "adminkormaza2@auditconsulting.ec",
    "adminjvinueza3@auditconsulting.ec",
)


def _column_exists(engine: Engine, table_name: str, column_name: str) -> bool:
    """True si la columna existe en la tabla. Portable entre SQLite y Postgres."""
    try:
        insp = inspect(engine)
        cols = {c["name"] for c in insp.get_columns(table_name)}
        return column_name in cols
    except Exception as err:
        logger.warning(f"No se pudo inspeccionar {table_name}: {err}")
        return False


def _add_role_column_if_missing(engine: Engine) -> None:
    """Agrega la columna `role` a la tabla users si no existe. Idempotente."""
    if _column_exists(engine, "users", "role"):
        return
    dialect = engine.dialect.name
    logger.info(f"Migrando: agregando columna users.role (dialect={dialect})")
    # SQLite y Postgres soportan la misma sintaxis basica de ALTER TABLE ADD COLUMN
    # con DEFAULT. Postgres tambien acepta NOT NULL en la misma sentencia si hay
    # default. SQLite lo tolera igual.
    ddl = (
        "ALTER TABLE users ADD COLUMN role VARCHAR(20) "
        "NOT NULL DEFAULT 'operador'"
    )
    with engine.begin() as conn:
        conn.execute(text(ddl))
    logger.info("Migracion OK: columna users.role agregada con default 'operador'")


def _promote_admins(db: Session, bootstrap_password: Optional[str]) -> None:
    """Asegura que los 3 emails de ADMIN_EMAILS existan y tengan role='admin'.

    - Si el user existe con otro rol: UPDATE role='admin' (deja password intacta).
    - Si el user NO existe:
        * Si `bootstrap_password` esta seteada: crea el user con esa password + role='admin'.
        * Si NO esta seteada: log warning y salta (no lo crea).
    """
    for email in ADMIN_EMAILS:
        email_n = email.strip().lower()
        existing = (
            db.query(models.User).filter(models.User.email == email_n).first()
        )
        if existing:
            if existing.role != "admin":
                logger.info(
                    f"Bootstrap: cambiando rol de {email_n} de "
                    f"'{existing.role}' a 'admin'"
                )
                existing.role = "admin"
                db.add(existing)
            if not existing.is_active:
                logger.info(f"Bootstrap: reactivando {email_n}")
                existing.is_active = True
                db.add(existing)
        else:
            if bootstrap_password:
                logger.info(f"Bootstrap: creando admin {email_n} (nuevo)")
                new_user = models.User(
                    email=email_n,
                    password_hash=get_password_hash(bootstrap_password),
                    is_active=True,
                    role="admin",
                )
                db.add(new_user)
            else:
                logger.warning(
                    f"Bootstrap: {email_n} no existe en DB y ADMIN_BOOTSTRAP_PASSWORD "
                    f"no esta seteada -> NO se crea. Setea la env var y redeploy, "
                    f"o crea el usuario manualmente y correra el promote."
                )
    db.commit()


def _normalize_other_roles(db: Session) -> None:
    """Para usuarios que no son admin: si tienen role NULL o vacio, setear
    a 'operador'. La columna es NOT NULL con default 'operador' pero puede
    haber corridas viejas donde qued\xf3 vacio (edge case tras ALTER TABLE
    en SQLite legacy). Idempotente.
    """
    # Solo aplica si hay filas afectadas; UPDATE es no-op si no hay match.
    affected = (
        db.query(models.User)
        .filter(
            (models.User.role.is_(None)) | (models.User.role == "")
        )
        .update({"role": "operador"}, synchronize_session=False)
    )
    if affected:
        logger.info(f"Normalizacion: {affected} user(s) sin rol -> 'operador'")
    db.commit()


def run_migrations(engine: Engine, session_factory) -> None:
    """Punto de entrada: corre todas las migraciones idempotentes.

    `session_factory` es SessionLocal (para no importar aqui y evitar
    ciclos). Se llama desde main.py despues de Base.metadata.create_all().

    Ordenado:
      1. ALTER TABLE users ADD COLUMN role (si falta)
      2. Normalizar role='operador' para filas viejas con NULL/vacio
      3. Promote/create los 3 admins iniciales

    Robusto: si algo falla, log y sigue arrancando. Preferimos que el
    servidor levante y el admin arregle a mano, en vez de quedar caido.
    """
    logger.info("=== Iniciando migraciones ===")
    try:
        _add_role_column_if_missing(engine)
    except Exception as err:
        logger.error(f"Migracion 'add role column' fallo: {err}. Continuando.")

    bootstrap_pwd = (os.getenv("ADMIN_BOOTSTRAP_PASSWORD", "") or "").strip()
    if not bootstrap_pwd:
        logger.warning(
            "ADMIN_BOOTSTRAP_PASSWORD no esta seteada. Los admins que no "
            "existan en DB NO seran creados. Los que existan si seran "
            "promovidos a role='admin'."
        )
    else:
        logger.info(
            "ADMIN_BOOTSTRAP_PASSWORD detectada. Se creara/promovera a los "
            f"{len(ADMIN_EMAILS)} admins."
        )

    db = session_factory()
    try:
        _normalize_other_roles(db)
        _promote_admins(db, bootstrap_pwd or None)
    except Exception as err:
        logger.error(f"Bootstrap admins fallo: {err}")
        db.rollback()
    finally:
        db.close()
    logger.info("=== Migraciones completadas ===")
