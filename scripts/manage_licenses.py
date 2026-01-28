from __future__ import annotations

import argparse
import secrets

from licensing_api import crud
from licensing_api.database import Base, engine, SessionLocal


def create_user_and_license(email: str, password: str, code: str | None = None):
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        user = crud.get_user_by_email(db, email=email)
        if user:
            raise SystemExit("El usuario ya existe.")
        user = crud.create_user(db, email=email, password=password)
        license_code = code or secrets.token_urlsafe(24)
        license_obj = crud.create_license(db, user, license_code)
        print("Usuario y licencia creados:")
        print("Email:", user.email)
        print("Licencia:", license_obj.code)
    finally:
        db.close()


def main():
    parser = argparse.ArgumentParser(description="Gestión manual de licencias.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_cmd = subparsers.add_parser("create", help="Crear usuario + licencia")
    create_cmd.add_argument("--email", required=True)
    create_cmd.add_argument("--password", required=True)
    create_cmd.add_argument("--code", required=False, help="Código personalizado")

    args = parser.parse_args()
    if args.command == "create":
        create_user_and_license(args.email, args.password, args.code)


if __name__ == "__main__":
    main()
