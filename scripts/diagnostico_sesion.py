"""Diagnostico de la sesion guardada. NO imprime ningun token.

Responde tres preguntas que, juntas, dicen por que se cierra una sesion:

  1. La app que corre, ¿es la que trae el arreglo del refresh token?
  2. La sesion guardada, ¿tiene refresh token?
  3. ¿Cuando vence el token que hay en mano?

Sin refresh token la sesion se cae a los 60 minutos y no hay forma de evitarlo:
el access token del SRI vive una hora y sin el de refresco hay que volver a
pedir la contrasena.

Uso:
    python scripts/diagnostico_sesion.py
    python scripts/diagnostico_sesion.py "C:/ruta/donde/esta/el/exe"
"""
from __future__ import annotations

import base64
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# La version desde la que el cliente sabe renovar el token.
VERSION_CON_REFRESH = (2026, 9, 24)


def _raiz() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).expanduser()
    return Path(os.getenv("APP_RUNTIME_DIR") or Path(__file__).resolve().parent.parent)


def _vence(token: str) -> str:
    """Fecha de expiracion de un JWT. Solo lee `exp`, nada mas."""
    try:
        cuerpo = token.split(".")[1]
        cuerpo += "=" * (-len(cuerpo) % 4)
        exp = json.loads(base64.urlsafe_b64decode(cuerpo)).get("exp")
        if not exp:
            return "sin fecha de expiracion"
        cuando = datetime.fromtimestamp(exp, tz=timezone.utc).astimezone()
        falta = cuando - datetime.now().astimezone()
        horas = falta.total_seconds() / 3600
        estado = f"vence en {horas:.1f} h" if horas > 0 else f"VENCIDO hace {-horas:.1f} h"
        return f"{cuando:%d/%m/%Y %H:%M} ({estado})"
    except Exception:
        return "no se pudo leer la fecha"


raiz = _raiz()
print(f"Carpeta analizada: {raiz}\n")

# ---------------------------------------------------------------- version
print("1) Version de la aplicacion")
archivo_version = next(
    (p for p in (raiz / "version.txt", raiz.parent / "version.txt") if p.exists()), None
)
if not archivo_version:
    print("   no se encontro version.txt: indica la carpeta del exe como argumento")
else:
    version = archivo_version.read_text(encoding="utf-8-sig").strip()
    try:
        partes = tuple(int(x) for x in version.split(".")[:3])
        al_dia = partes >= VERSION_CON_REFRESH
    except Exception:
        al_dia = False
    print(f"   {version}   ->   {'trae el arreglo' if al_dia else 'ANTERIOR AL ARREGLO'}")
    if not al_dia:
        print("   Actualiza la app: sin esto la sesion se cae a la hora, siempre.")

# ---------------------------------------------------------------- sesion
print("\n2) Sesion guardada")
carpeta = Path(os.getenv("SESSION_CACHE_DIR") or (raiz / ".session_cache"))
caches = sorted(carpeta.glob("session_cache*.json")) if carpeta.is_dir() else []
if not caches:
    print(f"   no hay sesion guardada en {carpeta}")
    print("   (es normal si nunca iniciaste sesion en este equipo)")
for cache in caches:
    try:
        datos = json.loads(cache.read_text(encoding="utf-8"))
    except Exception as err:
        print(f"   {cache.name}: no se pudo leer ({err})")
        continue
    print(f"   {cache.name}")
    tiene_refresh = bool(datos.get("refresh_token"))
    for clave in ("auth_token", "refresh_token", "device_fingerprint",
                  "license_validated", "license_last_check", "license_valid_until"):
        marca = "si" if datos.get(clave) else "NO"
        print(f"      {clave:22} {marca}")
    print()
    if not tiene_refresh:
        print("      >>> SIN refresh token: esta sesion se va a cerrar a los 60 min.")
        print("          Cierra sesion y vuelve a entrar: al iniciar con la version")
        print("          nueva, el servidor entrega el token de refresco y no")
        print("          vuelve a pasar.\n")

    # ------------------------------------------------------------ tokens
    print("3) Vencimiento de los tokens")
    for clave in ("auth_token", "refresh_token"):
        valor = datos.get(clave)
        print(f"   {clave:15} {_vence(valor) if valor else 'no hay'}")
    print()
