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


APP_NAME = "ROBOT_AUDIT_SRI"


def _raiz() -> Path:
    """Carpeta desde la que REALMENTE corre la app.

    El .exe se guarda donde sea, pero al abrirlo se copia a
    %LOCALAPPDATA%\\ROBOT_AUDIT_SRI y se relanza desde ahi. La sesion y las
    preferencias viven en esa carpeta, no en la de Descargas: buscarlas junto
    al archivo descargado no encuentra nada, y parece que el usuario nunca
    hubiera iniciado sesion.
    """
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).expanduser()
    if os.getenv("APP_RUNTIME_DIR"):
        return Path(os.environ["APP_RUNTIME_DIR"])
    local = os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")
    instalada = Path(local) / APP_NAME
    if instalada.is_dir():
        return instalada
    return Path(__file__).resolve().parent.parent


def _carpeta_cache(raiz: Path) -> Path:
    """Donde guarda la sesion, respetando `desktop_config.json` si existe."""
    if os.getenv("SESSION_CACHE_DIR"):
        return Path(os.environ["SESSION_CACHE_DIR"])
    config = raiz / "desktop_config.json"
    if config.exists():
        try:
            valor = (
                json.loads(config.read_text(encoding="utf-8-sig"))
                .get("SESSION_CACHE_DIR") or ""
            ).strip()
        except Exception:
            valor = ""
        if valor:
            ruta = Path(valor)
            return ruta if ruta.is_absolute() else (raiz / ruta)
    return raiz / ".session_cache"


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
print(f"Carpeta analizada: {raiz}")
print("(la app corre desde %LOCALAPPDATA%\\ROBOT_AUDIT_SRI, no desde donde")
print(" guardaste el .exe: al abrirlo se copia ahi y se relanza)\n")

# ---------------------------------------------------------------- version
print("1) Version de la aplicacion")
archivo_version = next(
    (p for p in (raiz / "version.txt", raiz.parent / "version.txt") if p.exists()), None
)
if not archivo_version:
    # En una instalacion real version.txt viaja DENTRO del .exe y se extrae a
    # una carpeta temporal, asi que aca no esta. La fecha del ejecutable sirve
    # igual para saber de cuando es la build.
    exe = raiz / "ROBOT_AUDIT_SRI.exe"
    if exe.exists():
        cuando = datetime.fromtimestamp(exe.stat().st_mtime)
        print(f"   ejecutable del {cuando:%d/%m/%Y %H:%M}")
        print("   La version exacta se lee en la barra superior de la app.")
    else:
        print(f"   no se encontro el ejecutable en {raiz}")
        print("   Indica la carpeta como argumento si la instalaste en otro lado.")
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
carpeta = _carpeta_cache(raiz)
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

# ---------------------------------------------------------------- bitacora
print("4) Ultimos eventos de sesion")
bitacora = carpeta / "sesion_eventos.log"
if not bitacora.exists():
    print(f"   no hay bitacora en {bitacora}")
    print("   (la escribe la version nueva; si falta, la app es anterior)")
else:
    lineas = bitacora.read_text(encoding="utf-8", errors="replace").splitlines()
    for linea in lineas[-15:]:
        print(f"   {linea}")
    if not lineas:
        print("   la bitacora esta vacia")
print()
