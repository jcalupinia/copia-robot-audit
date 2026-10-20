"""La migracion del servidor en los equipos ya instalados.

El problema que esto protege: cada instalacion guarda su `LICENSE_API_URL` en
`%LOCALAPPDATA%\\ROBOT_AUDIT_SRI\\desktop_config.json`, y ese archivo le gana al
valor compilado en el .exe. La auto-actualizacion reemplaza el ejecutable pero
no toca el config. Sin la migracion de `HOSTS_RETIRADOS`, publicar un .exe
nuevo apuntando a otro servidor no mueve a nadie: los equipos siguen hablando
con el viejo para siempre, y apagarlo los deja sin validar licencia y sin forma
de enterarse de ninguna actualizacion -- su unica fuente de `/updates/latest`
es justamente la URL vieja. No hay recuperacion remota de eso.

Uso:
    python tests_migracion_servidor.py
"""
from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

import desktop_launcher as dl

_fallos: list[str] = []


def chequear(etiqueta: str, obtenido, esperado) -> None:
    if obtenido == esperado:
        print(f"  [ok ] {etiqueta}")
        return
    _fallos.append(etiqueta)
    print(f"  [FALLA] {etiqueta}")
    print(f"          esperado: {esperado!r}")
    print(f"          obtenido: {obtenido!r}")


def probar_migracion() -> None:
    print("migracion del host retirado")
    viejo = "https://sri-robot-audit-ik01.onrender.com"
    nuevo = dl.DEFAULT_LICENSE_API_URL

    chequear(
        "el servidor viejo ya no es el default",
        dl._host_de(nuevo) in dl.HOSTS_RETIRADOS,
        False,
    )
    chequear(
        "un config con el host viejo se migra",
        dl._normalize_config({"LICENSE_API_URL": viejo})["LICENSE_API_URL"],
        nuevo,
    )
    chequear(
        "con barra final tambien se migra",
        dl._normalize_config({"LICENSE_API_URL": viejo + "/"})["LICENSE_API_URL"],
        nuevo,
    )
    chequear(
        "con mayusculas en el host tambien",
        dl._normalize_config(
            {"LICENSE_API_URL": "https://SRI-ROBOT-AUDIT-IK01.ONRENDER.COM"}
        )["LICENSE_API_URL"],
        nuevo,
    )
    chequear(
        "es idempotente: el nuevo no se vuelve a tocar",
        dl._normalize_config({"LICENSE_API_URL": nuevo})["LICENSE_API_URL"],
        nuevo,
    )
    # Quien apunto su instalacion a otro lado a proposito -- un servidor de
    # pruebas, una instancia propia -- no debe verse arrastrado.
    chequear(
        "una URL ajena se respeta",
        dl._normalize_config({"LICENSE_API_URL": "https://mi-servidor.local"})[
            "LICENSE_API_URL"
        ],
        "https://mi-servidor.local",
    )
    chequear(
        "vacio cae al default",
        dl._normalize_config({"LICENSE_API_URL": ""})["LICENSE_API_URL"],
        nuevo,
    )
    chequear(
        "ausente cae al default",
        dl._normalize_config({})["LICENSE_API_URL"],
        nuevo,
    )
    for basura in (None, 12345, [], {}):
        try:
            dl._normalize_config({"LICENSE_API_URL": basura})
            print(f"  [ok ] {basura!r} no explota")
        except Exception as err:  # pragma: no cover
            _fallos.append(f"{basura!r} explota")
            print(f"  [FALLA] {basura!r} explota: {err}")
    print()


def probar_persistencia() -> None:
    """La migracion tiene que quedar escrita, no repetirse en cada arranque."""
    print("el config migrado se persiste en disco")
    viejo = "https://sri-robot-audit-ik01.onrender.com"
    with tempfile.TemporaryDirectory() as tmp:
        ruta = Path(tmp) / dl.CONFIG_FILENAME
        ruta.write_text(
            json.dumps({"LICENSE_API_URL": viejo, "SESSION_CACHE_DIR": ".session_cache"}),
            encoding="utf-8",
        )
        original = json.loads(ruta.read_text(encoding="utf-8"))
        normalizado = dl._normalize_config(original)
        chequear("cambia respecto del archivo", normalizado != original, True)
        # `_load_config` escribe cuando difiere; se reproduce ese paso.
        dl._write_config(ruta, normalizado)
        releido = json.loads(ruta.read_text(encoding="utf-8"))
        chequear(
            "quedo escrito el servidor nuevo",
            releido["LICENSE_API_URL"],
            dl.DEFAULT_LICENSE_API_URL,
        )
        chequear(
            "en el segundo arranque ya no hay nada que migrar",
            dl._normalize_config(releido) == releido,
            True,
        )
        chequear(
            "no se perdio otra clave del config",
            releido.get("SESSION_CACHE_DIR"),
            ".session_cache",
        )
    print()


def probar_url_de_actualizacion() -> None:
    """La URL de updates se deriva de LICENSE_API_URL: migra con ella."""
    print("la URL de actualizacion acompana a la migracion")
    cfg = dl._normalize_config(
        {"LICENSE_API_URL": "https://sri-robot-audit-ik01.onrender.com"}
    )
    derivada = cfg["LICENSE_API_URL"].rstrip("/") + "/updates/latest"
    chequear(
        "apunta al servidor nuevo",
        derivada,
        dl.DEFAULT_LICENSE_API_URL.rstrip("/") + "/updates/latest",
    )
    chequear("no quedo rastro del viejo", "onrender.com" in derivada, False)
    print()


if __name__ == "__main__":
    print(f"servidor actual : {dl.DEFAULT_LICENSE_API_URL}")
    print(f"retirados       : {', '.join(dl.HOSTS_RETIRADOS)}\n")
    probar_migracion()
    probar_persistencia()
    probar_url_de_actualizacion()
    if _fallos:
        print(f"FALLARON {len(_fallos)}:")
        for f in _fallos:
            print(f"  - {f}")
        sys.exit(1)
    print("Todo en orden.")
