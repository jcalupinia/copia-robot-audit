import hashlib
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import threading
import webbrowser
from pathlib import Path
from urllib.parse import urljoin

import requests

APP_NAME = "ROBOT_AUDIT_SRI"
VERSION_FILENAME = "version.txt"

DEFAULT_LICENSE_API_URL = os.getenv("DEFAULT_LICENSE_API_URL", "https://sri-robot-audit-ik01.onrender.com")
DEFAULT_UPDATE_TOKEN = os.getenv("DEFAULT_UPDATE_TOKEN", "256ed0dd9849466ebd29888cebdafc52")

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if getattr(sys, "frozen", False):
    EXE_DIR = Path(sys.executable).resolve().parent
else:
    EXE_DIR = Path(__file__).resolve().parent
CONFIG_FILENAME = "desktop_config.json"
LOG_PATH = EXE_DIR / "desktop_launcher.log"
INSTALL_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / APP_NAME


def _default_config() -> dict:
    return {
        "LICENSE_API_URL": DEFAULT_LICENSE_API_URL,
        "UPDATE_TOKEN": DEFAULT_UPDATE_TOKEN,
        "SESSION_CACHE_DIR": ".session_cache",
    }


def _normalize_config(raw: dict | None) -> dict:
    cfg = {}
    if isinstance(raw, dict):
        cfg.update(raw)

    defaults = _default_config()
    for key, value in defaults.items():
        current = cfg.get(key)
        if current is None:
            cfg[key] = value
            continue
        if isinstance(current, str):
            current = current.strip()
        else:
            current = str(current)
        if key == "LICENSE_API_URL":
            cfg[key] = current or value
        elif key == "SESSION_CACHE_DIR":
            cfg[key] = current or ".session_cache"
        elif key == "UPDATE_TOKEN":
            cfg[key] = current or value
        else:
            cfg[key] = current
    return cfg


def _write_config(path: Path, config: dict) -> None:
    path.write_text(json.dumps(config, indent=2), encoding="utf-8")


def _set_console_title(title: str) -> None:
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(title)
    except Exception:
        pass


def _show_startup_message() -> None:
    _set_console_title(f"{APP_NAME} - Iniciando")
    print("Iniciando software, espere un momento...", flush=True)
    print("Preparando entorno y cargando la aplicacion.", flush=True)


def _load_version() -> str:
    candidates = [
        APP_DIR / VERSION_FILENAME,
        EXE_DIR / VERSION_FILENAME,
        Path(__file__).resolve().parent / VERSION_FILENAME,
    ]
    for path in candidates:
        if path.exists():
            try:
                return path.read_text(encoding="utf-8-sig").strip() or "0.0.0"
            except Exception:
                continue
    return "0.0.0"


APP_VERSION = _load_version()


def _fatal(message: str, exc: Exception | None = None) -> "SystemExit":
    try:
        LOG_PATH.write_text(f"{message}\n{exc or ''}\n", encoding="utf-8")
    except Exception:
        pass
    print(message)
    if exc:
        print(exc)
    try:
        input("Presiona Enter para salir...")
    except Exception:
        pass
    raise SystemExit(1)


def _log(message: str) -> None:
    try:
        LOG_PATH.write_text(f"{message}\n", encoding="utf-8")
    except Exception:
        pass


def _load_config():
    candidates = [
        EXE_DIR / CONFIG_FILENAME,
        Path.cwd() / CONFIG_FILENAME,
    ]
    for path in candidates:
        if path.exists():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8-sig"))
            except Exception as exc:
                _fatal(f"Config invalido en {path}.", exc)
            if not isinstance(loaded, dict):
                _fatal(f"Config invalido en {path}. Debe ser un objeto JSON.")
            normalized = _normalize_config(loaded)
            if normalized != loaded:
                try:
                    _write_config(path, normalized)
                except Exception:
                    pass
            return normalized
    _fatal(
        "Falta LICENSE_API_URL. Agrega esa URL en desktop_config.json "
        "(junto al .exe) y vuelve a abrir la app."
    )
    return {}


def _resolve_config_path() -> Path | None:
    candidates = [EXE_DIR / CONFIG_FILENAME, Path.cwd() / CONFIG_FILENAME]
    for path in candidates:
        if path.exists():
            return path
    return None


def _ensure_installed() -> None:
    if not getattr(sys, "frozen", False):
        return
    try:
        exe_path = Path(sys.executable).resolve()
    except Exception:
        return
    if exe_path.parent.resolve() == INSTALL_DIR.resolve():
        return

    try:
        INSTALL_DIR.mkdir(parents=True, exist_ok=True)
        target_exe = INSTALL_DIR / f"{APP_NAME}.exe"
        shutil.copy2(exe_path, target_exe)
        target_config = INSTALL_DIR / CONFIG_FILENAME

        config_src = _resolve_config_path()
        if config_src:
            shutil.copy2(config_src, target_config)
            try:
                loaded = json.loads(target_config.read_text(encoding="utf-8-sig"))
            except Exception:
                loaded = {}
            _write_config(target_config, _normalize_config(loaded))
        else:
            _write_config(target_config, _default_config())

        subprocess.Popen([str(target_exe)])
    except Exception as exc:
        _fatal("No se pudo instalar la app en AppData.", exc)
    raise SystemExit(0)


def _version_tuple(raw: str) -> tuple[int, ...] | None:
    if not raw:
        return None
    parts = re.findall(r"\d+", raw)
    if not parts:
        return None
    return tuple(int(p) for p in parts)


def _is_remote_newer(remote: str, local: str) -> bool:
    remote_tuple = _version_tuple(remote)
    local_tuple = _version_tuple(local)
    if remote_tuple and local_tuple:
        length = max(len(remote_tuple), len(local_tuple))
        remote_tuple = remote_tuple + (0,) * (length - len(remote_tuple))
        local_tuple = local_tuple + (0,) * (length - len(local_tuple))
        return remote_tuple > local_tuple
    return remote.strip() != local.strip()


def _download_file(url: str, dest: Path, token: str | None = None) -> None:
    headers = {}
    if token:
        headers["X-Update-Token"] = token
    with requests.get(url, headers=headers, stream=True, timeout=30) as response:
        if response.status_code >= 400:
            raise RuntimeError(f"Error descargando actualizacion: {response.status_code}")
        with dest.open("wb") as handle:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if chunk:
                    handle.write(chunk)


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _apply_update(new_exe: Path, hard_exit: bool = False) -> None:
    current_exe = Path(sys.executable).resolve()
    backup_exe = current_exe.with_suffix(".old.exe")
    bat_path = INSTALL_DIR / "update.bat"

    bat_path.write_text(
        "\n".join(
            [
                "@echo off",
                f"set \"NEW={new_exe}\"",
                f"set \"CUR={current_exe}\"",
                f"set \"OLD={backup_exe}\"",
                "ping 127.0.0.1 -n 3 > nul",
                ":loop",
                "del /f /q \"%OLD%\" > nul 2>&1",
                "move /Y \"%CUR%\" \"%OLD%\" > nul 2>&1",
                "move /Y \"%NEW%\" \"%CUR%\" > nul 2>&1",
                "if errorlevel 1 (",
                "  ping 127.0.0.1 -n 2 > nul",
                "  goto loop",
                ")",
                "start \"\" \"%CUR%\"",
            ]
        ),
        encoding="utf-8",
    )

    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(["cmd", "/c", str(bat_path)], creationflags=creation_flags)
    if hard_exit:
        os._exit(0)
    raise SystemExit(0)


def _check_for_update(config: dict) -> None:
    if not getattr(sys, "frozen", False):
        return
    update_url = (
        os.environ.get("UPDATE_URL")
        or (config.get("UPDATE_URL") if isinstance(config, dict) else None)
    )
    if not update_url:
        license_url = (
            os.environ.get("LICENSE_API_URL")
            or (config.get("LICENSE_API_URL") if isinstance(config, dict) else "")
        )
        if license_url:
            update_url = license_url.rstrip("/") + "/updates/latest"
    if not update_url:
        return

    token = os.environ.get("UPDATE_TOKEN")
    if not token and isinstance(config, dict):
        token = (config.get("UPDATE_TOKEN") or "").strip() or None

    headers = {}
    if token:
        headers["X-Update-Token"] = token

    try:
        response = requests.get(update_url, headers=headers, timeout=10)
        if response.status_code >= 400:
            return
        payload = response.json()
    except Exception:
        return

    remote_version = str(payload.get("version") or "").strip()
    if not remote_version:
        return
    if not _is_remote_newer(remote_version, APP_VERSION):
        return

    download_url = str(payload.get("url") or "").strip()
    if not download_url:
        return
    if not re.match(r"^https?://", download_url):
        download_url = urljoin(update_url, download_url)

    target = INSTALL_DIR / f"{APP_NAME}.new.exe"
    try:
        _log("Actualizando...")
        print("Actualizando...")
        _download_file(download_url, target, token=token)
        expected_sha = str(payload.get("sha256") or "").strip()
        if expected_sha:
            actual_sha = _sha256_file(target)
            if actual_sha.lower() != expected_sha.lower():
                target.unlink(missing_ok=True)
                return
        expected_size = payload.get("size")
        if isinstance(expected_size, int) and expected_size > 0:
            if target.stat().st_size != expected_size:
                target.unlink(missing_ok=True)
                return
        _apply_update(target)
    except Exception:
        try:
            target.unlink(missing_ok=True)
        except Exception:
            pass
        return


def _port_open(host: str, port: int, timeout: float = 0.2) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def _open_browser_when_ready(url: str, host: str, port: int, timeout: float = 30.0) -> None:
    start = time.time()
    while time.time() - start < timeout:
        if _port_open(host, port):
            try:
                webbrowser.open(url, new=0)
            except Exception:
                pass
            return
        time.sleep(0.3)


def _pick_port(preferred: int = 8501) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", preferred))
            return preferred
        except OSError:
            sock.bind(("127.0.0.1", 0))
            return sock.getsockname()[1]


def main():
    _show_startup_message()
    _ensure_installed()

    # Ensure bundled app files are importable
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
        os.environ["PYTHONPATH"] = str(APP_DIR)

    config = _load_config()
    os.environ.setdefault("APP_VERSION", APP_VERSION)
    os.environ.setdefault("UPDATE_IN_APP", "1")

    if os.environ.get("UPDATE_IN_APP", "1").lower() in ("0", "false", "no"):
        _check_for_update(config)
    os.environ.setdefault("PLAYWRIGHT_HEADLESS", "0")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("ENABLE_SESSION_CACHE", "1")
    os.environ.setdefault("RECIBIDOS_MANUAL_CONSULTA", "0")
    os.environ.setdefault("RECIBIDOS_CONSULTA_INTENTOS", "8")
    os.environ.setdefault("RECIBIDOS_CONSULTA_BACKOFF_BASE_SEC", "1.6")
    os.environ.setdefault("RECIBIDOS_AUTO_PRE_EXECUTE_MS", "500")
    os.environ.setdefault("RECIBIDOS_AUTO_POST_EXECUTE_MS", "350")
    os.environ.setdefault("RECIBIDOS_AUTO_RESULT_TIMEOUT_MS", "70000")

    if "LICENSE_API_URL" not in os.environ:
        license_url = (config.get("LICENSE_API_URL") or "").strip()
        if not license_url:
            _fatal("LICENSE_API_URL vacío en desktop_config.json.")
        os.environ["LICENSE_API_URL"] = license_url

    session_cache_dir = (config.get("SESSION_CACHE_DIR") or "").strip()
    if not session_cache_dir:
        session_cache_dir = str(EXE_DIR / ".session_cache")
    else:
        session_cache_path = Path(session_cache_dir)
        if not session_cache_path.is_absolute():
            session_cache_dir = str((EXE_DIR / session_cache_path).resolve())
    os.environ.setdefault("SESSION_CACHE_DIR", session_cache_dir)
    os.environ.setdefault("USER_PREFS_PATH", str(Path(session_cache_dir) / "user_prefs.json"))

    host = "127.0.0.1"
    port = _pick_port()
    url = f"http://{host}:{port}"
    print("Inicializando interfaz local...", flush=True)
    threading.Thread(
        target=_open_browser_when_ready,
        args=(url, host, port),
        daemon=True,
    ).start()

    try:
        import streamlit.web.cli as stcli
    except Exception as exc:
        _fatal("No se pudo iniciar Streamlit (faltan dependencias).", exc)

    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = host

    app_path = APP_DIR / "aplicacion.py"
    if not app_path.exists():
        _fatal(f"No se encontró aplicacion.py en {APP_DIR}")

    print("Abriendo navegador...", flush=True)
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address",
        host,
        "--server.port",
        str(port),
    ]

    try:
        stcli.main()
    except SystemExit:
        raise
    except Exception as exc:
        _fatal("Error iniciando Streamlit.", exc)


if __name__ == "__main__":
    main()
