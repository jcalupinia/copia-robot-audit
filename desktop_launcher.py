import json
import os
import socket
import sys
import time
import threading
import webbrowser
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "desktop_config.json"
LOG_PATH = BASE_DIR / "desktop_launcher.log"


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


def _load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _port_open(host: str, port: int, timeout: float = 0.2) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def _open_browser_when_ready(url: str, timeout: float = 30.0) -> None:
    start = time.time()
    while time.time() - start < timeout:
        if _port_open("127.0.0.1", 8501):
            try:
                webbrowser.open(url, new=0)
            except Exception:
                pass
            return
        time.sleep(0.3)


def main():
    config = _load_config()
    os.environ.setdefault("PLAYWRIGHT_HEADLESS", "0")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("ENABLE_SESSION_CACHE", "1")

    if "LICENSE_API_URL" not in os.environ and config.get("LICENSE_API_URL"):
        os.environ["LICENSE_API_URL"] = config["LICENSE_API_URL"]

    if "SESSION_CACHE_DIR" not in os.environ and config.get("SESSION_CACHE_DIR"):
        os.environ["SESSION_CACHE_DIR"] = config["SESSION_CACHE_DIR"]

    threading.Thread(
        target=_open_browser_when_ready,
        args=("http://127.0.0.1:8501",),
        daemon=True,
    ).start()

    try:
        import streamlit.web.cli as stcli
    except Exception as exc:
        _fatal("No se pudo iniciar Streamlit (faltan dependencias).", exc)

    sys.argv = [
        "streamlit",
        "run",
        str(BASE_DIR / "aplicacion.py"),
        "--server.address=127.0.0.1",
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]

    try:
        stcli.main()
    except SystemExit:
        raise
    except Exception as exc:
        _fatal("Error iniciando Streamlit.", exc)


if __name__ == "__main__":
    main()
