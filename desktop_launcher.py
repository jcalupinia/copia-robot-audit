import json
import os
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "desktop_config.json"


def _load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _try_open_browser(url: str, attempts: int = 20, delay: float = 0.3) -> None:
    for _ in range(attempts):
        try:
            webbrowser.open(url, new=0)
            return
        except Exception:
            time.sleep(delay)


def main():
    config = _load_config()
    os.environ.setdefault("PLAYWRIGHT_HEADLESS", "0")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("ENABLE_SESSION_CACHE", "1")

    if "LICENSE_API_URL" not in os.environ and config.get("LICENSE_API_URL"):
        os.environ["LICENSE_API_URL"] = config["LICENSE_API_URL"]

    if "SESSION_CACHE_DIR" not in os.environ and config.get("SESSION_CACHE_DIR"):
        os.environ["SESSION_CACHE_DIR"] = config["SESSION_CACHE_DIR"]

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(BASE_DIR / "aplicacion.py"),
        "--server.address=127.0.0.1",
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]

    proc = subprocess.Popen(cmd)

    # Wait for Streamlit to be ready before opening browser
    _try_open_browser("http://127.0.0.1:8501")

    raise SystemExit(proc.wait())


if __name__ == "__main__":
    main()
