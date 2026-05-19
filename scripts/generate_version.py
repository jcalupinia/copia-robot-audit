from __future__ import annotations

from datetime import datetime
from pathlib import Path
import subprocess


def _git_commit_count() -> str | None:
    try:
        count = subprocess.check_output(
            ["git", "rev-list", "--count", "HEAD"],
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except Exception:
        return None
    return count or None


def main() -> None:
    today = datetime.utcnow().strftime("%Y.%m.%d")
    count = _git_commit_count() or "0"
    version = f"{today}.{count}"
    Path("version.txt").write_text(version + "\n", encoding="utf-8")
    print(f"Version generada: {version}")


if __name__ == "__main__":
    main()