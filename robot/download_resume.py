from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any


def safe_user_token(user_email: str | None) -> str:
    raw = (user_email or "").strip().lower()
    token = re.sub(r"[^a-z0-9._-]+", "_", raw).strip("_")
    return token or "default"


def checkpoint_path(base_dir: str | Path, user_email: str | None) -> Path:
    base = Path(base_dir)
    base.mkdir(parents=True, exist_ok=True)
    return base / f"download_resume_{safe_user_token(user_email)}.json"


def _serialize(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(k): _serialize(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_serialize(v) for v in value]
    return value


def serialize_params(params: dict[str, Any]) -> dict[str, Any]:
    return _serialize(params)


def deserialize_params(params: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(params, dict):
        return {}
    data = dict(params)
    destino = data.get("destino")
    if destino:
        data["destino"] = Path(destino)
    return data


def load_checkpoint(path: str | Path) -> dict[str, Any] | None:
    path = Path(path)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def save_checkpoint(path: str | Path, payload: dict[str, Any]) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(_serialize(payload), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def delete_checkpoint(path: str | Path | None) -> None:
    if not path:
        return
    try:
        Path(path).unlink(missing_ok=True)
    except Exception:
        pass


def build_checkpoint_payload(user_email: str | None, params: dict[str, Any]) -> dict[str, Any]:
    now = datetime.now().isoformat(timespec="seconds")
    clean_params = dict(params)
    clean_params.pop("checkpoint_path", None)
    clean_params.pop("resume_download", None)
    return {
        "version": 1,
        "status": "pending",
        "user_email": (user_email or "").strip().lower(),
        "created_at": now,
        "updated_at": now,
        "params": serialize_params(clean_params),
        "summary": {
            "origen": clean_params.get("origen"),
            "tipo": clean_params.get("tipo"),
            "anio": clean_params.get("anio"),
            "mes": clean_params.get("mes"),
            "mes_fin": clean_params.get("mes_fin"),
            "dia": clean_params.get("dia"),
            "estado_emitidos": clean_params.get("estado_emitidos"),
            "formatos": list(clean_params.get("formatos") or []),
        },
        "progress": {
            "next_month": clean_params.get("mes"),
            "next_day": clean_params.get("dia") if clean_params.get("dia") not in (None, "") else 0,
            "last_completed_day": None,
            "last_completed_label": "",
            "completed_days": 0,
        },
        "last_error": "",
    }


def update_checkpoint(path: str | Path, **changes: Any) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    data.update(changes)
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def mark_checkpoint_failed(path: str | Path, error_message: str) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    data["status"] = "failed"
    data["last_error"] = str(error_message or "").strip()
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def mark_checkpoint_running(path: str | Path) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    data["status"] = "running"
    data["last_error"] = ""
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def update_checkpoint_progress(
    path: str | Path,
    *,
    next_month: int | None,
    next_day: int | None,
    last_completed_day: int | None,
    last_completed_label: str,
) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    progress = data.get("progress") if isinstance(data.get("progress"), dict) else {}
    completed_days = int(progress.get("completed_days") or 0) + (1 if last_completed_day else 0)
    progress.update(
        {
            "next_month": next_month,
            "next_day": next_day,
            "last_completed_day": last_completed_day,
            "last_completed_label": last_completed_label,
            "completed_days": completed_days,
        }
    )
    data["status"] = "running"
    data["progress"] = progress
    data["last_error"] = ""
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data
