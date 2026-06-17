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
            # Granularidad fina (paginacion en SRI). Cuando el robot se
            # interrumpe a mitad de un dia (por ej. timeout, SRI tira al
            # home, navegador cerrado), estos campos permiten retomar
            # exactamente desde la pagina y fila donde quedo, en lugar de
            # re-empezar el dia desde pag 1 fila 1. Se actualizan al
            # cerrar cada lote (cada 10 filas procesadas).
            "current_page": 1,           # pagina actual en la tabla del SRI (1-based)
            "current_row_index": 0,      # indice de fila ya completada (0-based)
            "total_rows_on_page": 0,     # total de filas en la pagina actual
        },
        "last_error": "",
        # cancel_reason: "user" si el usuario presiono Detener;
        # "error" o "" si fue una falla tecnica (timeout, navegador,
        # red, etc.). Permite auto-reanudacion SOLO en el segundo caso.
        "cancel_reason": "",
        # Contador de auto-reanudaciones para evitar loops infinitos
        # si el error es persistente.
        "auto_resume_attempts": 0,
    }


def update_checkpoint(path: str | Path, **changes: Any) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    data.update(changes)
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def mark_checkpoint_failed(
    path: str | Path,
    error_message: str,
    *,
    cancel_reason: str = "error",
) -> dict[str, Any] | None:
    """Marca el checkpoint como `failed`.

    cancel_reason:
      - "user"  → el usuario presiono Detener proceso. La app debe mostrar
                  el boton "Reanudar descarga" pero NO reanudar sola.
      - "error" → falla tecnica (timeout, navegador caido, red, etc.). La
                  app puede reanudar automaticamente al volver a abrir.
    """
    data = load_checkpoint(path)
    if not data:
        return None
    data["status"] = "failed"
    data["last_error"] = str(error_message or "").strip()
    data["cancel_reason"] = str(cancel_reason or "error").strip().lower()
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def mark_checkpoint_running(path: str | Path) -> dict[str, Any] | None:
    data = load_checkpoint(path)
    if not data:
        return None
    data["status"] = "running"
    data["last_error"] = ""
    # Al arrancar (manual o auto), limpiar el motivo previo para que el
    # proximo fallo sea evaluado independientemente.
    data["cancel_reason"] = ""
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data


def increment_auto_resume_attempts(path: str | Path) -> int:
    """Incrementa el contador de auto-reanudaciones y devuelve el nuevo
    valor. Se usa para limitar reanudaciones automaticas y evitar loops
    cuando el error de raiz es persistente (ej. portal SRI caido).
    """
    data = load_checkpoint(path)
    if not data:
        return 0
    current = int(data.get("auto_resume_attempts") or 0)
    data["auto_resume_attempts"] = current + 1
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return current + 1


def update_checkpoint_progress(
    path: str | Path,
    *,
    next_month: int | None,
    next_day: int | None,
    last_completed_day: int | None,
    last_completed_label: str,
    current_page: int | None = None,
    current_row_index: int | None = None,
    total_rows_on_page: int | None = None,
) -> dict[str, Any] | None:
    """Actualiza el progreso del checkpoint.

    Los 3 kwargs nuevos (`current_page`, `current_row_index`, `total_rows_on_page`)
    son OPCIONALES — si vienen en None NO se modifican los valores previos.
    Esto preserva compatibilidad con todas las llamadas existentes (Recibidos
    y cierres de mes en Emitidos) que solo manejan granularidad por dia.

    Llamada desde el flujo de Emitidos al cerrar cada lote (10 filas):
        update_checkpoint_progress(
            path, next_month=5, next_day=15, last_completed_day=None,
            last_completed_label="15/05/2024 (pag 8, lote 3 = 30/50 filas)",
            current_page=8, current_row_index=30, total_rows_on_page=50,
        )

    Al cerrar un dia completo (limpia el progreso fino para arrancar fresco
    el dia siguiente):
        update_checkpoint_progress(
            path, next_month=5, next_day=16, last_completed_day=15,
            last_completed_label="15/05/2024",
            current_page=1, current_row_index=0, total_rows_on_page=0,
        )
    """
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
    # Campos finos: solo updated si se pasa un valor (no None).
    if current_page is not None:
        progress["current_page"] = int(current_page)
    if current_row_index is not None:
        progress["current_row_index"] = int(current_row_index)
    if total_rows_on_page is not None:
        progress["total_rows_on_page"] = int(total_rows_on_page)
    data["status"] = "running"
    data["progress"] = progress
    data["last_error"] = ""
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    save_checkpoint(path, data)
    return data
