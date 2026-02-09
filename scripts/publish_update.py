from __future__ import annotations

import argparse
import hashlib
import os
from pathlib import Path

import boto3
import requests


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _load_version(version_path: Path) -> str:
    if version_path.exists():
        return version_path.read_text(encoding="utf-8-sig").strip()
    return ""


def _r2_client():
    account_id = os.getenv("R2_ACCOUNT_ID", "").strip()
    access_key = os.getenv("R2_ACCESS_KEY_ID", "").strip()
    secret_key = os.getenv("R2_SECRET_ACCESS_KEY", "").strip()
    endpoint = os.getenv("R2_ENDPOINT", "").strip()
    region = os.getenv("R2_REGION", "auto").strip() or "auto"
    if not endpoint and account_id:
        endpoint = f"https://{account_id}.r2.cloudflarestorage.com"
    if not (endpoint and access_key and secret_key):
        return None
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=region,
    )


def _upload_to_r2(exe_path: Path, object_key: str) -> None:
    bucket = os.getenv("R2_BUCKET", "").strip()
    if not bucket:
        raise RuntimeError("Falta R2_BUCKET.")
    client = _r2_client()
    if client is None:
        raise RuntimeError("Faltan credenciales R2.")
    extra = {"ContentType": "application/octet-stream"}
    cache_control = os.getenv("R2_CACHE_CONTROL", "").strip()
    if cache_control:
        extra["CacheControl"] = cache_control
    client.upload_file(str(exe_path), bucket, object_key, ExtraArgs=extra)


def _render_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }


def _render_update_env(api_key: str, service_id: str, env_vars: dict[str, str]) -> None:
    url = f"https://api.render.com/v1/services/{service_id}/env-vars"
    payload = [{"key": key, "value": value} for key, value in env_vars.items()]
    response = requests.put(url, json=payload, headers=_render_headers(api_key), timeout=20)
    if response.status_code >= 400:
        raise RuntimeError(f"Error actualizando env vars: {response.status_code} - {response.text}")


def _render_trigger_deploy(api_key: str, service_id: str) -> None:
    url = f"https://api.render.com/v1/services/{service_id}/deploys"
    response = requests.post(url, json={}, headers=_render_headers(api_key), timeout=20)
    if response.status_code >= 400:
        raise RuntimeError(f"Error al disparar deploy: {response.status_code} - {response.text}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Publica el exe en R2 y actualiza Render.")
    parser.add_argument("--exe", default="dist/ROBOT_AUDIT_SRI.exe")
    parser.add_argument("--version-file", default="version.txt")
    parser.add_argument("--object-key", default=os.getenv("R2_OBJECT_KEY") or "ROBOT_AUDIT_SRI.exe")
    args = parser.parse_args()

    exe_path = Path(args.exe)
    if not exe_path.exists():
        raise SystemExit(f"No se encontró el exe: {exe_path}")

    version = _load_version(Path(args.version_file))
    if not version:
        raise SystemExit("No se encontró version.txt.")

    sha256 = _sha256_file(exe_path)
    size = exe_path.stat().st_size

    print(f"Version: {version}")
    print(f"SHA256: {sha256}")
    print(f"Size: {size} bytes")

    # Subir a R2
    try:
        _upload_to_r2(exe_path, args.object_key)
        print("Subido a R2.")
    except Exception as exc:
        raise SystemExit(f"Error subiendo a R2: {exc}") from exc

    # Actualizar Render (si hay credenciales)
    render_key = os.getenv("RENDER_API_KEY", "").strip()
    render_service = os.getenv("RENDER_SERVICE_ID", "").strip()
    if render_key and render_service:
        env_vars = {
            "UPDATE_VERSION": version,
            "UPDATE_SHA256": sha256,
            "UPDATE_SIZE": str(size),
        }
        try:
            _render_update_env(render_key, render_service, env_vars)
            print("Env vars de Render actualizadas.")
        except Exception as exc:
            raise SystemExit(f"Error actualizando Render: {exc}") from exc

        if os.getenv("RENDER_TRIGGER_DEPLOY", "1").strip() in ("1", "true", "True", "yes", "YES"):
            try:
                _render_trigger_deploy(render_key, render_service)
                print("Deploy disparado en Render.")
            except Exception as exc:
                raise SystemExit(f"Error disparando deploy en Render: {exc}") from exc
    else:
        print("RENDER_API_KEY o RENDER_SERVICE_ID no configurado. Render no se actualiza.")


if __name__ == "__main__":
    main()
