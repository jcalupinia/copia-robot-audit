# =====================================================
#  SRI ROBOT AUDIT  APLICACIN PRINCIPAL STREAMLIT
# Versin estable para Render.com / Octubre 2025
# =====================================================

import os
import shutil
import hashlib
import sys
import platform
import time
import uuid
import json
import re
import unicodedata
import calendar
import base64
import contextlib
import functools
import html
import textwrap
from datetime import datetime, timedelta, date
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import requests
from urllib.parse import urljoin

from robot.downloader import (
    descargar_sri,
    set_user_notifier,
    MANUAL_CONSULTA_RECIBIDOS,
    _consolidar_reportes_excel,
    _extraer_datos_pdf_por_tipo_layout_first,
    _extraer_datos_pdf_retencion_emitido,
    _extraer_datos_pdf_nota_credito_emitido,
    _extraer_datos_pdf_nota_debito_emitido,
    _extraer_datos_pdf_factura_emitido,
    _extraer_datos_pdf_liquidacion_compra_emitido,
    _extraer_datos_xml_retencion_emitido,
    _extraer_datos_xml_nota_credito_emitido,
    _extraer_datos_xml_nota_debito_emitido,
    _extraer_datos_xml_factura_emitido,
    _extraer_datos_xml_liquidacion_compra_emitido,
    _guardar_reporte_pdf_excel,
    _guardar_reporte_pdf_retencion_excel,
    _guardar_reporte_pdf_retencion_emitidos_excel,
    _guardar_reporte_pdf_nota_credito_emitidos_excel,
    _guardar_reporte_pdf_nota_debito_emitidos_excel,
    _guardar_reporte_pdf_factura_emitidos_excel,
    _prefijo_tipo,
    _xml_files_por_tipo,
    _slug_tipo,
    EMITIDOS_FACTURA_REPORT_COLUMNS,
    EMITIDOS_RETENCION_REPORT_COLUMNS,
    PDF_REPORT_COLUMNS,
    RETENCION_REPORT_COLUMNS,
    TIPOS_MAP,
    ESTADOS_EMITIDOS_MAP,
)
from robot.parser import construir_reporte, _parse_recibido_xml
from robot.historial import registrar_descarga, obtener_historial   #  FIX import correcto
from robot.download_resume import (
    build_checkpoint_payload,
    checkpoint_path as build_download_checkpoint_path,
    delete_checkpoint as delete_download_checkpoint,
    deserialize_params as deserialize_download_params,
    load_checkpoint as load_download_checkpoint,
    mark_checkpoint_failed as mark_download_checkpoint_failed,
    mark_checkpoint_running as mark_download_checkpoint_running,
    save_checkpoint as save_download_checkpoint,
)
from licensing_client import LicensingClient
# Actualizador de escritorio
try:
    import desktop_launcher as _desktop_launcher
except Exception:
    _desktop_launcher = None

import asyncio
import threading
import queue

if os.name == "nt":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())


def _load_desktop_config() -> dict:
    if not getattr(sys, "frozen", False):
        return {}
    exe_dir = Path(sys.executable).resolve().parent
    candidates = [exe_dir / "desktop_config.json", Path.cwd() / "desktop_config.json"]
    for cand in candidates:
        if cand.exists():
            try:
                return json.loads(cand.read_text(encoding="utf-8-sig"))
            except Exception:
                return {}
    return {}


def _get_update_payload() -> dict | None:
    if _desktop_launcher is None:
        return None
    config = _load_desktop_config()
    update_url = os.getenv("UPDATE_URL") or config.get("UPDATE_URL")
    if not update_url:
        license_url = os.getenv("LICENSE_API_URL") or config.get("LICENSE_API_URL") or ""
        if license_url:
            update_url = license_url.rstrip("/") + "/updates/latest"
    if not update_url:
        return None
    token = os.getenv("UPDATE_TOKEN") or (config.get("UPDATE_TOKEN") or "")
    token = token.strip() or None
    headers = {}
    if token:
        headers["X-Update-Token"] = token
    try:
        resp = requests.get(update_url, headers=headers, timeout=8)
        if resp.status_code >= 400:
            return None
        payload = resp.json()
    except Exception:
        return None
    remote_version = str(payload.get("version") or "").strip()
    if not remote_version:
        return None
    local_version = os.getenv("APP_VERSION") or _desktop_launcher.APP_VERSION
    if not _desktop_launcher._is_remote_newer(remote_version, local_version):
        return None
    download_url = str(payload.get("url") or "").strip()
    if not download_url:
        return None
    if not re.match(r"^https?://", download_url):
        download_url = urljoin(update_url, download_url)
    payload["url"] = download_url
    if token:
        payload["token"] = token
    return payload


def _start_update(payload: dict) -> None:
    if _desktop_launcher is None:
        return
    target = _desktop_launcher.INSTALL_DIR / f"{_desktop_launcher.APP_NAME}.new.exe"
    _desktop_launcher._download_file(payload["url"], target, token=payload.get("token"))
    expected_sha = str(payload.get("sha256") or "").strip()
    if expected_sha:
        actual = _desktop_launcher._sha256_file(target)
        if actual.lower() != expected_sha.lower():
            target.unlink(missing_ok=True)
            return
    expected_size = payload.get("size")
    if isinstance(expected_size, int) and expected_size > 0:
        if target.stat().st_size != expected_size:
            target.unlink(missing_ok=True)
            return
    _desktop_launcher._apply_update(target, hard_exit=True)


def _auto_update_ui() -> None:
    if not getattr(sys, "frozen", False):
        return
    if _desktop_launcher is None:
        return
    if st.session_state.get("_update_checked"):
        msg = st.session_state.get("_update_message")
        if msg:
            st.info(msg)
        return
    st.session_state["_update_checked"] = True
    payload = _get_update_payload()
    if not payload:
        return
    msg = "Actualizando... la app se reiniciara en breve."
    st.session_state["_update_message"] = msg
    st.session_state["_update_modal"] = True
    st.info(msg)
    def _worker():
        try:
            _start_update(payload)
        except Exception:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def _render_update_modal() -> None:
    if not st.session_state.get("_update_modal"):
        return
    message = st.session_state.get("_update_message") or "Actualizando..."
    logo = _get_logo_data_uri()
    logo_html = f"<img src='{logo}' alt='Logo' style='width:72px;margin-bottom:12px'/>" if logo else ""
    st.markdown(
        f"""
<style>
.update-overlay {{
  position: fixed;
  inset: 0;
  background: rgba(6, 10, 18, 0.82);
  backdrop-filter: blur(6px);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9999;
}}
.update-card {{
  width: min(520px, 92vw);
  background: linear-gradient(160deg, #0f172a 0%, #0b1220 100%);
  border-radius: 22px;
  padding: 28px 30px;
  color: #f8fafc;
  box-shadow: 0 24px 60px rgba(2, 6, 23, 0.45);
  border: 1px solid rgba(148, 163, 184, 0.2);
  text-align: center;
}}
.update-title {{
  font-size: 20px;
  font-weight: 700;
  margin-bottom: 6px;
}}
.update-sub {{
  color: rgba(226, 232, 240, 0.8);
  font-size: 14px;
  margin-bottom: 16px;
}}
.update-spinner {{
  width: 42px;
  height: 42px;
  border-radius: 50%;
  border: 3px solid rgba(148, 163, 184, 0.3);
  border-top-color: #38bdf8;
  margin: 0 auto 14px;
  animation: spin 0.8s linear infinite;
}}
.update-progress {{
  height: 8px;
  border-radius: 999px;
  background: rgba(148, 163, 184, 0.18);
  overflow: hidden;
}}
.update-progress span {{
  display: block;
  height: 100%;
  width: 35%;
  background: linear-gradient(90deg, #38bdf8, #22d3ee, #a78bfa);
  animation: slide 1.2s ease-in-out infinite;
}}
@keyframes spin {{
  to {{ transform: rotate(360deg); }}
}}
@keyframes slide {{
  0% {{ transform: translateX(-120%); }}
  50% {{ transform: translateX(40%); }}
  100% {{ transform: translateX(120%); }}
}}
</style>
<div class="update-overlay">
  <div class="update-card">
    {logo_html}
    <div class="update-title">Actualizando ROBOT AUDIT SRI</div>
    <div class="update-sub">{message}</div>
    <div class="update-spinner"></div>
    <div class="update-progress"><span></span></div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def _render_manual_consultar_modal() -> None:
    msg = st.session_state.get("manual_consultar_hint")
    ts = st.session_state.get("manual_consultar_hint_ts")
    status = st.session_state.get("download_status")
    if not msg or not ts:
        return
    if status not in {"running", "cancelling"}:
        return
    # Modal temporal para avisar la accion manual requerida.
    if (time.time() - float(ts)) > 25:
        return
    st.markdown(
        f"""
<style>
.manual-consultar-overlay {{
  position: fixed;
  inset: 0;
  background: rgba(2, 6, 23, 0.45);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9990;
  pointer-events: none;
}}
.manual-consultar-card {{
  width: min(560px, 90vw);
  border-radius: 18px;
  padding: 20px 22px;
  background: linear-gradient(160deg, rgba(15, 23, 42, 0.96), rgba(10, 15, 28, 0.96));
  border: 1px solid rgba(148, 163, 184, 0.28);
  box-shadow: 0 24px 60px rgba(2, 6, 23, 0.45);
  color: #f8fafc;
}}
.manual-consultar-title {{
  font-size: 18px;
  font-weight: 800;
  margin-bottom: 6px;
}}
.manual-consultar-text {{
  font-size: 14px;
  color: rgba(226, 232, 240, 0.95);
  line-height: 1.45;
}}
.manual-consultar-help {{
  margin-top: 10px;
  font-size: 13px;
  color: rgba(148, 163, 184, 0.95);
}}
</style>
<div class="manual-consultar-overlay">
  <div class="manual-consultar-card">
    <div class="manual-consultar-title">Accion requerida</div>
    <div class="manual-consultar-text">Haz clic en <b>Consultar</b> en la ventana del navegador del SRI.</div>
    <div class="manual-consultar-help">{msg}</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def _render_download_finished_modal() -> None:
    if st.session_state.get("download_status") != "done":
        return
    if not st.session_state.get("download_finished_modal_open"):
        return

    params = st.session_state.get("download_params") or {}
    resultado = st.session_state.get("download_result") or {}
    origen = str(params.get("origen") or "").strip() or "Proceso"
    tipo = str(params.get("tipo") or "").strip() or "comprobantes"

    lineas = []
    if origen == "Emitidos":
        n_registros = int(resultado.get("n_registros", 0) or 0)
        n_xml = int(resultado.get("n_xml", 0) or 0)
        n_pdf = int(resultado.get("n_pdf", 0) or 0)
        lineas.append(f"Registros procesados: {n_registros}")
        if "XML" in (params.get("formatos") or []):
            lineas.append(f"XML descargados: {n_xml}")
        if "PDF" in (params.get("formatos") or []):
            lineas.append(f"PDF descargados: {n_pdf}")
    else:
        n_xml = int(resultado.get("n_xml", 0) or 0)
        n_pdf = int(resultado.get("n_pdf", 0) or 0)
        if "XML" in (params.get("formatos") or []):
            lineas.append(f"XML descargados: {n_xml}")
        if "PDF" in (params.get("formatos") or []):
            lineas.append(f"PDF descargados: {n_pdf}")
    carpeta_tipo = resultado.get("carpeta_tipo")
    if carpeta_tipo:
        lineas.append(f"Carpeta destino: {Path(carpeta_tipo)}")
    mensaje_verificacion = str(resultado.get("mensaje_verificacion") or "").strip()
    if mensaje_verificacion:
        lineas.append(f"Verificación: {mensaje_verificacion}")

    st.write(f"El proceso de {origen.lower()} para {tipo.lower()} ha finalizado correctamente.")
    for linea in lineas:
        st.caption(linea)
    if st.button("Cerrar", key="close_download_finished_modal", use_container_width=True):
        st.session_state["download_finished_modal_open"] = False
        st.rerun()


def _schedule_desktop_app_exit(delay_sec: float = 0.9) -> None:
    if st.session_state.get("_desktop_exit_scheduled"):
        return
    st.session_state["_desktop_exit_scheduled"] = True

    def _shutdown():
        time.sleep(max(0.1, float(delay_sec)))
        os._exit(0)

    threading.Thread(target=_shutdown, daemon=True).start()


def _render_close_app_modal() -> None:
    if not getattr(sys, "frozen", False):
        st.info("Esta opcion solo esta disponible en el ejecutable de escritorio.")
        return

    if st.session_state.get("_desktop_exit_in_progress"):
        st.success("Cerrando aplicacion...")
        st.caption("Se cerrara el servicio local y la ventana de comandos del ejecutable.")
        st.caption("Si la pestaña del navegador queda abierta, puedes cerrarla manualmente.")
        components.html(
            """
            <script>
            setTimeout(function() {
              try {
                window.open('', '_self');
                window.close();
              } catch (e) {}
            }, 250);
            </script>
            """,
            height=0,
        )
        st.stop()

    st.write("Esta accion cerrara la aplicacion completa en esta computadora.")
    st.caption("Usa esta opcion solo cuando quieras salir del sistema por completo.")
    col_cancel, col_confirm = st.columns(2)
    with col_cancel:
        if st.button("Cancelar", key="btn_cancel_close_app", use_container_width=True):
            st.rerun()
    with col_confirm:
        if st.button(
            "Cerrar aplicacion",
            key="btn_confirm_close_app",
            type="primary",
            use_container_width=True,
        ):
            st.session_state["_desktop_exit_in_progress"] = True
            _schedule_desktop_app_exit()
            st.rerun()



@st.cache_data(show_spinner=False)
def _get_logo_data_uri():
    logo_path = Path(__file__).parent / "AUDIT_IA_sin_fondo_transparente_FINAL.png"
    if not logo_path.exists():
        return None
    data = logo_path.read_bytes()
    encoded = base64.b64encode(data).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _logo_html(width):
    data_uri = _get_logo_data_uri()
    if not data_uri:
        return ""
    return f"<div style='text-align:center'><img src='{data_uri}' width='{width}'/></div>"


@functools.lru_cache(maxsize=2)
def _asset_data_uri(filename: str) -> str:
    """Devuelve un data URI base64 para un asset de la carpeta `assets/`.

    Cacheado a un solo I/O por archivo. Usado para hero/logo del mockup.
    """
    base_dir = Path(__file__).parent
    path = base_dir / "assets" / filename
    if not path.exists():
        return ""
    suffix = path.suffix.lower()
    mime = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".svg": "image/svg+xml",
        ".webp": "image/webp",
    }.get(suffix, "application/octet-stream")
    try:
        encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    except Exception:
        return ""
    return f"data:{mime};base64,{encoded}"


_group_card_counter = {"n": 0}


@contextlib.contextmanager
def _group_card(step: int | str, title: str, subtitle: str = ""):
    """Context manager para envolver una seccion como `.group` del mockup
    (card con borde + header numerado). Uso:

        with _group_card(1, "Credenciales", "Datos del SRI"):
            ...widgets de Streamlit...

    Pasa `key="group_card_N"` a st.container para que Streamlit emita
    `class="st-key-group_card_N"` en el wrapper stVerticalBlockBorderWrapper.
    El CSS targetea esa clase directamente — NO depende de :has() ni
    de la profundidad del marker en el DOM. El counter es por rerun
    (Streamlit reinicia el script, asi que el contador vuelve a 0).
    """
    _group_card_counter["n"] += 1
    container = st.container(
        border=True, key=f"group_card_{_group_card_counter['n']}"
    )
    with container:
        subtitle_html = f"<small>{html.escape(subtitle)}</small>" if subtitle else ""
        st.markdown(
            f"<div class='group-h-marker'>"
            f"<span class='step'>{step}</span>"
            f"<h3>{html.escape(str(title))}</h3>"
            f"{subtitle_html}"
            f"</div>",
            unsafe_allow_html=True,
        )
        yield container


def _slug_estado_emitidos(estado: str) -> str:
    estado_nombre = (ESTADOS_EMITIDOS_MAP.get(estado, estado) or "Sin Estado").strip() or "Sin Estado"
    estado_normalizado = (
        unicodedata.normalize("NFKD", estado_nombre).encode("ascii", "ignore").decode("ascii")
    )
    return re.sub(r"[^A-Za-z0-9]+", "_", estado_normalizado).strip("_") or "Sin_Estado"


def _buscar_reportes_mensuales(base_dir: Path, prefix: str) -> list[Path]:
    if not base_dir.exists():
        return []
    # Permite nombres con o sin guion bajo tras el anio:
    # - prefijoYYYYMM.xlsx
    # - prefijo_YYMM.xlsx
    # - prefijo_YYMMDD.xlsx
    regex = re.compile(
        r"^" + re.escape(prefix) + r"(?:_)?(\d{2}(?:\d{2})?)(?:_\d+)?\.xlsx$",
        re.IGNORECASE,
    )
    encontrados: list[Path] = []
    for ruta in base_dir.rglob(f"{prefix}*.xlsx"):
        if regex.match(ruta.name):
            encontrados.append(ruta)
    return sorted(encontrados, key=lambda p: p.name)


MESES_ES = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]

MESES_ES_MAP = {
    unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode("ascii").lower(): idx + 1
    for idx, nombre in enumerate(MESES_ES)
}


def _normalizar_texto_simple(valor: str) -> str:
    return unicodedata.normalize("NFKD", valor or "").encode("ascii", "ignore").decode("ascii").lower().strip()


def _mes_desde_texto(valor: str) -> int | None:
    texto = (valor or "").strip()
    if not texto:
        return None
    if texto.isdigit():
        mes_num = int(texto)
        if 1 <= mes_num <= 12:
            return mes_num
        return None
    return MESES_ES_MAP.get(_normalizar_texto_simple(texto))


def _parsear_token_fecha_reporte(token: str) -> tuple[int, int, int | None] | None:
    solo_digitos = re.sub(r"\D", "", token or "")
    if len(solo_digitos) == 6:
        anio = int(solo_digitos[:4])
        mes = int(solo_digitos[4:6])
        if 1 <= mes <= 12:
            return anio, mes, None
        return None
    if len(solo_digitos) == 8:
        anio = int(solo_digitos[:4])
        mes = int(solo_digitos[4:6])
        dia = int(solo_digitos[6:8])
        if 1 <= mes <= 12 and 1 <= dia <= 31:
            return anio, mes, dia
        dia = int(solo_digitos[:2])
        mes = int(solo_digitos[2:4])
        anio = int(solo_digitos[4:8])
        if 1 <= mes <= 12 and 1 <= dia <= 31:
            return anio, mes, dia
    return None


def _fecha_coincide_consolidacion(
    anio_archivo: int,
    mes_archivo: int,
    dia_archivo: int | None,
    modo_fecha: str,
    anio_objetivo: int,
    mes_inicio: int,
    mes_fin: int,
    dia_objetivo: int,
) -> bool:
    if anio_archivo != anio_objetivo:
        return False
    if modo_fecha == "Ano completo":
        return True
    if modo_fecha == "Rango de meses":
        return mes_inicio <= mes_archivo <= mes_fin
    if mes_archivo != mes_inicio:
        return False
    if dia_objetivo in (0, None):
        return True
    return dia_archivo == dia_objetivo


def _sufijo_periodo_consolidacion(
    modo_fecha: str,
    anio: int,
    mes_inicio: int,
    mes_fin: int,
    dia: int,
) -> str:
    if modo_fecha == "Ano completo":
        return f"{anio:04d}"
    if modo_fecha == "Rango de meses":
        return f"{anio:04d}{mes_inicio:02d}{mes_fin:02d}"
    if dia in (0, None):
        return f"{anio:04d}{mes_inicio:02d}"
    return f"{anio:04d}{mes_inicio:02d}{dia:02d}"


def _buscar_reportes_por_periodo(
    base_dir: Path,
    origen: str,
    tipo_slug: str,
    formato: str,
    modo_fecha: str,
    anio_objetivo: int,
    mes_inicio: int,
    mes_fin: int,
    dia_objetivo: int,
) -> list[Path]:
    if not base_dir.exists():
        return []
    formato_norm = (formato or "").strip().lower()
    if formato_norm not in {"xml", "pdf"}:
        return []
    origen_norm = (origen or "").strip().lower()
    patron_main = re.compile(
        r"^(recibidos|emitidos)_reporte_"
        + re.escape(formato_norm)
        + r"_"
        + re.escape(tipo_slug)
        + r"_(\d{6,8})(?:_\d+)?\.xlsx$",
        re.IGNORECASE,
    )
    patron_alt_xml_recibidos = None
    if formato_norm == "xml":
        patron_alt_xml_recibidos = re.compile(
            r"^reporte_"
            + re.escape(tipo_slug)
            + r"_(\d{4})_(\d{2})(?:_\d+)?\.xlsx$",
            re.IGNORECASE,
        )
    encontrados: list[Path] = []
    for ruta in base_dir.rglob("*.xlsx"):
        partes_norm = {_normalizar_texto_simple(parte) for parte in ruta.parts}
        if "consolidados" in partes_norm:
            continue
        nombre = ruta.name
        match_main = patron_main.match(nombre)
        if match_main:
            origen_arch = (match_main.group(1) or "").strip().lower()
            if origen_arch != origen_norm:
                continue
            fecha_info = _parsear_token_fecha_reporte(match_main.group(2))
            if not fecha_info:
                continue
            anio_arch, mes_arch, dia_arch = fecha_info
            if _fecha_coincide_consolidacion(
                anio_arch,
                mes_arch,
                dia_arch,
                modo_fecha,
                anio_objetivo,
                mes_inicio,
                mes_fin,
                dia_objetivo,
            ):
                encontrados.append(ruta)
            continue

        if patron_alt_xml_recibidos and origen_norm == "recibidos":
            match_alt = patron_alt_xml_recibidos.match(nombre)
            if not match_alt:
                continue
            anio_arch = int(match_alt.group(1))
            mes_arch = int(match_alt.group(2))
            if _fecha_coincide_consolidacion(
                anio_arch,
                mes_arch,
                None,
                modo_fecha,
                anio_objetivo,
                mes_inicio,
                mes_fin,
                dia_objetivo,
            ):
                encontrados.append(ruta)

    encontrados = sorted(dict.fromkeys(encontrados), key=lambda p: str(p).lower())
    return encontrados


def _extraer_fecha_desde_ruta_documento(
    ruta: Path,
    tipo_prefijo: str,
    tipo_slug_archivo: str | None = None,
) -> tuple[int, int, int | None] | None:
    # El slug usado para nombrar archivos al descargar viene de
    # `_slug_tipo(tipo_visible)` (p.ej. "nota_de_credito" para "Nota de Credito").
    # El derivado desde `tipo_prefijo` usa la etiqueta canonica del TIPO_LABEL_MAP
    # (p.ej. "NotaCredito" -> slug "notacredito"), que NO coincide con los
    # nombres de archivo reales. Aceptamos ambos por compatibilidad.
    slugs_validos: set[str] = set()
    if tipo_slug_archivo:
        slug_norm = _slug_tipo(tipo_slug_archivo)
        if slug_norm:
            slugs_validos.add(slug_norm)
    try:
        _, etiqueta_tipo = str(tipo_prefijo or "").split("_", 1)
        slug_derivado = _slug_tipo(etiqueta_tipo)
    except Exception:
        slug_derivado = _slug_tipo(tipo_prefijo or "")
    if slug_derivado:
        slugs_validos.add(slug_derivado)

    # Estructura nueva: .../<anio>/<mes>/<XML|PDF>/<tipo_slug>__YYYYMMDD__archivo.ext
    try:
        if ruta.parent.name.lower() in {"xml", "pdf"} and slugs_validos:
            mes_txt = ruta.parent.parent.name
            anio_txt = ruta.parent.parent.parent.name
            if anio_txt.isdigit():
                mes = _mes_desde_texto(mes_txt)
                if mes:
                    union_slugs = "|".join(re.escape(s) for s in slugs_validos)
                    patron_nombre = re.compile(
                        r"^(?:" + union_slugs + r")__(\d{8})__",
                        re.IGNORECASE,
                    )
                    match_nombre = patron_nombre.match(ruta.name)
                    if match_nombre:
                        fecha_token = match_nombre.group(1)
                        anio = int(fecha_token[:4])
                        mes_arch = int(fecha_token[4:6])
                        dia_arch = int(fecha_token[6:8])
                        if anio == int(anio_txt) and mes_arch == mes and 1 <= dia_arch <= 31:
                            return anio, mes_arch, dia_arch
    except Exception:
        pass

    # Estructura anterior: .../<anio>/<mes>/<dia>/<tipo_prefijo>/<XML|PDF>/archivo.ext
    try:
        tipo_dir = ruta.parent.parent
    except Exception:
        return None
    if tipo_dir.name != tipo_prefijo:
        return None
    try:
        dia_txt = tipo_dir.parent.name
        mes_txt = tipo_dir.parent.parent.name
        anio_txt = tipo_dir.parent.parent.parent.name
    except Exception:
        return None
    if not anio_txt.isdigit():
        return None
    anio = int(anio_txt)
    mes = _mes_desde_texto(mes_txt)
    if not mes:
        return None
    dia: int | None = None
    if dia_txt.isdigit():
        dia_int = int(dia_txt)
        if 1 <= dia_int <= 31:
            dia = dia_int
    return anio, mes, dia


def _colectar_documentos_por_periodo(
    base_dir: Path,
    tipo_prefijo: str,
    extension: str,
    modo_fecha: str,
    anio_objetivo: int,
    mes_inicio: int,
    mes_fin: int,
    dia_objetivo: int,
    tipo_slug_archivo: str | None = None,
) -> list[Path]:
    if not base_dir.exists():
        return []
    ext = (extension or "").strip().lower().lstrip(".")
    if ext not in {"xml", "pdf"}:
        return []
    encontrados: list[Path] = []
    for ruta in base_dir.rglob(f"*.{ext}"):
        try:
            partes_norm = {_normalizar_texto_simple(parte) for parte in ruta.parts}
            if "consolidados" in partes_norm:
                continue
            if ruta.parent.name.lower() != ext:
                continue
            fecha_info = _extraer_fecha_desde_ruta_documento(ruta, tipo_prefijo, tipo_slug_archivo)
            if not fecha_info:
                continue
            anio_arch, mes_arch, dia_arch = fecha_info
            if _fecha_coincide_consolidacion(
                anio_arch,
                mes_arch,
                dia_arch,
                modo_fecha,
                anio_objetivo,
                mes_inicio,
                mes_fin,
                dia_objetivo,
            ):
                encontrados.append(ruta)
        except Exception:
            continue
    return sorted(dict.fromkeys(encontrados), key=lambda p: str(p).lower())


def _copiar_documentos_unicos(archivos: list[Path], destino: Path) -> int:
    if not archivos:
        return 0
    destino.mkdir(parents=True, exist_ok=True)
    copiados = 0
    for origen in archivos:
        if not origen.exists():
            continue
        destino_final = destino / origen.name
        if destino_final.exists():
            sufijo = 1
            while True:
                candidato = destino / f"{origen.stem}_{sufijo}{origen.suffix}"
                if not candidato.exists():
                    destino_final = candidato
                    break
                sufijo += 1
        try:
            shutil.copy2(origen, destino_final)
            copiados += 1
        except Exception as err:
            print(f"[WARN] No se pudo copiar documento consolidado: {origen} ({err})")
    return copiados


def _consolidar_reportes_xml_desde_excels(reportes: list[Path], destino: Path) -> Path | None:
    rutas = [p for p in reportes if isinstance(p, Path) and p.exists()]
    if not rutas:
        return None
    dataframes: list[pd.DataFrame] = []
    columnas: list[str] | None = None
    for ruta in rutas:
        try:
            xls = pd.ExcelFile(ruta)
            hoja = "Cabecera" if "Cabecera" in xls.sheet_names else (xls.sheet_names[0] if xls.sheet_names else None)
            if not hoja:
                continue
            df = pd.read_excel(ruta, sheet_name=hoja)
        except Exception as err:
            print(f"[WARN] No se pudo leer reporte XML para consolidar: {ruta} ({err})")
            continue
        if df is None or df.empty:
            continue
        if columnas is None:
            columnas = list(df.columns)
        else:
            for col in df.columns:
                if col not in columnas:
                    columnas.append(col)
        dataframes.append(df)
    if not dataframes or not columnas:
        return None
    for idx, df in enumerate(dataframes):
        for col in columnas:
            if col not in df.columns:
                df[col] = ""
        dataframes[idx] = df[columnas]
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        combinado = pd.concat(dataframes, ignore_index=True)
        combinado.to_excel(destino, index=False)
        return destino
    except Exception as err:
        print(f"[WARN] No se pudo escribir reporte XML consolidado: {destino} ({err})")
        return None


def _normalizar_ruc(value: str | None) -> str:
    return "".join(ch for ch in (value or "") if ch.isdigit())


def _resolver_busqueda_consolidacion(
    carpeta_base: Path,
    origen: str,
    ruc_hint: str | None = None,
    estado_slug: str | None = None,
) -> Path:
    ruc_limpio = _normalizar_ruc(ruc_hint)
    candidatos: list[Path] = []
    vistos: set[str] = set()

    def _agregar(path: Path):
        clave = str(path.resolve()) if path.exists() else str(path)
        if clave in vistos:
            return
        vistos.add(clave)
        candidatos.append(path)

    if ruc_limpio and origen == "Emitidos" and estado_slug:
        _agregar(carpeta_base / ruc_limpio / origen / estado_slug)
    if origen == "Emitidos" and estado_slug:
        _agregar(carpeta_base / origen / estado_slug)
        _agregar(carpeta_base / estado_slug)
    if ruc_limpio:
        _agregar(carpeta_base / ruc_limpio / origen)
    _agregar(carpeta_base / origen)
    if ruc_limpio:
        _agregar(carpeta_base / ruc_limpio)
    _agregar(carpeta_base)

    for cand in candidatos:
        if not cand.exists():
            continue
        try:
            if next(cand.rglob("*.xlsx"), None):
                return cand
        except Exception:
            pass

    for cand in candidatos:
        if cand.exists():
            return cand

    return carpeta_base


def _init_download_state():
    if "download_status" not in st.session_state:
        st.session_state.download_status = "idle"
    if "download_thread" not in st.session_state:
        st.session_state.download_thread = None
    if "download_queue" not in st.session_state:
        st.session_state.download_queue = queue.Queue()
    if "download_messages" not in st.session_state:
        st.session_state.download_messages = []
    if "last_download_message" not in st.session_state:
        st.session_state.last_download_message = None
    if "download_result" not in st.session_state:
        st.session_state.download_result = None
    if "download_error" not in st.session_state:
        st.session_state.download_error = None
    if "download_params" not in st.session_state:
        st.session_state.download_params = None
    if "download_registered" not in st.session_state:
        st.session_state.download_registered = False
    if "running_notice_ts" not in st.session_state:
        st.session_state.running_notice_ts = None
    if "stop_notice_ts" not in st.session_state:
        st.session_state.stop_notice_ts = None
    if "manual_consultar_hint" not in st.session_state:
        st.session_state.manual_consultar_hint = None
    if "manual_consultar_hint_ts" not in st.session_state:
        st.session_state.manual_consultar_hint_ts = None
    if "download_finished_modal_open" not in st.session_state:
        st.session_state.download_finished_modal_open = False

def _drain_download_queue():
    q = st.session_state.download_queue
    while True:
        try:
            kind, payload = q.get_nowait()
        except queue.Empty:
            break
        if kind == "msg":
            mensaje = str(payload)
            st.session_state.download_messages.append(mensaje)
            st.session_state.last_download_message = (mensaje, time.time())
            mensaje_norm = mensaje.lower()
            if (
                "[accion]" in mensaje_norm
                and "consultar" in mensaje_norm
                and "recibidos" in mensaje_norm
            ):
                st.session_state.manual_consultar_hint = mensaje
                st.session_state.manual_consultar_hint_ts = time.time()
        elif kind == "done":
            st.session_state.download_result = payload
            st.session_state.download_status = "done"
            st.session_state.manual_consultar_hint = None
            st.session_state.manual_consultar_hint_ts = None
            st.session_state.download_finished_modal_open = True
        elif kind == "error":
            st.session_state.download_error = str(payload)
            st.session_state.download_status = "error"
            st.session_state.manual_consultar_hint = None
            st.session_state.manual_consultar_hint_ts = None

def _download_worker(params: dict, q: "queue.Queue"):
    from robot.downloader import descargar_sri, set_user_notifier, clear_cancel

    def _notify(msg: str):
        try:
            q.put(("msg", msg))
        except Exception:
            pass

    clear_cancel()
    set_user_notifier(_notify)
    checkpoint_path = params.get("checkpoint_path")
    try:
        if checkpoint_path:
            mark_download_checkpoint_running(checkpoint_path)
        resultado = descargar_sri(**params)
        if checkpoint_path:
            delete_download_checkpoint(checkpoint_path)
        q.put(("done", resultado))
    except Exception as err:
        if checkpoint_path:
            mark_download_checkpoint_failed(checkpoint_path, str(err))
        q.put(("error", str(err)))
    finally:
        set_user_notifier(None)


def _friendly_download_error_message(raw_error: str, origen: str | None = None) -> tuple[str, str]:
    err = (raw_error or "").strip()
    low = err.lower()
    origen_low = (origen or "").strip().lower()

    if "proceso cancelado por el usuario" in low:
        return "warning", "Proceso cancelado por el usuario."

    if (
        "[navegador]" in low
        or "no se pudo abrir el navegador" in low
        or "target page, context or browser has been closed" in low
        or "launch_persistent_context" in low
    ):
        return (
            "error",
            "No se pudo abrir el navegador. Verifica que Google Chrome esté "
            "instalado y cierra todas las ventanas de Chrome abiertas (incluido "
            "cualquier navegador que el robot haya dejado abierto). Si el "
            "problema persiste, reinicia el equipo e intenta nuevamente.",
        )

    if "no se pudo abrir el formulario de emitidos" in low or ("emitidos" in low and "timeout" in low):
        return (
            "warning",
            "No se pudo abrir el formulario de Emitidos en este momento. El portal del SRI puede estar lento. Vuelve a intentar en unos segundos.",
        )

    if "captcha" in low:
        return (
            "warning",
            "La consulta no pudo validarse por captcha. Vuelve a intentar. Si persiste, espera 1 o 2 minutos y prueba otra vez.",
        )

    if "indisponibilidad temporal" in low or ("portal del sri" in low and "indispon" in low):
        return (
            "warning",
            "El portal del SRI esta temporalmente no disponible. Intenta nuevamente en unos minutos.",
        )

    if "login del sri" in low or "credenciales" in low:
        return (
            "error",
            "No se pudo iniciar sesión en el portal del SRI. Verifica tus credenciales e intenta nuevamente.",
        )

    if "error http" in low:
        return (
            "error",
            "La consulta no pudo completarse por una respuesta inesperada del portal. Intenta nuevamente.",
        )

    if "timeout" in low:
        if origen_low == "emitidos":
            return (
                "warning",
                "La consulta de Emitidos tardo mas de lo esperado. Vuelve a intentar en unos segundos.",
            )
        return "warning", "La consulta tardo mas de lo esperado. Vuelve a intentar."

    return "error", "Ocurrio un inconveniente durante la consulta. Vuelve a intentar."

# ==============================
# CONFIGURACIN GENERAL
# ==============================
st.set_page_config(
    page_title="SRI Robot Audit",
    page_icon=str(Path(__file__).parent / "AUDIT_IA_sin_fondo_transparente_FINAL.png"),
    layout="wide",
    # Sidebar oculto: el mockup usa un topbar sticky en lugar de side menu.
    # El menu de perfil y el chequeo de actualizaciones se muestran ahora
    # en el header superior (ver _render_topbar) y en la pestana Ayuda.
    initial_sidebar_state="collapsed",
)
_auto_update_ui()
_render_update_modal()
# --- Tema de la interfaz (claro / oscuro) -----------------------------------
_THEME_TOKENS = {
    "dark": {
        # Paleta dark del mockup: base #070b14, panels #111a2e, lineas #1f2b45.
        # Acentos verde primario + azul secundario flotan como gradientes
        # radiales sutiles desde las esquinas (igual que el mockup).
        "--bg": ("radial-gradient(60rem 30rem at 90% -10%, rgba(91,140,255,0.14), transparent 60%),"
                 "radial-gradient(50rem 30rem at -10% 10%, rgba(34,197,94,0.12), transparent 60%),"
                 "linear-gradient(180deg, #070b14 0%, #0b1120 100%)"),
        "--text": "#eaf0fb",
        "--text-strong": "#ffffff",
        "--text-muted": "#94a1bd",
        "--text-label": "#cfd8ee",
        # Paneles solidos del mockup (#111a2e y #0e1626). Mantenemos los
        # tokens "glass" del codigo existente para no romper selectores —
        # solo cambia el color que entregan.
        "--glass": "#111a2e",
        "--glass-strong": "#111a2e",
        "--glass-hover": "#172238",
        "--border": "#1f2b45",
        "--border-strong": "#2a3a5e",
        "--input-bg": "#0e1626",
        "--input-bg-hover": "#172238",
        "--input-bg-focus": "rgba(34,197,94,0.08)",
        "--shadow": "0 24px 60px rgba(0,0,0,0.55)",
        "--sidebar-bg": "linear-gradient(180deg, #0b1120 0%, #070b14 100%)",
        # Titulo en gradiente verde→blanco→azul para reflejar la marca.
        "--title-grad": "linear-gradient(120deg, #22c55e 0%, #ffffff 55%, #5b8cff 100%)",
        "--table-head": "#0e1626",
        "--table-row": "transparent",
        "--table-row-alt": "rgba(255,255,255,0.02)",
        "--table-hover": "#172238",
        # Sombra y borde de los group cards. En dark no se necesita
        # elevacion (el borde + bg oscuro ya recortan); en light se usa
        # una sombra suave para diferenciar visualmente cada paso.
        "--card-shadow": "none",
        "--card-border": "#ced1adc0",
    },
    "light": {
        # Paleta light del mockup: bg #eef2f8, panel #ffffff, panel-2 #f4f7fb.
        "--bg": ("radial-gradient(60rem 30rem at 90% -10%, rgba(91,140,255,0.10), transparent 60%),"
                 "radial-gradient(50rem 30rem at -10% 10%, rgba(34,197,94,0.10), transparent 60%),"
                 "linear-gradient(180deg, #eef2f8 0%, #e6ebf3 100%)"),
        "--text": "#101828",
        "--text-strong": "#0d1426",
        "--text-muted": "#516081",
        "--text-label": "#34416a",
        "--glass": "#ffffff",
        "--glass-strong": "#ffffff",
        "--glass-hover": "#f4f7fb",
        "--border": "#d8e0ee",
        "--border-strong": "#c2cce0",
        "--input-bg": "#f4f7fb",
        "--input-bg-hover": "#ffffff",
        "--input-bg-focus": "rgba(34,197,94,0.10)",
        "--shadow": "0 20px 45px rgba(30,42,80,0.10)",
        "--sidebar-bg": "linear-gradient(180deg, #ffffff 0%, #f4f7fb 100%)",
        "--title-grad": "linear-gradient(120deg, #16a34a 0%, #1d2b50 55%, #3a5bb8 100%)",
        "--table-head": "#f4f7fb",
        "--table-row": "#ffffff",
        "--table-row-alt": "#f4f7fb",
        "--table-hover": "#e7edf6",
        # Sombra prominente + borde marcado para que cada step quede
        # claramente recortado del fondo claro — match con mockup.
        # Subimos el peso del borde (1.5px) y opacidades del shadow
        # a 0.22/0.14 para que los recuadros se vean SIN ambiguedad
        # incluso a primer vistazo.
        "--card-shadow": "0 2px 6px rgba(15,23,42,0.22), 0 12px 28px rgba(15,23,42,0.14)",
        "--card-border": "#666666",
    },
}

# Tokens fijos (no dependen del tema): acentos, radios, anillo de foco.
# El mockup usa VERDE como primario (#22c55e) y azul como secundario (#5b8cff).
_THEME_TOKENS_FIJOS = {
    "--accent": "#22c55e",
    "--accent-2": "#5b8cff",
    "--accent-soft": "rgba(34,197,94,0.16)",
    "--accent-tint": "#0f2a1c",
    "--success": "#22c55e",
    "--danger": "#f0556a",
    "--warn": "#e7b24a",
    "--radius": "14px",
    "--radius-lg": "18px",
    "--ring": "0 0 0 3px rgba(34,197,94,0.28)",
}


def _build_global_css(theme: str) -> str:
    """Construye el bloque <style> global para el tema indicado.

    Todo el CSS usa var(--token); solo cambia el bloque :root segun el tema,
    de modo que claro y oscuro comparten EXACTAMENTE el mismo diseno.
    """
    tokens = dict(_THEME_TOKENS_FIJOS)
    tokens.update(_THEME_TOKENS.get(theme, _THEME_TOKENS["dark"]))
    root = ":root{" + "".join(f"{k}:{v};" for k, v in tokens.items()) + "}"
    # Override solo en modo claro: el logo blanco (logo-acg-white.png) no
    # debe llevar mix-blend-mode:screen porque blend "screen" con blanco
    # sobre un fondo claro hace que el logo se funda y desaparezca.
    extra = ""
    if theme == "light":
        extra = "\n.brand img{mix-blend-mode:normal !important;}\n"
    return "<style>\n" + _CSS_IMPORT + "\n" + root + "\n" + _CSS_REGLAS + extra + "\n</style>"


def _render_theme_toggle() -> None:
    """Boton fijo arriba a la derecha para alternar tema claro/oscuro."""
    tema = st.session_state.get("ui_theme", "dark")
    if tema == "dark":
        etiqueta = "☀️  Modo claro"
    else:
        etiqueta = "\U0001f319  Modo oscuro"
    if st.button(etiqueta, key="toggle_theme", help="Cambiar entre tema claro y oscuro"):
        st.session_state["ui_theme"] = "light" if tema == "dark" else "dark"
        st.rerun()


_CSS_IMPORT = "@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');"

_CSS_REGLAS = """
/* ===================== Ocultar UI de Streamlit ===================== */
header[data-testid="stHeader"],
div[data-testid="stToolbar"],
div[data-testid="stToolbarActions"],
div[data-testid="stHeaderActionElements"],
button[title="Deploy"],
button[title*="Deploy"],
button[aria-label="Deploy"],
button[aria-label*="Deploy"],
a[title="Deploy"],
a[title*="Deploy"],
[data-testid="stAppDeployButton"],
[data-testid="baseButton-headerNoPadding"],
div[data-testid="stDecoration"]{
  display:none !important;
}

/* ===================== Fondo y tipografia base ===================== */
.stApp{
  background:var(--bg);
  background-attachment:fixed;
  color:var(--text);
  font-family:'Plus Jakarta Sans', system-ui, sans-serif;
}
.stApp, .stApp p, .stApp span, .stApp label, .stApp li,
.stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6{
  font-family:'Plus Jakarta Sans', system-ui, sans-serif;
}
.stApp div[data-testid="stAppViewContainer"] .main,
.stApp div[data-testid="stAppViewContainer"] .main h1,
.stApp div[data-testid="stAppViewContainer"] .main h2,
.stApp div[data-testid="stAppViewContainer"] .main h3,
.stApp div[data-testid="stAppViewContainer"] .main h4,
.stApp div[data-testid="stAppViewContainer"] .main h5,
.stApp div[data-testid="stAppViewContainer"] .main h6,
.stApp div[data-testid="stAppViewContainer"] .main p,
.stApp div[data-testid="stAppViewContainer"] .main span,
.stApp div[data-testid="stAppViewContainer"] .main label,
.stApp [data-testid="stMarkdownContainer"]{
  color:var(--text);
}

/* ===================== Titulos ===================== */
.app-title{
  font-size:clamp(2.1rem, 2.7vw, 2.8rem) !important;
  font-weight:800 !important;
  line-height:1.14 !important;
  letter-spacing:-0.02em;
  background:var(--title-grad);
  -webkit-background-clip:text;
  background-clip:text;
  -webkit-text-fill-color:transparent;
}
.section-title,
.historial-title{
  font-size:clamp(1.4rem, 2vw, 1.85rem) !important;
  font-weight:700 !important;
  line-height:1.2 !important;
  color:var(--text-strong) !important;
  letter-spacing:-0.01em;
}
/* ===== Pills de estado y autorizacion (historial — estilo mockup) ===== */
.pill{
  display:inline-flex; align-items:center; gap:.35rem;
  padding:.18rem .65rem;
  border-radius:999px;
  font-size:.78rem; font-weight:600;
  border:1px solid transparent;
  line-height:1;
}
/* Estado: ok verde, pendiente amarillo, error rojo, otros gris */
.pill-ok{
  background:color-mix(in srgb, #10b981 18%, transparent);
  color:#10b981; border-color:color-mix(in srgb, #10b981 35%, transparent);
}
.pill-pendiente{
  background:color-mix(in srgb, #f59e0b 18%, transparent);
  color:#f59e0b; border-color:color-mix(in srgb, #f59e0b 35%, transparent);
}
.pill-error{
  background:color-mix(in srgb, #ef4444 18%, transparent);
  color:#ef4444; border-color:color-mix(in srgb, #ef4444 35%, transparent);
}
.pill-other{
  background:color-mix(in srgb, var(--text-muted) 15%, transparent);
  color:var(--text-muted); border-color:var(--border);
}
/* Autorizacion: Autorizados azul, No autorizados naranja */
.pill-auth-yes{
  background:color-mix(in srgb, #3b82f6 18%, transparent);
  color:#3b82f6; border-color:color-mix(in srgb, #3b82f6 35%, transparent);
}
.pill-auth-no{
  background:color-mix(in srgb, #f97316 18%, transparent);
  color:#f97316; border-color:color-mix(in srgb, #f97316 35%, transparent);
}
.pill-auth-none{
  background:color-mix(in srgb, var(--text-muted) 15%, transparent);
  color:var(--text-muted); border-color:var(--border);
}

/* Total de operaciones — barra inferior del historial con badge verde */
.historial-total{
  display:flex; align-items:center; justify-content:space-between;
  margin-top:.8rem; padding:.7rem 1rem;
  background:color-mix(in srgb, var(--glass) 90%, transparent);
  border:1px solid var(--border); border-radius:12px;
}
.historial-total-label{
  font-weight:600; color:var(--text);
}
.historial-total-badge{
  display:inline-flex; align-items:center; gap:.35rem;
  padding:.22rem .7rem;
  border-radius:999px;
  font-size:.82rem; font-weight:700;
  background:color-mix(in srgb, #10b981 20%, transparent);
  color:#10b981;
  border:1px solid color-mix(in srgb, #10b981 40%, transparent);
}

.auth-title{
  text-align:center;
  font-size:1.95rem;
  font-weight:800;
  letter-spacing:-0.01em;
  color:var(--text-strong) !important;
  margin-bottom:1.5rem;
}

/* ===================== Tarjeta glass (formularios) ===================== */
div[data-testid="stForm"]{
  background:var(--glass-strong);
  backdrop-filter:blur(22px) saturate(150%);
  -webkit-backdrop-filter:blur(22px) saturate(150%);
  border:1px solid var(--border-strong);
  border-radius:var(--radius-lg);
  padding:2.4rem 2.6rem;
  box-shadow:var(--shadow), inset 0 1px 0 rgba(255,255,255,0.07);
  position:relative;
}
div[data-testid="stForm"]::before{
  content:"";
  position:absolute;
  inset:0 0 auto 0;
  height:1px;
  border-radius:var(--radius-lg) var(--radius-lg) 0 0;
  background:linear-gradient(90deg, transparent, rgba(34,197,94,0.7), rgba(91,140,255,0.7), transparent);
}
div[data-testid="stForm"] label,
div[data-testid="stForm"] p,
div[data-testid="stForm"] span{
  color:var(--text);
}

/* ===================== Inputs de texto / numero / area ===================== */
/* El fondo se aplica al contenedor BaseWeb, a sus divs internos Y al input
   real: BaseWeb pinta un fondo propio en una capa interna que el contenedor
   exterior no cubre. Antes el input quedaba transparent y en modo claro se
   veia el fondo oscuro nativo (campos negros). */
.stApp div[data-baseweb="input"],
.stApp div[data-baseweb="base-input"],
.stApp div[data-baseweb="textarea"],
.stApp div[data-baseweb="input"] > div,
.stApp div[data-baseweb="base-input"] > div,
.stApp [data-testid="stTextInput"] input,
.stApp [data-testid="stNumberInput"] input,
.stApp [data-testid="stTextArea"] textarea,
.stApp input,
.stApp textarea{
  background:var(--input-bg) !important;
  color:var(--text) !important;
}
.stApp div[data-baseweb="input"],
.stApp div[data-baseweb="base-input"],
.stApp div[data-baseweb="textarea"]{
  border:1px solid var(--border) !important;
  border-radius:var(--radius) !important;
  transition:border-color .18s ease, box-shadow .18s ease;
}
.stApp div[data-baseweb="input"]:hover,
.stApp div[data-baseweb="base-input"]:hover,
.stApp div[data-baseweb="textarea"]:hover{
  border-color:var(--border-strong) !important;
}
.stApp div[data-baseweb="input"]:focus-within,
.stApp div[data-baseweb="base-input"]:focus-within,
.stApp div[data-baseweb="textarea"]:focus-within{
  border-color:var(--accent) !important;
  box-shadow:var(--ring) !important;
}
.stApp input::placeholder,
.stApp textarea::placeholder{
  color:var(--text-muted) !important;
}

/* ===================== Selects ===================== */
.stApp div[data-baseweb="select"] > div{
  background:var(--input-bg) !important;
  border:1px solid var(--border) !important;
  border-radius:var(--radius) !important;
  transition:border-color .18s ease, box-shadow .18s ease;
}
.stApp div[data-baseweb="select"] > div:hover{
  border-color:var(--border-strong) !important;
}
.stApp div[data-baseweb="select"] > div:focus-within{
  border-color:var(--accent) !important;
  box-shadow:var(--ring) !important;
}
.stApp [data-baseweb="select"] div{
  color:var(--text) !important;
}
div[data-testid="stForm"] div[data-baseweb="input"] button{
  background:transparent !important;
  box-shadow:none !important;
}
div[data-testid="stForm"] div[data-baseweb="input"] svg,
.stApp div[data-baseweb="input"] svg{
  color:var(--text) !important;
  fill:var(--text) !important;
}

/* ===================== Inputs deshabilitados (paths, lectura) =====================
   Sin esto el navegador oscurece/desatura el texto y en light queda
   casi invisible. Forzamos color sólido y opacidad alta. */
.stApp [data-testid="stTextInput"] input:disabled,
.stApp [data-testid="stNumberInput"] input:disabled,
.stApp [data-testid="stTextArea"] textarea:disabled,
.stApp input:disabled,
.stApp textarea:disabled{
  color:var(--text) !important;
  -webkit-text-fill-color:var(--text) !important;
  background:var(--input-bg) !important;
  opacity:.92 !important;
}

/* ===================== Stepper buttons del number_input (+/-) =====================
   Streamlit por defecto los pinta oscuros; en modo claro se ven
   como bloques negros. Forzamos var(--input-bg) + texto temado. */
.stApp [data-testid="stNumberInput"] button{
  background:var(--input-bg) !important;
  color:var(--text) !important;
  border:1px solid var(--border) !important;
  box-shadow:none !important;
}
.stApp [data-testid="stNumberInput"] button:hover{
  background:var(--input-bg-hover) !important;
  border-color:var(--accent) !important;
}
.stApp [data-testid="stNumberInput"] button svg,
.stApp [data-testid="stNumberInput"] button [data-testid="stIcon"]{
  color:var(--text) !important;
  fill:var(--text) !important;
}

/* ===================== Inline `<code>` (markdown backticks) =====================
   `st.caption(f"... `{path}`")` genera <code> que Streamlit pinta de
   negro por defecto. En light se ven barras negras ilegibles.
   Los tematizamos para que respeten claro/oscuro. */
.stApp [data-testid="stMarkdownContainer"] code,
.stApp [data-testid="stCaptionContainer"] code,
.stApp [data-testid="stCaption"] code,
.stApp p code,
.stApp small code,
.stApp code{
  background:var(--input-bg) !important;
  color:var(--text) !important;
  border:1px solid var(--border) !important;
  padding:.08rem .38rem !important;
  border-radius:6px !important;
  font-size:.86em !important;
  font-family:'JetBrains Mono', ui-monospace, SFMono-Regular, Menlo, monospace !important;
}

/* ===================== Boton primario (submit de formularios) ===================== */
div[data-testid="stForm"] button[kind="primaryFormSubmit"],
div[data-testid="stForm"] button[data-testid="baseButton-primary"]{
  background:linear-gradient(135deg, var(--accent) 0%, var(--accent-2) 100%) !important;
  color:#ffffff !important;
  border:none !important;
  border-radius:12px !important;
  font-weight:700 !important;
  letter-spacing:0.01em;
  padding:0.6rem 1.2rem !important;
  box-shadow:0 12px 28px rgba(34,197,94,0.32) !important;
  transition:transform .15s ease, box-shadow .15s ease, filter .15s ease;
}
div[data-testid="stForm"] button[kind="primaryFormSubmit"]:hover,
div[data-testid="stForm"] button[data-testid="baseButton-primary"]:hover{
  transform:translateY(-2px);
  filter:brightness(1.08);
  box-shadow:0 16px 36px rgba(34,197,94,0.45) !important;
}
div[data-testid="stForm"] button[kind="primaryFormSubmit"]:disabled,
div[data-testid="stForm"] button[data-testid="baseButton-primary"]:disabled{
  background:var(--glass-hover) !important;
  color:var(--text-muted) !important;
  box-shadow:none !important;
  transform:none;
}

/* ===================== Botones generales ===================== */
.stApp .stButton > button{
  border-radius:12px !important;
  border:1px solid var(--border-strong) !important;
  background:var(--glass-strong) !important;
  color:var(--text) !important;
  font-weight:600 !important;
  transition:transform .15s ease, background .15s ease, border-color .15s ease, box-shadow .15s ease;
}
.stApp .stButton > button:hover{
  background:var(--glass-hover) !important;
  border-color:var(--accent) !important;
  transform:translateY(-1px);
}
/* Botones Iniciar/Detener proceso — los aria-label ahora incluyen
   los emojis ▶️ / ⏹️ porque agregamos iconos al label visible.
   Mantenemos los selectores antiguos por compatibilidad si alguien
   revierte el cambio del label. */
button[aria-label*="Iniciar proceso"]{
  background:linear-gradient(135deg, #16a34a 0%, #22c55e 100%) !important;
  border:none !important;
  color:#ffffff !important;
  font-weight:700 !important;
  box-shadow:0 12px 28px rgba(34,197,94,0.32) !important;
}
button[aria-label*="Iniciar proceso"]:hover{
  transform:translateY(-2px);
  filter:brightness(1.08);
  box-shadow:0 16px 36px rgba(34,197,94,0.45) !important;
}
button[aria-label*="Detener proceso"]{
  background:linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
  border:none !important;
  color:#ffffff !important;
  font-weight:700 !important;
  box-shadow:0 12px 28px rgba(239,68,68,0.30) !important;
}
button[aria-label*="Detener proceso"]:hover{
  transform:translateY(-2px);
  filter:brightness(1.08);
  box-shadow:0 16px 36px rgba(239,68,68,0.42) !important;
}

/* ===================== Boton toggle de tema arriba a la derecha ===================== */
.st-key-toggle_theme{
  position:static !important;
  width:0 !important;
  max-width:0 !important;
  height:0 !important;
  min-width:0 !important;
  margin:0 !important;
  padding:0 !important;
  overflow:visible !important;
}
.st-key-toggle_theme button{
  position:fixed !important;
  top:14px !important;
  right:18px !important;
  left:auto !important;
  bottom:auto !important;
  width:auto !important;
  max-width:max-content !important;
  min-height:0 !important;
  padding:0.4rem 1rem !important;
  border-radius:999px !important;
  background:var(--glass-strong) !important;
  border:1px solid var(--border-strong) !important;
  color:var(--text) !important;
  font-size:0.86rem !important;
  font-weight:600 !important;
  backdrop-filter:blur(14px) saturate(150%);
  -webkit-backdrop-filter:blur(14px) saturate(150%);
  box-shadow:0 8px 22px rgba(0,0,0,0.22) !important;
  z-index:9999 !important;
}
.st-key-toggle_theme button:hover{
  border-color:var(--accent) !important;
  background:var(--glass-hover) !important;
  transform:translateY(-1px);
}

/* ===================== Enlace de recuperar contrasena ===================== */
.auth-reset-wrap{
  display:flex;
  justify-content:flex-end;
  margin-top:12px;
}
.auth-reset-link{
  display:inline-flex;
  align-items:center;
  justify-content:center;
  padding:9px 15px;
  border-radius:10px;
  border:1px solid var(--border-strong);
  background:var(--glass-strong);
  color:var(--accent) !important;
  font-size:0.9rem;
  font-weight:600;
  text-decoration:none !important;
  transition:background .15s ease, border-color .15s ease, color .15s ease;
}
.auth-reset-link:hover{
  background:var(--accent-soft);
  border-color:var(--accent);
  color:var(--accent) !important;
  text-decoration:none !important;
}

/* ===================== Tabs ===================== */
.stApp [data-baseweb="tab-list"]{
  gap:6px;
  background:var(--glass);
  border:1px solid var(--border);
  border-radius:14px;
  padding:6px;
}
.stApp [data-baseweb="tab"]{
  border-radius:10px !important;
  color:var(--text-muted) !important;
  font-weight:600 !important;
  padding:0.5rem 1rem !important;
  transition:background .15s ease, color .15s ease;
}
.stApp [data-baseweb="tab"]:hover{
  background:var(--glass-hover) !important;
  color:var(--text) !important;
}
.stApp [data-baseweb="tab"][aria-selected="true"]{
  background:linear-gradient(135deg, var(--accent-soft), rgba(91,140,255,0.16)) !important;
  color:var(--text-strong) !important;
}
.stApp [data-baseweb="tab-highlight"],
.stApp [data-baseweb="tab-border"]{
  display:none !important;
}

/* ===================== Sidebar ===================== */
section[data-testid="stSidebar"]{
  background:var(--sidebar-bg) !important;
  border-right:1px solid var(--border);
  backdrop-filter:blur(12px);
}
section[data-testid="stSidebar"] *{
  color:var(--text) !important;
}
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
section[data-testid="stSidebar"] p{
  text-align:center;
}
section[data-testid="stSidebar"] [data-testid="stImage"]{
  display:flex;
  justify-content:center;
}
section[data-testid="stSidebar"] [data-testid="stImage"] img,
section[data-testid="stSidebar"] img{
  margin-left:auto !important;
  margin-right:auto !important;
}
/* Linea de version en sidebar */
.sidebar-version-line{
  text-align:center !important;
  font-size:0.95rem !important;
  color:var(--text) !important;
  margin:0.4rem 0 0.55rem 0;
}
.sidebar-version-line strong{
  color:var(--text-strong) !important;
  font-weight:700 !important;
}
/* Boton "Buscar actualizaciones" en sidebar */
.st-key-btn_buscar_update button{
  background:linear-gradient(135deg, var(--accent) 0%, var(--accent-2) 100%) !important;
  color:#ffffff !important;
  border:none !important;
  border-radius:12px !important;
  font-weight:700 !important;
  letter-spacing:0.01em;
  padding:0.55rem 1rem !important;
  box-shadow:0 10px 24px rgba(34,197,94,0.32) !important;
  transition:transform .15s ease, filter .15s ease, box-shadow .15s ease;
}
.st-key-btn_buscar_update button *{
  color:#ffffff !important;
}
.st-key-btn_buscar_update button:hover{
  transform:translateY(-1px);
  filter:brightness(1.08);
  box-shadow:0 14px 32px rgba(34,197,94,0.45) !important;
}

/* ===================== Captions visibles en dark mode ===================== */
.stApp [data-testid="stCaptionContainer"],
.stApp [data-testid="stCaptionContainer"] *,
.stApp [data-testid="stCaption"],
.stApp [data-testid="stCaption"] *,
.stApp [data-testid="stMarkdownContainer"] small,
.stApp [data-testid="stMarkdownContainer"] small *,
.stApp [data-testid="stMarkdown"] small,
.stApp [data-testid="stMarkdown"] small *{
  color:var(--text) !important;
  opacity:1 !important;
}

/* ===================== Labels de widgets ===================== */
.stApp [data-testid="stWidgetLabel"],
.stApp [data-testid="stWidgetLabel"] p,
.stApp label[data-testid="stWidgetLabel"],
.stApp label[data-testid="stWidgetLabel"] p,
.stApp div[data-testid="stWidgetLabel"] label,
.stApp div[data-testid="stWidgetLabel"] label p{
  font-size:1.0rem !important;
  font-weight:600 !important;
  letter-spacing:0.01em;
  color:var(--text-label) !important;
}
.stApp input,
.stApp textarea,
.stApp [data-baseweb="select"] div,
.stApp [data-baseweb="radio"] span,
.stApp [data-baseweb="checkbox"] span{
  font-size:1.0rem !important;
}

/* ===================== Divisores ===================== */
.stApp hr{
  border:none;
  height:1px;
  background:linear-gradient(90deg, transparent, var(--border-strong), transparent);
}

/* ===================== Alertas ===================== */
.stApp div[data-testid="stAlert"]{
  border-radius:12px;
  border:1px solid var(--border-strong);
  backdrop-filter:blur(8px);
}

/* ===================== Boton de tour ===================== */
.st-key-btn_open_tour{
  display:flex;
  justify-content:flex-start;
  align-items:center;
  padding-top:0.3rem;
}
.st-key-btn_open_tour button{
  width:auto !important;
  min-height:2rem !important;
  padding:0.35rem 0.9rem !important;
  border-radius:999px !important;
  border:1px solid var(--border-strong) !important;
  background:var(--glass-strong) !important;
  color:var(--text-label) !important;
  font-size:0.88rem !important;
  font-weight:600 !important;
  box-shadow:none !important;
}
.st-key-btn_open_tour button::before{
  content:"";
  display:inline-block;
  width:0.46rem;
  height:0.46rem;
  margin-right:0.5rem;
  border-radius:999px;
  background:var(--accent);
  box-shadow:0 0 0 3px rgba(34,197,94,0.22);
  vertical-align:middle;
}
.st-key-btn_open_tour button:hover{
  color:var(--text-strong) !important;
  border-color:var(--accent) !important;
  background:var(--accent-soft) !important;
}

/* ============================================================ */
/* ============== COMPONENTES MIGRADOS DEL MOCKUP ============== */
/* ============================================================ */

/* App ocupa TODO el ancho del viewport — no la confinamos a 1080px (eso
   metia el topbar y el login dentro de un "box" centrado). Mantenemos
   solo un padding minimo a los costados para que el contenido no se
   pegue al borde. */
.stApp div[data-testid="stMainBlockContainer"],
.stApp section.main > div.block-container{
  max-width:none !important;
  padding-top:0 !important;
  padding-left:0 !important;
  padding-right:0 !important;
  padding-bottom:0 !important;
}

/* === Topbar (sticky, full-bleed arriba) ===
   Spaneo el topbar al borde del viewport: sin border-radius, sin
   border lateral, solo una linea inferior. Margenes negativos para
   "salir" del padding interno que Streamlit pone alrededor del main. */
.app-topbar{
  position:sticky; top:0; z-index:30;
  display:flex; align-items:center; justify-content:space-between;
  gap:1rem; padding:.85rem clamp(1.2rem,3vw,2rem);
  background:color-mix(in srgb, var(--glass) 86%, transparent);
  backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px);
  border:0; border-bottom:1px solid var(--border);
  border-radius:0;
  margin:0 0 1.4rem 0;
  width:100%;
}

/* Contenido de los tabs: lo centramos a un ancho amplio (1400px) para
   que las cards/forms respiren horizontalmente y se vean alargadas
   como en el mockup, sin afectar al topbar (que sigue al 100%).
   IMPORTANTE: width:100% es necesario porque Streamlit por defecto
   colapsa el ancho del contenedor de tabs al ancho del contenido mas
   grande. Sin esto, tabs con contenido chico (como Descarga con sus
   group cards) se ven mas angostas que tabs con contenido ancho
   (como Reportes con su tabla). */
.stApp [data-testid="stTabs"]{
  width:100%;
  max-width:1400px;
  margin:0 auto;
  padding:0 clamp(1rem,2.5vw,2.5rem);
}
/* Forzar el panel de contenido del tab activo a ocupar todo el ancho
   disponible. Sin esto, st.columns + group cards dentro de un tab
   pueden hacer que el panel se "encoja" al ancho minimo necesario. */
.stApp [data-testid="stTabs"] [data-baseweb="tab-panel"]{
  width:100%;
}
/* Forzar group cards (st.container border) a ocupar 100% del ancho
   del tab. Sin esto, el border wrapper de Streamlit a veces se
   ajusta al ancho del contenido. */
.stApp [data-testid="stVerticalBlockBorderWrapper"]{
  width:100% !important;
}
/* ===== Topbar FIXED al viewport (no sticky) =====
   Usamos position:fixed para garantizar que el topbar quede PEGADO
   al borde superior del viewport, ignorando cualquier padding
   heredado de los wrappers de Streamlit (que en versiones nuevas
   meten padding-top/left que cuesta sobreescribir con sticky).
   El spacer (`.topbar-spacer`) renderizado justo despues empuja
   el contenido para que no quede tapado. */
.stApp .st-key-topbar_container{
  position:fixed !important;
  top:0 !important;
  left:0 !important;
  right:0 !important;
  z-index:50 !important;
  height:85px !important;
  background:color-mix(in srgb, var(--glass) 94%, transparent) !important;
  backdrop-filter:blur(14px);
  -webkit-backdrop-filter:blur(14px);
  border:0 !important;
  border-bottom:1px solid var(--border) !important;
  border-radius:0 !important;
  padding:0 clamp(1rem, 2.5vw, 1.8rem) !important;
  margin:0 !important;
  width:100% !important;
  max-width:none !important;
  box-shadow:none !important;
  /* NO usar display:flex aqui — eso colapsaba las columnas internas
     de Streamlit. El stVerticalBlock interno YA es display:flex
     column. Solo necesitamos que ocupe 100% width/height del topbar. */
  overflow:hidden;
}
/* El stVerticalBlock interno toma el ancho/alto completo del topbar.
   Su display:flex column ya lo mete Streamlit por defecto. */
.stApp .st-key-topbar_container > [data-testid="stVerticalBlock"]{
  width:100% !important;
  height:100% !important;
  flex-direction:column !important;
  justify-content:center !important;
}
/* La fila de columnas (stHorizontalBlock) debe ocupar el 100% del
   ancho disponible, mantener align-items:center y position:relative
   para que el titulo absoluto se centre matematicamente respecto a
   ESTA fila (que abarca todo el viewport). */
.stApp .st-key-topbar_container [data-testid="stHorizontalBlock"]{
  width:100% !important;
  align-items:center !important;
  position:relative;
}
/* Cada columna individual debe respetar el alineamiento center sin
   colapsar. */
.stApp .st-key-topbar_container [data-testid="column"]{
  display:flex !important;
  align-items:center !important;
}
/* Spacer renderizado debajo del topbar para que el contenido
   real (tabs, formularios) no quede tapado por el topbar fijo. */
.topbar-spacer{
  height:85px !important;
  width:100%;
  display:block;
  margin:0 0 1rem 0;
}

/* FIX doble linea: el `st.markdown('<div class=is-topbar-marker>')`
   crea un element-container que ocupa una fila vertical adentro del
   topbar. Escondemos TODO el element-container del marker (no solo
   el div interno) para que no genere espaciado fantasma. */
.stApp .st-key-topbar_container [data-testid="element-container"]:has(.is-topbar-marker){
  display:none !important;
}
.is-topbar-marker{display:none}

/* Quitar el `gap: 1rem` por defecto de Streamlit dentro del topbar.
   Sin esto, queda una banda vacia entre el marker oculto y la fila
   de columnas que se ve como "segunda barra". */
.stApp .st-key-topbar_container > [data-testid="stVerticalBlock"]{
  gap:0 !important;
}
.stApp .st-key-topbar_container > [data-testid="stVerticalBlock"] > [data-testid="element-container"]{
  margin:0 !important;
}
/* El stHorizontalBlock (la fila de columnas) tambien debe estar a 0 */
.stApp .st-key-topbar_container [data-testid="stHorizontalBlock"]{
  margin:0 !important;
  gap:.7rem !important;
}

/* === Layout 3-zonas con titulo CENTRADO al viewport ===
   En vez de depender de pesos de columnas (asimetrico por brand
   ancho + 3 controles a la derecha), posicionamos el titulo en
   absolute relative al topbar para garantizar centrado real al
   ancho del viewport. */
.stApp .st-key-topbar_container{position:sticky; /* contexto para absolute */}
.stApp .st-key-topbar_container [data-testid="stHorizontalBlock"]{
  position:relative;
}
.topbar-title{
  position:absolute;
  left:50%;
  top:50%;
  transform:translate(-50%, -50%);
  font-size:.95rem; color:var(--text-muted); font-weight:600;
  text-align:center; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
  pointer-events:none;
  max-width:50vw;
}
.topbar-title b{color:var(--text)}

/* Brand (solo logo) */
.brand{display:flex; align-items:center; justify-content:center; gap:0; white-space:nowrap; margin:0 !important; padding:0 !important; line-height:1; width:100%; height:auto}
.brand img{height: clamp(35px, 10vw, 90px); width:auto; mix-blend-mode:screen; flex-shrink:0; margin:0 !important; padding:0; margin-top:-25px !important}
.brand .divider{display:none}
.brand span{display:none}

/* Asegurar que el elemento-container del brand también esté centrado */
.stApp .st-key-topbar_container [data-testid="column"] [data-testid="element-container"]:has(.brand){
  display:flex !important;
  align-items:center !important;
  justify-content:center !important;
  width:100% !important;
  height:100% !important;
  margin:0 !important;
  padding:0 !important;
}

/* Chip de version */
.ver-chip{
  display:inline-flex; align-items:center; gap:.35rem;
  font-size:.72rem; font-weight:600; color:var(--text-muted);
  border:1px solid var(--border); border-radius:999px;
  padding:.25rem .65rem; background:var(--input-bg);
}
.ver-chip::before{content:""; width:6px; height:6px; border-radius:50%; background:var(--accent)}

/* Theme toggle button DENTRO del topbar — estilo pill compacto */
.stApp .st-key-btn_topbar_theme button{
  background:var(--input-bg) !important;
  border:1px solid var(--border) !important;
  border-radius:999px !important;
  padding:.35rem .7rem !important;
  font-size:.8rem !important;
  font-weight:600 !important;
  color:var(--text) !important;
  min-height:auto !important;
  height:auto !important;
  box-shadow:none !important;
  white-space:nowrap;
}
.stApp .st-key-btn_topbar_theme button:hover{
  background:var(--input-bg-hover) !important;
  border-color:var(--accent) !important;
}

/* Hamburguesa (popover trigger ☰) — boton compacto cuadrado en el
   topbar. Reemplaza al pill de avatar+email que era permanente. */
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > button{
  background:var(--input-bg) !important;
  border:1px solid var(--border) !important;
  border-radius:10px !important;
  padding:.35rem .55rem !important;
  font-size:1.05rem !important;
  font-weight:700 !important;
  color:var(--text) !important;
  min-height:auto !important;
  height:auto !important;
  width:auto !important;
  min-width:38px !important;
  line-height:1 !important;
  display:inline-flex !important;
  align-items:center !important;
  justify-content:center !important;
  box-shadow:none !important;
}
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > button:hover{
  border-color:var(--accent) !important;
  background:var(--input-bg-hover) !important;
}

/* === Contenido del menu hamburguesa ===
   Adentro del popover: avatar pequeno + sesion activa + email + divider
   + boton Cerrar sesion. */
.user-menu-row{
  display:flex; align-items:center; gap:.7rem;
  padding:.55rem .35rem;
}
.user-menu-avatar{
  width:36px; height:36px; flex:0 0 36px;
  border-radius:50%;
  display:grid; place-items:center;
  font-size:.78rem; font-weight:800; color:#04130b;
  background:linear-gradient(135deg, var(--accent), var(--accent-2));
}
.user-menu-info{display:flex; flex-direction:column; gap:.1rem; min-width:0; flex:1}
.user-menu-label{
  font-size:.7rem; font-weight:700;
  text-transform:uppercase; letter-spacing:.05em;
  color:var(--text-muted);
}
.user-menu-email{
  font-size:.88rem; font-weight:600;
  color:var(--text-strong);
  word-break:break-all;
  overflow:hidden; text-overflow:ellipsis;
}
.user-menu-divider{
  border:0; border-top:1px solid var(--border);
  margin:.4rem 0;
}

/* Gap pequeno entre items del popover */
[data-testid="stPopover"] [data-testid="stVerticalBlock"]{
  gap:.4rem !important;
}

/* ===== Panel flotante del popover (dropdown del menu de usuario) =====
   El panel se monta via portal FUERA de la cadena `.stApp`, asi que
   los selectores prefijados con `.stApp` no lo alcanzan. Usamos los
   testids/atributos de Streamlit y BaseWeb directos para tematizarlo.
   Sin esto el panel toma un fondo oscuro por default y en modo claro
   el email queda casi invisible (texto dark sobre bg dark). */
[data-testid="stPopoverBody"],
[data-baseweb="popover"] [data-baseweb="popover-inner"],
[data-baseweb="popover"] > div > div{
  background:var(--glass) !important;
  color:var(--text) !important;
  border:1px solid var(--border) !important;
  border-radius:12px !important;
  box-shadow:var(--shadow) !important;
}
[data-testid="stPopoverBody"] [data-testid="stMarkdownContainer"],
[data-baseweb="popover"] [data-testid="stMarkdownContainer"]{
  color:var(--text) !important;
}
/* Boton "Cerrar sesion" dentro del popover — tematizar para que en
   modo claro tenga bg claro y texto legible. */
.st-key-btn_popover_logout button,
.st-key-btn_popover_close_app button{
  background:var(--input-bg) !important;
  color:var(--text) !important;
  border:1px solid var(--border) !important;
  border-radius:10px !important;
  font-weight:600 !important;
  box-shadow:none !important;
}
.st-key-btn_popover_logout button:hover,
.st-key-btn_popover_close_app button:hover{
  background:var(--input-bg-hover) !important;
  border-color:var(--accent) !important;
  color:var(--text-strong) !important;
}

@media (max-width:760px){
  .topbar-title{display:none}
  .ver-chip{display:none}
}

/* === Pegar el header al borde superior + remover padding lateral ===
   Reseteamos padding/margin en TODA la cadena de wrappers Streamlit
   con mayor especificidad (body .stApp ...) para ganarle al CSS
   default de Streamlit. Sin esto queda gap arriba y a los lados. */
html, body{
  margin:0 !important;
  padding:0 !important;
}
body .stApp{
  padding:0 !important;
  margin:0 !important;
}
body .stApp [data-testid="stAppViewContainer"]{
  padding:0 !important;
  margin:0 !important;
}
body .stApp section.main,
body .stApp section[data-testid="stMain"],
body .stApp section.main > div.block-container,
body .stApp [data-testid="stMainBlockContainer"]{
  padding-top:0 !important;
  padding-left:0 !important;
  padding-right:0 !important;
  margin:0 !important;
  max-width:none !important;
}

/* Ocultar el toggle del sidebar collapsed (boton flecha arriba-izq)
   para que no se solape con el topbar fixed. Como el sidebar de la
   app esta vacio (todo se movio al topbar), no hay razon para
   exponer el toggle al usuario. */
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"],
button[kind="header"][data-testid="baseButton-header"]{
  display:none !important;
}

/* Compensar el contenido lateral con un padding pequeno DENTRO del
   spacer (no del topbar) para que tabs y forms no toquen el borde. */
body .stApp section.main > div.block-container{
  padding-left:clamp(1rem, 2.5vw, 2.5rem) !important;
  padding-right:clamp(1rem, 2.5vw, 2.5rem) !important;
}
body .stApp .st-key-topbar_container{
  /* El topbar es fixed → vive afuera del block-container, por eso
     necesita su propio padding (ya definido arriba como 0 ... 1.8rem). */
}

/* === Controles derechos uniformes: h-11 (44px), text-sm, rounded-xl ===
   Los 3 controles (chip version + boton tema + icono usuario) comparten
   altura, border-radius, font-size y padding base para verse coherentes
   como una sola "fila de pills". */
.stApp .st-key-topbar_container .ver-chip,
.stApp .st-key-btn_topbar_theme button,
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > button{
  height:44px !important;
  min-height:44px !important;
  display:inline-flex !important;
  align-items:center !important;
  justify-content:center !important;
  font-size:.875rem !important;
  font-weight:600 !important;
  border:1px solid var(--border) !important;
  background:var(--input-bg) !important;
  box-shadow:none !important;
  line-height:1 !important;
  border-radius:12px !important;
  transition:border-color .15s ease, background .15s ease;
}
/* Chip version: padding pill-style + dot verde antes */
.stApp .st-key-topbar_container .ver-chip{
  padding:0 1rem !important;
  color:var(--text-muted) !important;
  gap:.45rem;
}
/* Boton tema: cuadrado 44x44 icon-only (sol/luna), mismo tamano que
   el popover de usuario a la derecha. Sin texto, sin padding lateral. */
.stApp .st-key-btn_topbar_theme button{
  padding:0 !important;
  width:44px !important;
  min-width:44px !important;
  font-size:1.15rem !important;
  color:var(--text) !important;
}
.stApp .st-key-btn_topbar_theme button:hover{
  border-color:var(--accent) !important;
  background:var(--input-bg-hover) !important;
}
/* Boton Perfil: pill con icono 👤 + texto "Perfil". Antes era
   cuadrado icon-only de 44x44; ahora se expande al texto pero
   mantiene la altura 44px para alinear con el toggle de tema.
   Selector permisivo para cubrir cualquier nesting que Streamlit
   aplique al popover (con/sin un div extra). */
.stApp .st-key-topbar_container [data-testid="stPopover"] button,
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > button,
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > div > button{
  padding:0 .9rem !important;
  height:44px !important;
  min-height:44px !important;
  font-size:.875rem !important;
  font-weight:600 !important;
  color:var(--text) !important;
  background:var(--input-bg) !important;
  border:1px solid var(--border) !important;
  border-radius:12px !important;
  box-shadow:none !important;
  white-space:nowrap !important;
  gap:.4rem !important;
}
.stApp .st-key-topbar_container [data-testid="stPopover"] button:hover,
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > button:hover,
.stApp .st-key-topbar_container [data-testid="stPopover"] > div > div > button:hover{
  border-color:var(--accent) !important;
  background:var(--input-bg-hover) !important;
}
/* Por si Streamlit aplica un span/div interno con color hardcoded */
.stApp .st-key-topbar_container [data-testid="stPopover"] button *{
  color:var(--text) !important;
}
/* Gap entre los 3 controles unificado */
.stApp .st-key-topbar_container [data-testid="stHorizontalBlock"]{
  gap:12px !important;
}

/* === pagehead === */
.pagehead{margin:0 0 1.1rem 0}
.pagehead h1{
  font-size:clamp(1.4rem, 2.6vw, 1.95rem) !important;
  font-weight:800 !important; letter-spacing:-.02em !important;
  line-height:1.1 !important; margin:0 !important;
  background:var(--title-grad); -webkit-background-clip:text; background-clip:text;
  -webkit-text-fill-color:transparent;
}
.pagehead p{color:var(--text-muted); margin:.35rem 0 0 !important; font-size:.92rem}

/* === Group cards (steps numerados) ===
   Aplicamos look "group" al wrapper que Streamlit renderiza para los
   containers que llevan nuestro header `.group-h-marker` (truco con :has).
*/
/* Streamlit emite `class="st-key-group_card_N"` en el wrapper de
   stVerticalBlockBorderWrapper cuando pasamos `key="group_card_N"`
   a st.container(border=True). Targeteamos esa clase directamente
   en vez de :has() — esto es robusto a la version del navegador
   y al nesting que Streamlit aplique internamente. */
.stApp [class*="st-key-group_card_"]{
  background:var(--glass) !important;
  border:1.5px solid var(--card-border, var(--border)) !important;
  border-radius:var(--radius) !important;
  padding:1.1rem 1.2rem !important;
  margin-bottom:1rem !important;
  box-shadow:var(--card-shadow, none) !important;
}
/* Cualquier OTRO `stVerticalBlockBorderWrapper` (que NO sea card
   numerado — sin la clase st-key-group_card_) queda neutro:
   sin borde, sin sombra, sin padding extra, ni background. */
.stApp [data-testid="stVerticalBlockBorderWrapper"]:not([class*="st-key-group_card_"]){
  border:none !important;
  box-shadow:none !important;
  padding:0 !important;
  background:transparent !important;
}
.group-h-marker{
  display:flex; align-items:center; gap:.7rem;
  margin:0 0 .9rem 0;
}
.group-h-marker .step{
  width:26px; height:26px; flex:0 0 26px;
  border-radius:8px; display:grid; place-items:center;
  font-weight:800; font-size:.8rem;
  color:var(--accent); background:var(--accent-soft);
  border:1px solid color-mix(in srgb, var(--accent) 35%, transparent);
}
.group-h-marker h3{
  font-size:1.05rem !important; font-weight:700 !important;
  margin:0 !important; color:var(--text-strong) !important;
  letter-spacing:-.01em;
}
.group-h-marker small{
  color:var(--text-muted); font-weight:500;
  margin-left:auto; font-size:.78rem;
}

/* === Quickstart (Ayuda) === */
.quickstart{
  display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));
  gap:1rem; margin:.4rem 0 1.3rem 0;
}
.qs{
  background:var(--glass); border:1px solid var(--border);
  border-radius:var(--radius); padding:1.1rem;
}
.qs .n{
  color:var(--accent); font-family:ui-monospace,"SF Mono",Menlo,Consolas,monospace;
  font-weight:700; font-size:.85rem;
}
.qs h4{margin:.4rem 0 .3rem; font-size:1rem; color:var(--text-strong)}
.qs p{color:var(--text-muted); font-size:.85rem; margin:0}

/* === Accordion === */
.acc{background:var(--glass); border:1px solid var(--border);
  border-radius:var(--radius); overflow:hidden}
.acc details{border-bottom:1px solid var(--border)}
.acc details:last-child{border-bottom:0}
.acc summary{
  list-style:none; cursor:pointer;
  display:flex; align-items:center; gap:.8rem;
  padding:1rem 1.2rem; font-weight:600; color:var(--text);
}
.acc summary::-webkit-details-marker{display:none}
.acc summary .ic{color:var(--accent)}
.acc summary .pm{margin-left:auto; color:var(--text-muted); transition:.2s}
.acc details[open] .pm{transform:rotate(45deg)}
.acc .body{padding:0 1.2rem 1.2rem 3.4rem; color:var(--text-muted); font-size:.9rem}

/* === Chips === */
.chip{
  display:inline-flex; align-items:center; gap:.35rem;
  font-size:.74rem; font-weight:700;
  padding:.2rem .55rem; border-radius:999px;
  border:1px solid transparent; white-space:nowrap;
}
.chip::before{content:""; width:6px; height:6px; border-radius:50%; background:currentColor}
.chip-ok{color:var(--accent); background:var(--accent-soft);
  border-color:color-mix(in srgb, var(--accent) 30%, transparent)}
.chip-auth{color:var(--accent-2); background:color-mix(in srgb, var(--accent-2) 14%, transparent);
  border-color:color-mix(in srgb, var(--accent-2) 30%, transparent)}
.chip-warn{color:var(--warn); background:color-mix(in srgb, var(--warn) 14%, transparent)}

/* === About card === */
.about-card{
  background:var(--glass); border:1px solid var(--border);
  border-radius:var(--radius); padding:1.2rem;
}
.about-card .row{display:flex; align-items:center; gap:.7rem; flex-wrap:wrap}
.about-card .label{color:var(--text-muted); font-size:.82rem; font-weight:600}
.about-card .value{
  font-family:ui-monospace,"SF Mono",Menlo,Consolas,monospace;
  font-weight:700; color:var(--text-strong);
  background:var(--input-bg); border:1px solid var(--border);
  padding:.25rem .65rem; border-radius:999px; font-size:.82rem;
}
.about-card h4{margin:0 0 .8rem; color:var(--text-strong); font-size:1.05rem}

/* === Login con hero lateral, FULL BLEED ===
   El hero ocupa toda la altura de la pantalla y media de ancho — sin
   box ni rounded corners, replicando la pantalla `#login` del mockup. */
.login-shell{display:flex; gap:0; min-height:100vh;
  background:transparent; border-radius:0; overflow:hidden}
.login-visual{position:relative; flex:1.1;
  min-height:calc(100vh - 20px);
  border-radius:0; overflow:hidden; border:0;
  margin:0;
}
.login-visual img{position:absolute; inset:0; width:100%; height:100%;
  object-fit:cover; object-position:34% 46%; display:block}

/* Cuando se renderiza el login, el wrapping de st.columns mete un padding
   que rompe el "full-bleed". Lo neutralizamos. */
.stApp [data-testid="stHorizontalBlock"]:has(.login-visual){
  gap:0 !important;
}
.stApp [data-testid="stHorizontalBlock"]:has(.login-visual) > div{
  padding:0 !important;
}
/* La columna del form recupera un padding interno para que el card no
   toque el borde derecho de la pantalla. */
.stApp [data-testid="stHorizontalBlock"]:has(.login-visual) > div:last-child{
  padding:clamp(2rem,5vw,4rem) clamp(1.5rem,4vw,3rem) !important;
  display:flex; align-items:center;
}
.login-visual .lv-fade{position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(to right, transparent 30%,
    color-mix(in srgb, var(--bg) 60%, transparent) 70%,
    color-mix(in srgb, var(--bg) 90%, transparent) 100%)}
.login-visual .lv-brand{
  position:absolute; left:1.4rem; bottom:1.4rem;
  display:flex; align-items:center; gap:.6rem; z-index:2;
}
.login-visual .lv-brand .b-badge{
  width:36px; height:36px; border-radius:10px;
  display:grid; place-items:center;
  border:1.5px solid var(--accent); background:rgba(7,13,8,.6);
}
.login-visual .lv-brand b{display:block; font-size:.95rem; font-weight:800; color:#fff; line-height:1.1}
.login-visual .lv-brand span.kicker{
  display:block; font-size:.66rem; letter-spacing:.14em;
  text-transform:uppercase; color:color-mix(in srgb, var(--accent) 70%, white);
}

/* Login form card */
.login-form-card{
  background:var(--glass); border:1px solid var(--border);
  border-radius:18px; padding:1.6rem 1.5rem 1.4rem;
  box-shadow:0 30px 60px -25px rgba(0,0,0,.55);
  position:relative; overflow:hidden;
}
.login-form-card .beam{
  position:absolute; top:0; left:0; right:0; height:2px;
  background:linear-gradient(90deg, transparent, var(--accent), var(--accent-2), transparent);
}
.login-form-card h2{
  font-size:1.6rem !important; font-weight:800 !important;
  margin:0 0 .35rem !important; color:var(--text-strong) !important;
  letter-spacing:-.02em;
}
.login-form-card .sub{color:var(--text-muted); font-size:.92rem; margin:0 0 1rem}
.badge-device{
  display:inline-flex; align-items:center; gap:.4rem;
  font-size:.72rem; font-weight:700; color:var(--accent);
  background:var(--accent-soft);
  border:1px solid color-mix(in srgb, var(--accent) 30%, transparent);
  padding:.3rem .7rem; border-radius:999px; margin-bottom:.9rem;
}
.badge-device::before{content:""; width:6px; height:6px; border-radius:50%; background:currentColor}

/* Esconder logo gigante centrado en login */
.login-shell + div [data-testid="stImage"]{display:none}
"""


if "ui_theme" not in st.session_state:
    st.session_state["ui_theme"] = "dark"
st.markdown(_build_global_css(st.session_state["ui_theme"]), unsafe_allow_html=True)
# El theme toggle ahora vive DENTRO del topbar (ver `_render_topbar`),
# por eso ya no llamamos al toggle suelto flotante aqui.


_init_download_state()
_drain_download_queue()

# Variables para Playwright (Render / Docker)
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "/ms-playwright")
os.environ.setdefault("PYPPETEER_HOME", "/ms-playwright")

BASE_DIR = Path(__file__).parent
if getattr(sys, "frozen", False):
    RUNTIME_DIR = Path(os.getenv("APP_RUNTIME_DIR", Path(sys.executable).parent))
else:
    RUNTIME_DIR = Path(os.getenv("APP_RUNTIME_DIR", BASE_DIR))
DESC_DIR = RUNTIME_DIR / "descargas"
DESC_DIR.mkdir(exist_ok=True, parents=True)
LICENSE_CLIENT = LicensingClient()
SESSION_CACHE_DIR = Path(os.getenv("SESSION_CACHE_DIR", RUNTIME_DIR / ".session_cache"))
SESSION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
SESSION_CACHE = SESSION_CACHE_DIR / "session_cache.json"
ENABLE_SESSION_CACHE = os.getenv("ENABLE_SESSION_CACHE", "1").strip().lower() not in {"0", "false", "no"}
PREFERENCES_FILE = Path(os.getenv("USER_PREFS_PATH", SESSION_CACHE_DIR / "user_prefs.json"))
DOWNLOAD_RESUME_DIR = SESSION_CACHE_DIR / "download_resume"
DOWNLOAD_RESUME_DIR.mkdir(parents=True, exist_ok=True)
_SESSION_CACHE_MEMORY: dict[str, dict] = {}

def _generate_device_fingerprint() -> str:
    raw = f"{platform.node()}|{platform.system()}|{platform.release()}|{uuid.getnode()}"
    return hashlib.sha256(raw.encode()).hexdigest()


def _get_or_init_client_device_id() -> str | None:
    cached_id = st.session_state.get("_device_id")
    if cached_id:
        return str(cached_id)
    if "device_id" in st.query_params:
        try:
            st.query_params.pop("device_id", None)
        except Exception:
            st.query_params.clear()
    prefs = {}
    if PREFERENCES_FILE.exists():
        try:
            prefs = json.loads(PREFERENCES_FILE.read_text(encoding="utf-8"))
        except Exception:
            prefs = {}
    device_id = prefs.get("device_id")
    if device_id:
        st.session_state["_device_id"] = device_id
        return str(device_id)
    # Create a stable device id and persist it locally (avoids reload loop)
    device_id = uuid.uuid4().hex
    st.session_state["_device_id"] = device_id
    prefs["device_id"] = device_id
    try:
        PREFERENCES_FILE.write_text(json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return str(device_id)


def _get_device_id_from_query() -> str | None:
    device_id = st.query_params.get("device_id")
    if isinstance(device_id, list):
        device_id = device_id[0] if device_id else None
    return str(device_id) if device_id else None


def _session_cache_path(device_id: str | None) -> Path:
    if not device_id:
        return SESSION_CACHE
    safe_id = "".join(ch for ch in device_id if ch.isalnum() or ch in ("-", "_"))
    return SESSION_CACHE_DIR / f"session_cache_{safe_id}.json"


def _require_client_device_id() -> str | None:
    return _get_or_init_client_device_id()


def _persist_session_state():
    if not ENABLE_SESSION_CACHE:
        return
    device_id = st.session_state.get("_device_id") or _get_device_id_from_query()
    cache_path = _session_cache_path(device_id)
    payload = {}
    for key in (
        "auth_token",
        "user_email",
        "license_validated",
        "device_fingerprint",
        "license_last_check",
    ):
        if key in st.session_state:
            payload[key] = st.session_state[key]
    if payload:
        _SESSION_CACHE_MEMORY[str(device_id or "default")] = dict(payload)
        cache_path.write_text(json.dumps(payload))


def _load_cached_session(device_id: str | None = None):
    if not ENABLE_SESSION_CACHE:
        return
    if "auth_token" in st.session_state:
        return
    cache_path = _session_cache_path(device_id)
    data = None
    if cache_path.exists():
        try:
            data = json.loads(cache_path.read_text())
        except Exception:
            data = None
    if data is None:
        data = _SESSION_CACHE_MEMORY.get(str(device_id or "default"))
    if not data:
        return
    for key, value in data.items():
        st.session_state[key] = value


def _clear_cached_session():
    if not ENABLE_SESSION_CACHE:
        return
    device_id = st.session_state.get("_device_id") or _get_device_id_from_query()
    cache_path = _session_cache_path(device_id)
    _SESSION_CACHE_MEMORY.pop(str(device_id or "default"), None)
    try:
        cache_path.unlink(missing_ok=True)
    except Exception:
        pass


def _clear_cached_auth_only():
    if not ENABLE_SESSION_CACHE:
        return
    device_id = st.session_state.get("_device_id") or _get_device_id_from_query()
    cache_path = _session_cache_path(device_id)
    data = None
    if cache_path.exists():
        try:
            data = json.loads(cache_path.read_text())
        except Exception:
            data = None
    if data is None:
        data = _SESSION_CACHE_MEMORY.get(str(device_id or "default"))
    if not data:
        return
    data.pop("auth_token", None)
    _SESSION_CACHE_MEMORY[str(device_id or "default")] = dict(data)
    try:
        if data:
            cache_path.write_text(json.dumps(data))
        else:
            cache_path.unlink(missing_ok=True)
    except Exception:
        pass


def _load_user_preferences():
    if st.session_state.get("_prefs_loaded"):
        return
    data = {}
    if PREFERENCES_FILE.exists():
        try:
            data = json.loads(PREFERENCES_FILE.read_text())
        except Exception:
            data = {}
    st.session_state["download_base_dir"] = data.get("download_base_dir", str(DESC_DIR))
    st.session_state["consolidate_base_dir"] = data.get(
        "consolidate_base_dir",
        st.session_state["download_base_dir"],
    )
    st.session_state["first_use_tour_completed"] = bool(data.get("first_use_tour_completed", False))
    st.session_state["first_use_tour_prompt_dismissed"] = bool(data.get("first_use_tour_prompt_dismissed", False))
    st.session_state["_prefs_loaded"] = True


def _persist_user_preferences():
    data = {}
    if PREFERENCES_FILE.exists():
        try:
            data = json.loads(PREFERENCES_FILE.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    data["download_base_dir"] = st.session_state.get("download_base_dir", str(DESC_DIR))
    data["consolidate_base_dir"] = st.session_state.get(
        "consolidate_base_dir",
        data.get("download_base_dir", str(DESC_DIR)),
    )
    data["first_use_tour_completed"] = bool(st.session_state.get("first_use_tour_completed", False))
    data["first_use_tour_prompt_dismissed"] = bool(st.session_state.get("first_use_tour_prompt_dismissed", False))
    try:
        PREFERENCES_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception as err:
        st.warning(f"No se pudo guardar la configuración local: {err}")


def _onboarding_steps() -> list[dict[str, str]]:
    return [
        {
            "title": "Credenciales",
            "content": "Ingresa RUC y clave del SRI. Luego selecciona el origen: Recibidos o Emitidos.",
        },
        {
            "title": "Fechas y filtros",
            "content": "Configura modo de fecha (Mes y dia, Rango o Año completo) y el tipo de comprobante.",
        },
        {
            "title": "Formatos",
            "content": "Elige XML y/o PDF según tu necesidad. En Emitidos puedes marcar formatos individuales.",
        },
        {
            "title": "Carpeta de descarga",
            "content": "Selecciona la carpeta base donde se guardarán los documentos y reportes.",
        },
        {
            "title": "Ejecución",
            "content": "Haz clic en Iniciar proceso y espera el resumen final con reportes descargables.",
        },
    ]


def _start_first_use_tour(reset_step: bool = True) -> None:
    st.session_state["first_use_tour_active"] = True
    st.session_state["first_use_tour_prompt_dismissed"] = True
    _persist_user_preferences()
    if reset_step:
        st.session_state["first_use_tour_step"] = 0


def _finish_first_use_tour() -> None:
    st.session_state["first_use_tour_active"] = False
    st.session_state["first_use_tour_completed"] = True
    st.session_state["first_use_tour_prompt_dismissed"] = True
    _persist_user_preferences()


def _render_first_use_tour() -> None:
    steps = _onboarding_steps()
    total = len(steps)
    step_idx = int(st.session_state.get("first_use_tour_step", 0))
    step_idx = max(0, min(step_idx, total - 1))
    st.session_state["first_use_tour_step"] = step_idx
    step = steps[step_idx]

    st.markdown("#### Tour de primer uso")
    st.info(f"Paso {step_idx + 1} de {total}: {step['title']}")
    st.caption(step["content"])

    col_prev, col_next, col_later, col_finish, col_skip = st.columns([1, 1, 1, 1, 1.1])
    with col_prev:
        if st.button("Anterior", key="tour_prev", use_container_width=True, disabled=step_idx == 0):
            st.session_state["first_use_tour_step"] = max(0, step_idx - 1)
            st.rerun()
    with col_next:
        if st.button("Siguiente", key="tour_next", use_container_width=True, disabled=step_idx >= total - 1):
            st.session_state["first_use_tour_step"] = min(total - 1, step_idx + 1)
            st.rerun()
    with col_later:
        if st.button("Mas tarde", key="tour_later", use_container_width=True):
            st.session_state["first_use_tour_active"] = False
            st.session_state["first_use_tour_prompt_dismissed"] = True
            _persist_user_preferences()
            st.rerun()
    with col_finish:
        if st.button("Finalizar", key="tour_finish", use_container_width=True):
            _finish_first_use_tour()
            st.success("Tour completado. Puedes verlo nuevamente desde la pestaña Ayuda.")
            st.rerun()
    with col_skip:
        if st.button("Omitir y no mostrar", key="tour_skip", use_container_width=True):
            _finish_first_use_tour()
            st.rerun()


def _render_first_use_prompt() -> None:
    st.write("Quieres ver un recorrido rapido para aprender el uso del sistema?")
    col_p1, col_p2, col_p3 = st.columns([1, 1, 1.1])
    with col_p1:
        if st.button("Iniciar recorrido", key="tour_prompt_start", use_container_width=True):
            _start_first_use_tour(reset_step=True)
            st.rerun()
    with col_p2:
        if st.button("Mas tarde", key="tour_prompt_later", use_container_width=True):
            st.session_state["first_use_tour_prompt_dismissed"] = True
            _persist_user_preferences()
            st.rerun()
    with col_p3:
        if st.button("No volver a mostrar", key="tour_prompt_never", use_container_width=True):
            _finish_first_use_tour()
            st.rerun()


def _get_download_base_dir() -> Path:
    base = Path(st.session_state.get("download_base_dir") or str(DESC_DIR)).expanduser()
    base.mkdir(parents=True, exist_ok=True)
    return base


def _current_download_checkpoint_path(user_email: str | None = None) -> Path:
    email = (user_email or st.session_state.get("user_email") or "").strip().lower()
    return build_download_checkpoint_path(DOWNLOAD_RESUME_DIR, email)


def _format_download_resume_period(summary: dict | None) -> str:
    summary = summary or {}
    meses = [
        "",
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]
    try:
        anio = int(summary.get("anio") or 0)
    except Exception:
        anio = 0
    try:
        mes = int(summary.get("mes") or 0)
    except Exception:
        mes = 0
    try:
        mes_fin = int(summary.get("mes_fin") or 0)
    except Exception:
        mes_fin = 0
    try:
        dia = int(summary.get("dia") or 0)
    except Exception:
        dia = 0

    if anio and mes == 1 and mes_fin == 12 and dia in (0, None):
        return f"Año completo {anio}"
    if anio and mes and mes_fin and mes_fin >= mes and dia in (0, None):
        return f"{meses[mes]} {anio} a {meses[mes_fin]} {anio}"
    if anio and mes and dia not in (0, None):
        return f"{dia:02d}/{mes:02d}/{anio}"
    if anio and mes:
        return f"{meses[mes]} {anio}"
    return "Periodo no disponible"


def _load_pending_download_checkpoint_for_current_user() -> dict | None:
    path = _current_download_checkpoint_path()
    data = load_download_checkpoint(path)
    if not data:
        return None
    if str(data.get("status") or "").strip().lower() == "completed":
        delete_download_checkpoint(path)
        return None
    data["_path"] = str(path)
    return data


def _start_download_process(
    params: dict,
    *,
    resume_download: bool = False,
    checkpoint_payload: dict | None = None,
) -> None:
    checkpoint_file = _current_download_checkpoint_path()
    if resume_download:
        if load_download_checkpoint(checkpoint_file):
            mark_download_checkpoint_running(checkpoint_file)
        else:
            payload = checkpoint_payload or build_checkpoint_payload(st.session_state.get("user_email"), params)
            save_download_checkpoint(checkpoint_file, payload)
            mark_download_checkpoint_running(checkpoint_file)
    else:
        payload = checkpoint_payload or build_checkpoint_payload(st.session_state.get("user_email"), params)
        save_download_checkpoint(checkpoint_file, payload)

    st.session_state.download_messages = []
    st.session_state.download_result = None
    st.session_state.download_error = None
    st.session_state.download_registered = False
    st.session_state.download_finished_modal_open = False
    st.session_state.download_status = "running"
    st.session_state.running_notice_ts = time.time()
    st.session_state.stop_notice_ts = None
    st.session_state.manual_consultar_hint = None
    st.session_state.manual_consultar_hint_ts = None
    st.session_state.download_params = dict(params)

    worker_params = dict(params)
    worker_params["checkpoint_path"] = str(checkpoint_file)
    worker_params["resume_download"] = bool(resume_download)
    worker = threading.Thread(
        target=_download_worker,
        args=(worker_params, st.session_state.download_queue),
        daemon=True,
    )
    st.session_state.download_thread = worker
    worker.start()


def _select_directory_dialog(initial_dir: str | None = None):
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as err:
        return None, f"No se puede abrir el selector nativo: {err}"
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(initialdir=initial_dir or str(Path.home()))
        root.destroy()
        if path:
            return path, None
        return None, None
    except Exception as err:
        return None, str(err)


def _request_password_reset(email: str) -> None:
    LICENSE_CLIENT.request_password_reset(email.strip())


def _confirm_password_reset(token: str, new_password: str) -> None:
    LICENSE_CLIENT.confirm_password_reset(token.strip(), new_password)


def _preview_password_reset(token: str) -> str:
    data = LICENSE_CLIENT.preview_password_reset(token.strip())
    return str(data.get("email") or "").strip()


def _handle_reset_query_token():
    params = st.query_params
    if "reset_request" in params:
        st.session_state["reset_request_mode"] = True
        st.query_params.clear()
        return
    token_values = params.get("reset_token")
    if not token_values:
        return
    if isinstance(token_values, str):
        token = token_values
    else:
        token = token_values[0]
    try:
        email = _preview_password_reset(token)
    except Exception:
        st.warning("El enlace de recuperación no es válido o ya expiró.")
        st.query_params.clear()
        return
    if not email:
        st.warning("No se pudo validar el enlace de recuperación.")
        st.query_params.clear()
        return
    st.session_state["password_recovery_mode"] = True
    st.session_state["active_reset_token"] = token
    st.session_state["recovery_email"] = email
    st.query_params.clear()


def _normalize_compare_text(value: str | None) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")
    return text


def _canonical_tipo(value: str | None) -> str:
    text = _normalize_compare_text(value)
    if "retencion" in text:
        return "retenciones"
    if "nota" in text and "credito" in text:
        return "notas_de_credito"
    if "nota" in text and "debito" in text:
        return "notas_de_debito"
    if "liquidacion" in text:
        return "liquidacion_de_compra"
    if "guia" in text or "remision" in text:
        return "guia_de_remision"
    if "factura" in text:
        return "facturas"
    return text


def _infer_origin_and_status_from_path(path: Path) -> tuple[str, str]:
    parts = [_normalize_compare_text(part) for part in path.parts]
    if "emitidos" in parts:
        origen = "Emitidos"
    elif "recibidos" in parts:
        origen = "Recibidos"
    else:
        origen = "Desconocido"
    estado = ""
    if origen == "Emitidos":
        if any(part == "no_autorizados" for part in parts):
            estado = "No autorizados"
        elif any(part == "autorizados" for part in parts):
            estado = "Autorizados"
    return origen, estado


def _parse_report_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    candidatos = [
        text,
        text[:10],
        text.replace("T", " ")[:19],
    ]
    formatos = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ]
    for candidato in candidatos:
        for formato in formatos:
            try:
                return datetime.strptime(candidato, formato).date()
            except Exception:
                continue
    return None


def _report_row_key(row: dict) -> str:
    if not isinstance(row, dict):
        return ""
    for key in ("claveAcceso", "CLAVE_ACCESO", "Clave de Acceso", "Número de Autorización"):
        value = str(row.get(key) or "").strip()
        if value:
            return value
    numero = str(
        row.get("numeroComprobante")
        or row.get("SERIE_COMPROBANTE")
        or row.get("Número de Documento de Sustento")
        or row.get("Secuencial")
        or ""
    ).strip()
    fecha = str(
        row.get("fechaEmision")
        or row.get("FECHA_EMISION")
        or row.get("Fecha de Emisión")
        or row.get("Fecha de Autorización")
        or ""
    ).strip()
    ruc = str(row.get("rucEmisor") or row.get("RUC_EMISOR") or row.get("RUC Emisor") or "").strip()
    return "|".join(part for part in (numero, fecha, ruc) if part)


def _map_xml_to_standard_row(cabecera: dict, detalles: list[dict]) -> dict:
    row = {col: "" for col in PDF_REPORT_COLUMNS}
    row["tipoDocumento"] = cabecera.get("DESCRIPCION_DOC", "")
    row["rucEmisor"] = cabecera.get("RUC_EMISOR", "")
    row["razonSocialEmisor"] = cabecera.get("RAZON_SOCIAL_EMISOR", "")
    row["nombreComercial"] = cabecera.get("NOMBRE_COMERCIAL_EMISOR", "")
    row["direccionMatrizEmisor"] = cabecera.get("DIR_MATRIZ", "")
    row["direccionSucursalEmisor"] = cabecera.get("DIR_ESTABLECIMIENTO", "")
    row["obligadoContabilidad"] = cabecera.get("OBLIGADO_CONTABILIDAD", "")
    row["tipoContribuyenteRIMPE"] = cabecera.get("CONTRIBUYENTE_RIMPE", "")
    row["numeroComprobante"] = cabecera.get("SERIE_COMPROBANTE", "")
    row["establecimiento"] = cabecera.get("ESTAB", "")
    row["puntoEmision"] = cabecera.get("PTO_EMI", "")
    row["secuencial"] = cabecera.get("SECUENCIAL", "")
    row["fechaEmision"] = cabecera.get("FECHA_EMISION", "")
    row["fechaAutorizacion"] = cabecera.get("FECHA_AUTORIZACION", "")
    row["razonSocialComprador"] = cabecera.get("RAZON_SOCIAL_COMPRADOR", "")
    row["identificacionComprador"] = cabecera.get("IDENTIFICACION_COMPRADOR", "")
    row["direccionComprador"] = cabecera.get("DIRECCION_COMPRADOR", "")
    row["comprobanteModificado"] = cabecera.get("NUM_DOC_MODIFICADO", "")
    row["fechaEmisionModificado"] = cabecera.get("FECHA_EMISION_DOC_SUSTENTO", "")
    row["razonModificacion"] = cabecera.get("MOTIVO_MODIFICACION", "")
    row["valorModificacion"] = cabecera.get("VALOR_MODIFICACION") or cabecera.get("VALOR_MODIFICACION_XML") or ""
    row["descripcionesProductos"] = " | ".join(
        str(det.get("DESCRIPCION") or "").strip()
        for det in (detalles or [])
        if str(det.get("DESCRIPCION") or "").strip()
    )
    row["subtotalSinImpuestos"] = cabecera.get("TOTAL_SIN_IMPUESTOS", "")
    row["totalDescuento"] = cabecera.get("TOTAL_DESCUENTO", "")
    row["propina"] = cabecera.get("PROPINA", "")
    row["valorTotal"] = cabecera.get("IMPORTE_TOTAL", "") or cabecera.get("VALOR_TOTAL", "")
    row["ambiente"] = cabecera.get("AMBIENTE", "")
    row["emision"] = cabecera.get("TIPO_EMISION", "")
    row["claveAcceso"] = cabecera.get("CLAVE_ACCESO", "")
    row["informacionAdicional"] = cabecera.get("INFO_ADICIONAL_JSON", "")
    return row


def _map_xml_to_retencion_row(cabecera: dict, retenciones: list[dict]) -> dict:
    row = {col: "" for col in RETENCION_REPORT_COLUMNS}
    row["tipoDocumento"] = cabecera.get("DESCRIPCION_DOC", "")
    row["rucEmisor"] = cabecera.get("RUC_EMISOR", "")
    row["razonSocialEmisor"] = cabecera.get("RAZON_SOCIAL_EMISOR", "")
    row["nombreComercial"] = cabecera.get("NOMBRE_COMERCIAL_EMISOR", "")
    row["direccionMatrizEmisor"] = cabecera.get("DIR_MATRIZ", "")
    row["direccionSucursalEmisor"] = cabecera.get("DIR_ESTABLECIMIENTO", "")
    row["obligadoContabilidad"] = cabecera.get("OBLIGADO_CONTABILIDAD", "")
    row["fechaAutorizacion"] = cabecera.get("FECHA_AUTORIZACION", "")
    row["ambiente"] = cabecera.get("AMBIENTE", "")
    row["emision"] = cabecera.get("TIPO_EMISION", "")
    row["numeroComprobante"] = cabecera.get("SERIE_COMPROBANTE", "")
    row["establecimiento"] = cabecera.get("ESTAB", "")
    row["puntoEmision"] = cabecera.get("PTO_EMI", "")
    row["secuencial"] = cabecera.get("SECUENCIAL", "")
    row["fechaEmision"] = cabecera.get("FECHA_EMISION", "")
    row["razonSocialSujetoRetenido"] = cabecera.get("RAZON_SOCIAL_COMPRADOR", "")
    row["identificacionSujetoRetenido"] = cabecera.get("IDENTIFICACION_COMPRADOR", "")
    row["claveAcceso"] = cabecera.get("CLAVE_ACCESO", "")
    row["informacionAdicional"] = cabecera.get("INFO_ADICIONAL_JSON", "")
    iva_idx = 0
    ir_idx = 0
    for item in retenciones or []:
        tipo_imp = _normalize_compare_text(item.get("TIPO_IMPUESTO", ""))
        suffix = ""
        if "iva" in tipo_imp:
            suffix = "" if iva_idx == 0 else "_1"
            iva_idx += 1
            row[f"Base_Imponible_Ret_IVA{suffix}"] = item.get("BASE_IMPONIBLE", "")
            row[f"Impuesto_Ret_IVA{suffix}"] = item.get("TIPO_IMPUESTO", "")
            row[f"Porcentaje_Ret_IVA{suffix}"] = item.get("PORCENTAJE_RETENER", "")
            row[f"Valor_Retenido_IVA{suffix}"] = item.get("VALOR_RETENIDO", "")
        else:
            suffix = "" if ir_idx == 0 else "_1"
            ir_idx += 1
            row[f"Base_Imponible_Ret_IR{suffix}"] = item.get("BASE_IMPONIBLE", "")
            row[f"Impuesto_Ret_IR{suffix}"] = item.get("TIPO_IMPUESTO", "")
            row[f"Porcentaje_Ret_IR{suffix}"] = item.get("PORCENTAJE_RETENER", "")
            row[f"Valor_Retenido_IR{suffix}"] = item.get("VALOR_RETENIDO", "")
        row["Comprobante_Sustento"] = row["Comprobante_Sustento"] or item.get("DOC_SUSTENTO_TIPO", "")
        row["Numero_Sustento"] = row["Numero_Sustento"] or item.get("DOC_SUSTENTO_SECUENCIAL", "")
        row["Fecha_Emision_Sustento"] = row["Fecha_Emision_Sustento"] or item.get("FECHA_EMISION_DOC_SUSTENTO", "")
        row["Ejercicio_Fiscal"] = row["Ejercicio_Fiscal"] or item.get("PERIODO_FISCAL", "")
    return row


def _build_custom_report_from_folder(
    base_dir: Path,
    *,
    origen: str,
    tipo: str,
    fecha_inicio: date,
    fecha_fin: date,
    estado_emitidos: str | None = None,
) -> dict:
    base_dir = Path(base_dir).expanduser()
    target_tipo = _canonical_tipo(tipo)
    is_retencion = target_tipo == "retenciones"
    is_nota_credito = target_tipo == "notas_de_credito"
    is_nota_debito = target_tipo == "notas_de_debito"
    is_factura_emitida = target_tipo == "facturas" and origen == "Emitidos"
    is_liquidacion_emitida = target_tipo == "liquidacion_de_compra" and origen == "Emitidos"
    rows: list[dict] = []
    seen_keys: set[str] = set()
    xml_count = 0
    pdf_count = 0
    errores: list[str] = []

    for xml_path in sorted(base_dir.rglob("*.xml")):
        path_origen, path_estado = _infer_origin_and_status_from_path(xml_path)
        if path_origen != origen:
            continue
        if origen == "Emitidos" and estado_emitidos and path_estado and path_estado != estado_emitidos:
            continue
        if is_retencion and origen == "Emitidos":
            row = _extraer_datos_xml_retencion_emitido(xml_path)
            fecha_doc = _parse_report_date(row.get("Fecha de Emisión")) or _parse_report_date(row.get("Fecha de Autorización"))
            if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
                continue
            key = _report_row_key(row)
            if key:
                seen_keys.add(key)
            rows.append(row)
            xml_count += 1
            continue
        if is_nota_credito and origen == "Emitidos":
            row = _extraer_datos_xml_nota_credito_emitido(xml_path)
            fecha_doc = _parse_report_date(row.get("Fecha de Emisión")) or _parse_report_date(row.get("Fecha de Autorización"))
            if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
                continue
            key = _report_row_key(row)
            if key:
                seen_keys.add(key)
            rows.append(row)
            xml_count += 1
            continue
        if is_nota_debito and origen == "Emitidos":
            row = _extraer_datos_xml_nota_debito_emitido(xml_path)
            fecha_doc = _parse_report_date(row.get("Fecha de Emisión")) or _parse_report_date(row.get("Fecha de Autorización"))
            if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
                continue
            key = _report_row_key(row)
            if key:
                seen_keys.add(key)
            rows.append(row)
            xml_count += 1
            continue
        if is_factura_emitida:
            row = _extraer_datos_xml_factura_emitido(xml_path)
            fecha_doc = _parse_report_date(row.get("Fecha de Emisión")) or _parse_report_date(row.get("Fecha de Autorización"))
            if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
                continue
            key = _report_row_key(row)
            if key:
                seen_keys.add(key)
            rows.append(row)
            xml_count += 1
            continue
        if is_liquidacion_emitida:
            row = _extraer_datos_xml_liquidacion_compra_emitido(xml_path)
            fecha_doc = _parse_report_date(row.get("fechaEmision")) or _parse_report_date(row.get("fechaAutorizacion"))
            if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
                continue
            key = _report_row_key(row)
            if key:
                seen_keys.add(key)
            rows.append(row)
            xml_count += 1
            continue
        cabecera, detalles, _, _, _, retenciones, error_entry, _ = _parse_recibido_xml(xml_path)
        if error_entry or not cabecera:
            continue
        if _canonical_tipo(cabecera.get("DESCRIPCION_DOC")) != target_tipo:
            continue
        fecha_doc = _parse_report_date(cabecera.get("FECHA_EMISION")) or _parse_report_date(cabecera.get("FECHA_AUTORIZACION"))
        if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
            continue
        row = _map_xml_to_retencion_row(cabecera, retenciones) if is_retencion else _map_xml_to_standard_row(cabecera, detalles)
        key = _report_row_key(row)
        if key:
            seen_keys.add(key)
        rows.append(row)
        xml_count += 1

    for pdf_path in sorted(base_dir.rglob("*.pdf")):
        path_origen, path_estado = _infer_origin_and_status_from_path(pdf_path)
        if path_origen != origen:
            continue
        if origen == "Emitidos" and estado_emitidos and path_estado and path_estado != estado_emitidos:
            continue
        try:
            if is_retencion and origen == "Emitidos":
                row = _extraer_datos_pdf_retencion_emitido(pdf_path)
            elif is_nota_credito and origen == "Emitidos":
                row = _extraer_datos_pdf_nota_credito_emitido(pdf_path)
            elif is_nota_debito and origen == "Emitidos":
                row = _extraer_datos_pdf_nota_debito_emitido(pdf_path)
            elif is_factura_emitida:
                row = _extraer_datos_pdf_factura_emitido(pdf_path)
            elif is_liquidacion_emitida:
                row = _extraer_datos_pdf_liquidacion_compra_emitido(pdf_path)
            else:
                row = _extraer_datos_pdf_por_tipo_layout_first(
                    pdf_path,
                    es_retencion=is_retencion,
                    es_nota_credito=is_nota_credito,
                    es_nota_debito=is_nota_debito,
                )
        except Exception as err:
            errores.append(f"{pdf_path.name}: {err}")
            continue
        tipo_row = row.get("tipoDocumento") if isinstance(row, dict) else ""
        if is_retencion and origen == "Emitidos":
            tipo_row = "retenciones"
        elif is_nota_credito and origen == "Emitidos":
            tipo_row = "notas_de_credito"
        elif is_nota_debito and origen == "Emitidos":
            tipo_row = "notas_de_debito"
        elif is_factura_emitida:
            tipo_row = "factura"
        elif is_liquidacion_emitida:
            tipo_row = "liquidacion_de_compra"
        if _canonical_tipo(tipo_row) != target_tipo:
            continue
        key = _report_row_key(row)
        if key and key in seen_keys:
            continue
        fecha_doc = (
            _parse_report_date(row.get("fechaEmision"))
            or _parse_report_date(row.get("fechaAutorizacion"))
            or _parse_report_date(row.get("Fecha de Emisión"))
            or _parse_report_date(row.get("Fecha de Autorización"))
        )
        if not fecha_doc or fecha_doc < fecha_inicio or fecha_doc > fecha_fin:
            continue
        if key:
            seen_keys.add(key)
        rows.append(row)
        pdf_count += 1

    if not rows:
        return {
            "ok": False,
            "message": "No se encontraron documentos en el rango indicado.",
            "xml_count": 0,
            "pdf_count": 0,
            "errors": errores,
        }

    report_dir = base_dir / "Reportes_personalizados"
    report_dir.mkdir(parents=True, exist_ok=True)
    origen_slug = _normalize_compare_text(origen)
    tipo_slug = _canonical_tipo(tipo)
    sufijo = f"{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}"
    output_path = report_dir / f"reporte_personalizado_{origen_slug}_{tipo_slug}_{sufijo}.xlsx"
    if output_path.exists():
        timestamp = datetime.now().strftime("%H%M%S")
        output_path = report_dir / f"reporte_personalizado_{origen_slug}_{tipo_slug}_{sufijo}_{timestamp}.xlsx"

    guardado = (
        _guardar_reporte_pdf_retencion_emitidos_excel(rows, output_path)
        if is_retencion and origen == "Emitidos"
        else _guardar_reporte_pdf_nota_credito_emitidos_excel(rows, output_path)
        if is_nota_credito and origen == "Emitidos"
        else _guardar_reporte_pdf_nota_debito_emitidos_excel(rows, output_path)
        if is_nota_debito and origen == "Emitidos"
        else _guardar_reporte_pdf_factura_emitidos_excel(rows, output_path)
        if is_factura_emitida
        else _guardar_reporte_pdf_retencion_excel(rows, output_path)
        if is_retencion
        else _guardar_reporte_pdf_excel(rows, output_path)
    )
    if not guardado or not output_path.exists():
        return {
            "ok": False,
            "message": "No se pudo generar el archivo Excel del reporte personalizado.",
            "xml_count": xml_count,
            "pdf_count": pdf_count,
            "errors": errores,
        }
    return {
        "ok": True,
        "path": output_path,
        "xml_count": xml_count,
        "pdf_count": pdf_count,
        "rows": len(rows),
        "errors": errores,
    }


def _inject_login_background_css():
    """Inyecta:
      - CSS que aplica `assets/hero.jpg` como background fixed full-viewport
        sobre `.stApp`, con un gradient overlay (mas oscuro hacia la derecha
        para legibilidad del form).
      - Marker `.is-login-page` para que las reglas CSS scopeadas se activen
        (selector `:has(.is-login-page)`).
      - Reglas que reposicionan el main container de Streamlit y dejan los
        widgets alineados a la derecha en un card flotante semi-transparente.
      - Brand badge fijo abajo-izquierda (lv-brand-floating).

    Se llama una sola vez por render de login / reset / recovery, antes de
    cualquier otro widget. NO usa `st.columns`.
    """
    hero_uri = _asset_data_uri("hero.jpg")
    if not hero_uri:
        # Sin asset, no podemos hacer full-bleed — fallback minimal.
        logo_html = _logo_html(220)
        if logo_html:
            st.markdown(logo_html, unsafe_allow_html=True)
        return
    # NOTA: el f-string esta indentado adentro de la funcion. Streamlit's
    # markdown procesa lineas con 4+ espacios como bloques de codigo, asi
    # que strippeamos TODA leading whitespace de cada linea antes de
    # rendear (textwrap.dedent no alcanza porque los selectores CSS multi-
    # linea tienen indentacion adicional de continuacion que sigue arriba
    # del umbral de 4 espacios despues del dedent).
    _login_html = f"""
        <div class="is-login-page" aria-hidden="true"></div>
        <style>
        /* Hero como background del viewport entero + overlay para
           que el form sea legible. El overlay es mas opaco hacia la
           derecha (donde flota el card). */
        .stApp:has(.is-login-page){{
          background:
            linear-gradient(to right,
              rgba(7,11,20,0.28) 0%,
              rgba(7,11,20,0.50) 42%,
              rgba(7,11,20,0.82) 68%,
              rgba(7,11,20,0.94) 90%,
              rgba(7,11,20,0.97) 100%
            ),
            url('{hero_uri}') no-repeat center / cover fixed !important;
          background-attachment: fixed !important;
        }}
        /* Convertimos el main container del login en un panel vertical de
           460px de ancho alineado a la derecha del viewport. Como TODO el
           contenido del login (titulo, subtitulo, card) vive dentro de
           este container, basta con constrenirlo aqui para que todos los
           hijos respeten el ancho y queden alineados verticalmente.
           min-height 100vh + flex centrado lo posiciona en el medio
           vertical. */
        .stApp:has(.is-login-page) [data-testid="stMainBlockContainer"],
        .stApp:has(.is-login-page) section.main > div.block-container{{
          min-height: 100vh !important;
          max-width: 460px !important;
          margin-left: auto !important;
          margin-right: clamp(2rem, 6vw, 5rem) !important;
          padding: 1.5rem 0 !important;
          display: flex !important;
          flex-direction: column !important;
          justify-content: center !important;
        }}
        /* Form como card flotante semi-transparente con blur. */
        .stApp:has(.is-login-page) [data-testid="stForm"]{{
          background: rgba(15, 23, 42, 0.82) !important;
          border: 1px solid var(--border) !important;
          border-radius: 18px !important;
          padding: 1.5rem 1.4rem 1.3rem !important;
          backdrop-filter: blur(10px);
          -webkit-backdrop-filter: blur(10px);
          box-shadow: 0 30px 60px -25px rgba(0,0,0,0.65) !important;
          position: relative;
          overflow: hidden;
        }}
        /* Beam superior del card (linea gradient verde-azul). */
        .stApp:has(.is-login-page) [data-testid="stForm"]::before{{
          content: "";
          position: absolute; top: 0; left: 0; right: 0; height: 2px;
          background: linear-gradient(90deg,
            transparent, var(--accent), var(--accent-2), transparent);
          pointer-events: none;
        }}
        /* Brand badge fijo abajo-izquierda (mockup). */
        .lv-brand-floating{{
          position: fixed;
          left: clamp(1.2rem, 2.5vw, 2.4rem);
          bottom: clamp(1.4rem, 4vh, 2.4rem);
          display: flex;
          align-items: center;
          gap: .65rem;
          z-index: 20;
          pointer-events: none;
        }}
        .lv-brand-floating .b-badge{{
          width: 36px; height: 36px;
          border-radius: 10px;
          display: grid; place-items: center;
          border: 1.5px solid var(--accent);
          background: rgba(7,13,8,0.6);
          flex: 0 0 auto;
        }}
        .lv-brand-floating b{{
          display: block;
          font-size: .92rem; font-weight: 800;
          color: #fff; line-height: 1.1;
          letter-spacing: .01em;
        }}
        .lv-brand-floating .kicker{{
          display: block;
          font-size: .66rem;
          letter-spacing: .14em;
          text-transform: uppercase;
          color: color-mix(in srgb, var(--accent) 75%, white);
          margin-top: 2px;
        }}
        /* En pantallas chicas, el form vuelve a centrarse y el brand
           se vuelve mas chico para no superponer el card. */
        @media (max-width: 720px) {{
          .stApp:has(.is-login-page) [data-testid="stMainBlockContainer"]
            [data-testid="stVerticalBlock"]:first-of-type{{
            margin-left: auto; margin-right: auto;
            max-width: 100%;
          }}
          .lv-brand-floating .kicker{{ display: none; }}
        }}

        /* === Header centrado ARRIBA del card === */
        .auth-header-centered{{
          text-align: center;
          margin: 0 0 1.1rem 0;
        }}
        .auth-header-centered h2{{
          font-size: 1.75rem !important;
          font-weight: 800 !important;
          margin: 0 !important;
          color: var(--text-strong) !important;
          letter-spacing: -.02em;
          line-height: 1.15;
        }}
        .auth-header-centered p{{
          margin: .45rem 0 0 !important;
          color: var(--text-muted) !important;
          font-size: .92rem;
        }}

        /* === Card unificado (badge + form + link) ===
           Cuando un `st.container()` contiene el marker `.is-login-card`,
           su stVerticalBlock interno recibe el aspecto de card: bg semi
           transparente, borde, radius, beam superior gradient. */
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card){{
          background: rgba(15, 23, 42, 0.82) !important;
          border: 1px solid var(--border) !important;
          border-radius: 18px !important;
          padding: 1.3rem 1.4rem 1.2rem !important;
          backdrop-filter: blur(10px);
          -webkit-backdrop-filter: blur(10px);
          box-shadow: 0 30px 60px -25px rgba(0,0,0,0.65) !important;
          position: relative;
          overflow: hidden;
        }}
        /* Beam superior del card unificado (linea gradient verde-azul). */
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card)::before{{
          content: "";
          position: absolute; top: 0; left: 0; right: 0; height: 2px;
          background: linear-gradient(90deg,
            transparent, var(--accent), var(--accent-2), transparent);
          pointer-events: none;
        }}

        /* El marker no se muestra. */
        .is-login-card{{display: none;}}

        /* Form DENTRO del card unificado: quitar su propio bg/borde/sombra
           para que no se vea "card dentro de card". El padding interno lo
           da el container padre. */
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card) [data-testid="stForm"]{{
          background: transparent !important;
          border: 0 !important;
          border-radius: 0 !important;
          padding: 0 !important;
          backdrop-filter: none;
          -webkit-backdrop-filter: none;
          box-shadow: none !important;
        }}
        /* El form ya tiene un `::before` (beam) — lo eliminamos porque
           el beam ahora vive en el container padre. */
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card) [data-testid="stForm"]::before{{
          content: none !important;
        }}

        /* Link "¿Olvidaste tu contraseña?" / "Volver" centrado al final
           del card. */
        .stApp:has(.is-login-page) .auth-reset-wrap{{
          text-align: center;
          margin: .6rem 0 0;
        }}
        .stApp:has(.is-login-page) .auth-reset-link{{
          color: var(--accent);
          text-decoration: none;
          font-weight: 600;
          font-size: .92rem;
        }}
        .stApp:has(.is-login-page) .auth-reset-link:hover{{
          text-decoration: underline;
        }}

        /* Boton "Volver a iniciar sesión" estilizado como link cuando vive
           dentro del card unificado (reset_request y password_recovery). */
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card)
            [class*="st-key-btn_back_login"] button{{
          background: transparent !important;
          border: 0 !important;
          color: var(--accent) !important;
          font-weight: 600 !important;
          padding: .2rem 0 !important;
          box-shadow: none !important;
          min-height: auto !important;
        }}
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card)
            [class*="st-key-btn_back_login"] button:hover{{
          text-decoration: underline;
          background: transparent !important;
          color: var(--accent) !important;
        }}
        .stApp:has(.is-login-page) [data-testid="stVerticalBlock"]:has(
            > [data-testid="element-container"] > .is-login-card)
            [class*="st-key-btn_back_login"]{{
          text-align: center;
          margin-top: .4rem;
        }}
        </style>
        <div class="lv-brand-floating">
          <span class="b-badge">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
              <path d="M12 3v11m0 0l4-4m-4 4l-4-4M5 20h14"
                stroke="currentColor" stroke-width="2"
                stroke-linecap="round" stroke-linejoin="round"
                style="color:var(--accent)"/>
            </svg>
          </span>
          <span><b>ROBOT&nbsp;SRI&nbsp;AUDIT</b><span class="kicker">Descarga y auditoría del SRI</span></span>
        </div>
        """
    # Strippear leading whitespace de cada linea: el CSS no se altera (los
    # selectores multi-linea siguen siendo validos), y se evita que
    # markdown lo trate como bloque de codigo.
    _login_html = "\n".join(line.lstrip() for line in _login_html.split("\n"))
    # Usamos st.html (Streamlit 1.33+) en vez de st.markdown porque st.html
    # bypassa el parser de markdown completamente. st.markdown procesa
    # caracteres especiales (backticks, asteriscos, indentacion) dentro del
    # contenido aunque este envuelto en <style>, lo que rompia el bloque y
    # mostraba CSS como texto en la pantalla. Si st.html no esta disponible
    # (Streamlit < 1.33), cae a markdown como fallback.
    if hasattr(st, "html"):
        st.html(_login_html)
    else:
        st.markdown(_login_html, unsafe_allow_html=True)


def _render_auth_header(title: str, subtitle: str):
    """Encabezado centrado ARRIBA del card del login (titulo + subtitulo).
    El badge y la beam ahora viven adentro del card (en `_render_login` y
    sus variantes), no aqui."""
    st.markdown(
        f"""
        <div class="auth-header-centered">
          <h2>{html.escape(title)}</h2>
          <p>{html.escape(subtitle)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_reset_request():
    # Layout: hero como background full-viewport; widgets en columna unica
    # alineados a la derecha por CSS (sin st.columns). Titulo + subtitulo
    # van centrados arriba del card; badge + form + link van adentro del
    # card unificado.
    _inject_login_background_css()
    _render_auth_header(
        "Recuperar contraseña",
        "Te enviaremos un enlace para restablecerla",
    )
    card = st.container()
    with card:
        st.markdown('<div class="is-login-card"></div>', unsafe_allow_html=True)
        st.markdown(
            '<span class="badge-device">Sesión vinculada al dispositivo</span>',
            unsafe_allow_html=True,
        )
        with st.form("password_request_form"):
            email = st.text_input(
                "Correo electrónico",
                value=st.session_state.get("recovery_email", ""),
            )
            send = st.form_submit_button("Enviar enlace", type="primary")
        if send:
            if not email:
                st.error("Ingresa el correo registrado.")
            else:
                try:
                    _request_password_reset(email.strip())
                    st.success("Si el correo existe, enviaremos un enlace de recuperación.")
                    st.session_state["reset_request_mode"] = False
                    st.session_state["password_recovery_mode"] = False
                    st.query_params.clear()
                    st.rerun()
                except Exception as err:
                    st.error(f"No se pudo enviar el correo: {err}")
        if st.button("← Volver a iniciar sesión", key="btn_back_login_reset_request"):
            st.session_state["reset_request_mode"] = False
            st.session_state.pop("recovery_email", None)
            st.query_params.clear()
            st.rerun()


def _render_password_recovery():
    st.session_state.setdefault("password_recovery_mode", False)
    _inject_login_background_css()
    _render_auth_header(
        "Restablecer contraseña",
        "Crea tu nueva contraseña para finalizar el acceso",
    )
    active_token = st.session_state.get("active_reset_token") or ""
    recovery_email = st.session_state.get("recovery_email") or ""
    card = st.container()
    with card:
        st.markdown('<div class="is-login-card"></div>', unsafe_allow_html=True)
        if not active_token or not recovery_email:
            st.warning("Abre el enlace de recuperación desde tu correo para continuar.")
        else:
            st.markdown(
                '<span class="badge-device">Sesión vinculada al dispositivo</span>',
                unsafe_allow_html=True,
            )
            with st.form("password_recovery_form"):
                st.text_input("Correo electrónico", value=recovery_email, disabled=True)
                new_password = st.text_input("Nueva contraseña", type="password")
                confirm_password = st.text_input("Confirmar contraseña", type="password")
                submitted = st.form_submit_button("Guardar contraseña", type="primary")
                if submitted:
                    if not new_password or not confirm_password:
                        st.error("Completa todos los campos.")
                    elif new_password != confirm_password:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        try:
                            _confirm_password_reset(active_token, new_password)
                            st.success("Tu contraseña se actualizó correctamente.")
                            st.session_state["password_recovery_mode"] = False
                            st.session_state.pop("recovery_email", None)
                            st.session_state.pop("active_reset_token", None)
                            st.rerun()
                        except Exception as err:
                            st.error(f"No se pudo actualizar la contraseña: {err}")
        if st.button("← Volver a iniciar sesión", key="btn_back_login_password_recovery"):
            st.session_state["password_recovery_mode"] = False
            st.session_state.pop("recovery_email", None)
            st.session_state.pop("active_reset_token", None)
            st.query_params.clear()
            st.rerun()


def _render_login():
    _handle_reset_query_token()
    st.session_state.setdefault("password_recovery_mode", False)
    st.session_state.setdefault("reset_request_mode", False)
    if st.session_state["password_recovery_mode"]:
        _render_password_recovery()
        return
    if st.session_state["reset_request_mode"]:
        _render_reset_request()
        return
    # Layout: hero como background full-viewport del .stApp. Titulo +
    # subtitulo se renderizan CENTRADOS arriba del card. Badge, form y
    # link "¿Olvidaste tu contraseña?" van DENTRO del card unificado
    # creado por `st.container()` + marker `.is-login-card` (CSS via
    # `:has()` lo trata como una sola caja).
    _inject_login_background_css()
    _render_auth_header(
        "Iniciar sesión",
        "Accede a tu Robot de auditoría del SRI",
    )
    card = st.container()
    with card:
        st.markdown('<div class="is-login-card"></div>', unsafe_allow_html=True)
        st.markdown(
            '<span class="badge-device">Sesión vinculada al dispositivo</span>',
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            email = st.text_input("Correo electrónico")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar sesión", type="primary")
        st.markdown(
            "<div class='auth-reset-wrap'>"
            "<a class='auth-reset-link' href='?reset_request=1'>&iquest;Olvidaste tu contrase&ntilde;a?</a>"
            "</div>",
            unsafe_allow_html=True,
        )
    if submitted:
        if not email or not password:
            st.error("Completa todos los campos.")
        else:
            try:
                email_clean = email.strip()
                cached_email = (st.session_state.get("user_email") or "").strip().lower()
                cached_ok = bool(
                    st.session_state.get("license_validated")
                    and st.session_state.get("device_fingerprint")
                    and cached_email == email_clean.lower()
                )
                token = LICENSE_CLIENT.login(email_clean, password)
                st.session_state["auth_token"] = token
                st.session_state["user_email"] = email_clean
                if cached_ok:
                    st.session_state["license_validated"] = True
                else:
                    st.session_state["license_validated"] = False
                    st.session_state.pop("license_last_check", None)
                _persist_session_state()
                if cached_ok:
                    st.success("Inicio de sesión exitoso.")
                else:
                    st.success("Inicio de sesión exitoso. Continúa con la activación.")
                st.rerun()
            except Exception as err:
                st.error(f"Error al autenticar: {err}")


def _render_activation():
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    _, center_col, _ = st.columns([1, 0.9, 1])
    with center_col:
        top_left_col, _ = st.columns([1.2, 4.8])
        with top_left_col:
            if st.button(" Volver a inicio de sesión", key="btn_back_login_activation"):
                device_id = st.session_state.get("_device_id") or _get_device_id_from_query()
                _clear_cached_auth_only()
                for key in ("auth_token", "user_email", "license_validated", "license_last_check"):
                    st.session_state.pop(key, None)
                if device_id:
                    st.session_state["_device_id"] = device_id
                st.rerun()

        logo_html = _logo_html(130)
        if logo_html:
            st.markdown(logo_html, unsafe_allow_html=True)
        st.markdown(
            "<h1 style='text-align:center; margin: 0.9rem 0;'>Activación de licencia</h1>",
            unsafe_allow_html=True,
        )
        st.warning("Introduce tu código de licencia para vincular este equipo.")
        client_device_id = _require_client_device_id()
        if not client_device_id:
            st.stop()
        default_fp = st.session_state.get("device_fingerprint") or hashlib.sha256(
            client_device_id.encode()
        ).hexdigest()
        st.session_state["device_fingerprint"] = default_fp
        with st.form("activation_form"):
            code = st.text_input("Código de licencia")
            fingerprint = st.text_input(
                "Identificador del equipo",
                value=default_fp,
                help="Este identificador se genera automáticamente para este equipo.",
                disabled=True,
            )
            submitted = st.form_submit_button("Activar licencia", type="primary")
            if submitted:
                if not code:
                    st.error("Debes ingresar tu código de licencia.")
                else:
                    try:
                        LICENSE_CLIENT.activate_license(
                            st.session_state["auth_token"],
                            code.strip(),
                            default_fp,
                        )
                        st.session_state["license_validated"] = True
                        st.session_state["license_last_check"] = time.time()
                        _persist_session_state()
                        st.success("Licencia activada correctamente.")
                        st.rerun()
                    except Exception as err:
                        st.error(f"No se pudo activar la licencia: {err}")
def _ensure_access():
    if "auth_token" not in st.session_state:
        client_device_id = _require_client_device_id()
        if not client_device_id:
            st.stop()
        _load_cached_session(client_device_id)
    if "auth_token" not in st.session_state:
        _render_login()
        st.stop()

    client_device_id = _require_client_device_id()
    if not client_device_id:
        st.stop()
    fingerprint = st.session_state.get("device_fingerprint") or hashlib.sha256(
        client_device_id.encode()
    ).hexdigest()
    st.session_state["device_fingerprint"] = fingerprint

    # Validar siempre contra el backend para que la cuenta dependa de Render.
    try:
        LICENSE_CLIENT.get_profile(st.session_state["auth_token"])
    except Exception:
        _clear_cached_auth_only()
        for key in ("auth_token", "user_email", "license_validated", "license_last_check"):
            st.session_state.pop(key, None)
        st.warning("Tu sesión ya no es válida. Inicia sesión nuevamente.")
        _render_login()
        st.stop()

    # Revalidar licencia en cada acceso (sin bypass local por cache).
    try:
        LICENSE_CLIENT.validate_license(st.session_state["auth_token"], fingerprint)
        st.session_state["license_validated"] = True
        st.session_state["license_last_check"] = time.time()
        _persist_session_state()
    except Exception:
        st.session_state["license_validated"] = False
        st.session_state.pop("license_last_check", None)
        _persist_session_state()
        _render_activation()
        st.stop()


_ensure_access()
_load_user_preferences()
if "first_use_tour_active" not in st.session_state:
    st.session_state["first_use_tour_active"] = False
if "first_use_tour_step" not in st.session_state:
    st.session_state["first_use_tour_step"] = 0
if "first_use_tour_prompt_shown_session" not in st.session_state:
    st.session_state["first_use_tour_prompt_shown_session"] = False
DEVICE_FINGERPRINT = st.session_state.get("device_fingerprint") or st.session_state.get("user_email")

# ==============================
# TOPBAR (sticky header del mockup)
# ==============================
# Versión dinámica del app — en el .exe compilado viene de version.txt;
# en modo dev mostramos un placeholder. La usamos en el topbar (chip
# discreto) y tambien en Ayuda → Acerca de la aplicación.
if _desktop_launcher is not None and _desktop_launcher.APP_VERSION:
    _app_version_display = _desktop_launcher.APP_VERSION
else:
    _app_version_display = "3.0 (dev)"


def _render_topbar(app_version: str) -> None:
    """Topbar sticky pegado arriba con brand a la izquierda, titulo al
    centro, version + theme toggle + profile popover a la derecha.

    El theme toggle y el menu de perfil (Cerrar sesion / Cerrar app)
    viven todos DENTRO del topbar — no hay botones flotantes sueltos
    en otras partes de la pantalla.
    """
    user_email = st.session_state.get("user_email") or "No disponible"
    initials_seed = (user_email.split("@")[0] or "U").replace(".", " ")
    avatar_initials = "".join(part[0].upper() for part in initials_seed.split() if part)[:2] or "U"

    # Logo segun el tema activo: en modo claro usamos la variante blanca
    # (logo-acg-white.png) para que el isotipo siga siendo legible sobre
    # el fondo claro. En modo oscuro se mantiene la version original.
    _tema_actual = st.session_state.get("ui_theme", "dark")
    _logo_filename = "logo-acg-white.png" if _tema_actual == "light" else "logo-acg.png"
    logo_uri = _asset_data_uri(_logo_filename)
    logo_img = (
        f'<img src="{logo_uri}" alt="Audit Consulting Group"/>'
        if logo_uri
        else ""
    )
    is_frozen = getattr(sys, "frozen", False)

    # Container con key para que el CSS le aplique aspecto de topbar
    # (sticky, glass background, border-bottom). Streamlit emite la clase
    # `st-key-topbar_container` sobre el wrapper, que es lo que apunta el CSS.
    # Adentro: 5 columnas alineadas verticalmente al centro.
    try:
        bar = st.container(key="topbar_container")
    except TypeError:
        # `key` en st.container llego en Streamlit 1.34; fallback con marker.
        bar = st.container()
    with bar:
        st.markdown('<div class="is-topbar-marker"></div>', unsafe_allow_html=True)
        # Pesos: brand izquierda + spacer central (titulo es absolute, no
        # ocupa flujo) + boton tema + icono usuario. La version se removio
        # del topbar — sigue disponible en Ayuda > Acerca de la aplicacion.
        col_weights = [1.2, 7.6, 0.6, 1.0]
        try:
            cols = st.columns(col_weights, vertical_alignment="center")
        except TypeError:
            # vertical_alignment llego en Streamlit 1.36; fallback sin el kwarg.
            cols = st.columns(col_weights)

        with cols[0]:
            st.markdown(
                f"""
                <div class="brand">
                  {logo_img}
                </div>
                """,
                unsafe_allow_html=True,
            )
        with cols[1]:
            st.markdown(
                '<div class="topbar-title">SRI Robot Audit |<b> Descarga y Reporte Automático</b></div>',
                unsafe_allow_html=True,
            )
        with cols[2]:
            # Theme toggle DENTRO del topbar (reemplaza al boton flotante
            # _render_theme_toggle que vivia arriba a la derecha).
            tema = st.session_state.get("ui_theme", "dark")
            # Solo icono (sin texto "Claro"/"Oscuro") para mantener el
            # boton compacto cuadrado de 44px como el popover de usuario.
            label = "☀️" if tema == "dark" else "🌙"
            if st.button(label, key="btn_topbar_theme", help="Cambiar tema claro/oscuro"):
                st.session_state["ui_theme"] = "light" if tema == "dark" else "dark"
                st.rerun()
        with cols[3]:
            # Menu de usuario: trigger compacto con silueta de persona
            # (👤) en vez de hamburguesa. El email del usuario y el
            # boton Cerrar sesion viven ADENTRO del panel desplegable
            # — NO en el topbar como chip permanente.
            with st.popover("👤 Perfil", use_container_width=True, help="Menú de usuario"):
                # Email del usuario adentro del panel
                st.markdown(
                    f"""
                    <div class="user-menu-row">
                      <span class="user-menu-avatar">{html.escape(avatar_initials)}</span>
                      <div class="user-menu-info">
                        <span class="user-menu-label">Sesión activa</span>
                        <span class="user-menu-email">{html.escape(user_email)}</span>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown('<hr class="user-menu-divider"/>', unsafe_allow_html=True)
                # Boton Cerrar sesion — mismo handler que tenia antes (no toco la
                # logica de auth: solo se mueve de ubicacion visual).
                if st.button(
                    "🚪  Cerrar sesión",
                    key="btn_popover_logout",
                    use_container_width=True,
                ):
                    device_id = st.session_state.get("_device_id") or _get_device_id_from_query()
                    _clear_cached_auth_only()
                    st.session_state.clear()
                    if device_id:
                        st.session_state["_device_id"] = device_id
                    st.rerun()
                if is_frozen:
                    if st.button(
                        "⏻  Cerrar app",
                        key="btn_popover_close_app",
                        use_container_width=True,
                    ):
                        st.session_state["open_close_app_dialog"] = True

    # Spacer renderizado ABAJO del topbar (que ahora es position:fixed)
    # para empujar el contenido real (tabs, formularios) y que no quede
    # tapado. Su altura matchea el `height` del topbar en el CSS.
    st.markdown('<div class="topbar-spacer" aria-hidden="true"></div>', unsafe_allow_html=True)


_render_topbar(_app_version_display)


# ==============================
# INTERFAZ PRINCIPAL
# ==============================

if hasattr(st, "dialog"):
    @st.dialog("Recorrido rapido del sistema")
    def _tour_dialog():
        _render_first_use_tour()

    @st.dialog("Bienvenido")
    def _tour_prompt_dialog():
        _render_first_use_prompt()

    @st.dialog("Proceso terminado")
    def _download_finished_dialog():
        _render_download_finished_modal()

    @st.dialog("Cerrar aplicacion")
    def _close_app_dialog():
        _render_close_app_modal()
else:
    def _tour_dialog():
        _render_first_use_tour()

    def _tour_prompt_dialog():
        _render_first_use_prompt()

    def _download_finished_dialog():
        _render_download_finished_modal()

    def _close_app_dialog():
        _render_close_app_modal()

if st.session_state.pop("open_close_app_dialog", False):
    _close_app_dialog()

if (
    not st.session_state.get("first_use_tour_completed", False)
    and not st.session_state.get("first_use_tour_prompt_dismissed", False)
    and not st.session_state.get("first_use_tour_active", False)
    and not st.session_state.get("first_use_tour_prompt_shown_session", False)
):
    st.session_state["first_use_tour_prompt_shown_session"] = True
    _tour_prompt_dialog()

if st.session_state.get("first_use_tour_active", False):
    _tour_dialog()

if st.session_state.get("download_finished_modal_open", False):
    _download_finished_dialog()

tab1, tab2, tab3, tab4 = st.tabs(
    [" Descarga de Comprobantes", " Reportes e Historial", " Consolidacion de documentos", " Ayuda"]
)

# =====================================================
# TAB 1  DESCARGA Y PROCESAMIENTO AUTOMTICO
# =====================================================
with tab1:
    pending_download_checkpoint = None
    if st.session_state.download_status not in {"running", "cancelling"}:
        pending_download_checkpoint = _load_pending_download_checkpoint_for_current_user()
    if pending_download_checkpoint:
        resume_summary = pending_download_checkpoint.get("summary") if isinstance(pending_download_checkpoint.get("summary"), dict) else {}
        resume_progress = pending_download_checkpoint.get("progress") if isinstance(pending_download_checkpoint.get("progress"), dict) else {}
        resume_period = _format_download_resume_period(resume_summary)
        resume_last_point = str(resume_progress.get("last_completed_label") or "").strip()
        resume_origen = str(resume_summary.get("origen") or "No disponible")
        resume_tipo = str(resume_summary.get("tipo") or "No disponible")
        resume_formatos = ", ".join(resume_summary.get("formatos") or []) or "No disponible"
        st.info(
            f"Descarga pendiente detectada. Origen: {resume_origen}. Tipo: {resume_tipo}. "
            f"Periodo: {resume_period}. Ultimo punto guardado: {resume_last_point or 'inicio del proceso'}."
        )
        st.caption(f"Formatos: {resume_formatos}")
        resume_error = str(pending_download_checkpoint.get("last_error") or "").strip()
        if resume_error:
            st.caption(f"Ultimo error registrado: {resume_error}")
        col_resume_1, col_resume_2 = st.columns([1, 1])
        with col_resume_1:
            if st.button("Reanudar descarga", key="btn_resume_download", use_container_width=True, type="secondary"):
                resume_params = deserialize_download_params(pending_download_checkpoint.get("params"))
                if not resume_params:
                    st.error("No se pudo recuperar la configuración de la descarga pendiente.")
                else:
                    _start_download_process(resume_params, resume_download=True)
                    st.rerun()
        with col_resume_2:
            if st.button("Descartar", key="btn_discard_resume_download", use_container_width=True):
                delete_download_checkpoint(pending_download_checkpoint.get("_path"))
                st.rerun()

    col_title, col_tour_link = st.columns([5, 1.6])
    with col_title:
        st.markdown(
            "<div class='pagehead'>"
            "<h1>Descarga de Comprobantes</h1>"
            "<p>Ingreso de credenciales y filtros — el robot descarga todo automáticamente.</p>"
            "</div>",
            unsafe_allow_html=True,
        )
    with col_tour_link:
        if st.button("Primera vez? Ver tour", key="btn_open_tour"):
            _start_first_use_tour(reset_step=True)
            st.rerun()

    with _group_card(1, "Credenciales", "Datos del SRI"):
        col_ruc, col_clave = st.columns([1, 1])
        with col_ruc:
            ruc_input = st.text_input("RUC", placeholder="Ejemplo: 0999999001")
            ruc = re.sub(r"\s+", "", ruc_input or "").strip()
            ci_adicional_input = ""
        with col_clave:
            clave = st.text_input("Clave del SRI", type="password", placeholder="********")

    with _group_card(2, "Filtros", "Qué descargar"):
        col_origen, col_tipo = st.columns([1, 1])
        with col_origen:
            origen = st.selectbox("Origen de comprobantes", ["Recibidos", "Emitidos"], index=0)
            if origen == "Recibidos":
                tipo_opciones = [
                    "Facturas",
                    "Retenciones",
                    "Notas de crédito",
                    "Notas de débito",
                    "Liquidación de compra",
                ]
            else:
                tipo_opciones = [
                    "Facturas",
                    "Liquidación de compra",
                    "Retenciones",
                    "Notas de crédito",
                    "Notas de débito",
                    "Guía de remisión",
                ]
        with col_tipo:
            tipo = st.selectbox("Tipo de comprobante", tipo_opciones)

        estado_emitidos = None
        establecimiento_input = None
        punto_emision_input = None
        formatos = []
        descargar_pdf_emitidos = False
        descargar_xml_emitidos = False
        anio_emitidos = datetime.now().year
        mes_emitidos = datetime.now().month
        dia_emitidos = datetime.now().day
        mes_fin_emitidos = datetime.now().month
        anio_recibidos = datetime.now().year
        mes_recibidos = datetime.now().month
        dia_recibidos = 0
        mes_fin_recibidos = datetime.now().month
        meses_es = [
            "Enero",
            "Febrero",
            "Marzo",
            "Abril",
            "Mayo",
            "Junio",
            "Julio",
            "Agosto",
            "Septiembre",
            "Octubre",
            "Noviembre",
            "Diciembre",
        ]
        modo_fechas_recibidos = "Mes y dí­a"
        modo_fechas_emitidos = "Mes y día"

        if origen == "Recibidos":
            modo_fechas_recibidos = st.radio(
                "Modo de fecha",
                ["Mes y día", "Rango de meses", "Año completo"],
                horizontal=True,
                key="modo_fechas_recibidos",
            )
            if modo_fechas_recibidos == "Rango de meses":
                col_r1, col_r2, col_r3 = st.columns([1, 1, 1])
                with col_r1:
                    anio_recibidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                with col_r2:
                    mes_inicio_label = st.selectbox(
                        "Mes inicio",
                        meses_es,
                        index=mes_recibidos - 1,
                        key="mes_inicio_recibidos",
                    )
                with col_r3:
                    mes_fin_label = st.selectbox(
                        "Mes fin",
                        meses_es,
                        index=mes_fin_recibidos - 1,
                        key="mes_fin_recibidos",
                    )
                mes_recibidos = meses_es.index(mes_inicio_label) + 1
                mes_fin_recibidos = meses_es.index(mes_fin_label) + 1
                dia_recibidos = 0
            elif modo_fechas_recibidos == "Año completo":
                col_r1, col_r2, col_r3 = st.columns([1, 1, 1])
                with col_r1:
                    anio_recibidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                mes_recibidos = 1
                mes_fin_recibidos = 12
                dia_recibidos = 0
            else:
                col_r1, col_r2, col_r3 = st.columns([1, 1, 1])
                with col_r1:
                    anio_recibidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                with col_r2:
                    mes_label = st.selectbox(
                        "Mes",
                        meses_es,
                        index=mes_recibidos - 1,
                        key="mes_recibidos",
                    )
                mes_recibidos = meses_es.index(mes_label) + 1
                with col_r3:
                    dia_recibidos = st.number_input(
                        "Día (0 = Todos)", min_value=0, max_value=31, value=0, step=1,
                        help="Elige 0 para descargar todo el mes o un día específico (1-31).",
                    )
            formatos = st.multiselect("Formatos a descargar", ["XML", "PDF"], default=["XML", "PDF"])
        else:
            modo_fechas_emitidos = st.radio(
                "Modo de fecha",
                ["Mes y día", "Rango de meses", "Año completo"],
                horizontal=True,
                key="modo_fechas_emitidos",
            )
            if modo_fechas_emitidos == "Rango de meses":
                col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
                with col_f1:
                    anio_emitidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                with col_f2:
                    mes_inicio_label = st.selectbox(
                        "Mes inicio",
                        meses_es,
                        index=mes_emitidos - 1,
                        key="mes_inicio_emitidos",
                    )
                with col_f3:
                    mes_fin_label = st.selectbox(
                        "Mes fin",
                        meses_es,
                        index=mes_fin_emitidos - 1,
                        key="mes_fin_emitidos",
                    )
                mes_emitidos = meses_es.index(mes_inicio_label) + 1
                mes_fin_emitidos = meses_es.index(mes_fin_label) + 1
                dia_emitidos = 0
            elif modo_fechas_emitidos == "Año completo":
                col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
                with col_f1:
                    anio_emitidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                mes_emitidos = 1
                mes_fin_emitidos = 12
                dia_emitidos = 0
            else:
                col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
                with col_f1:
                    anio_emitidos = st.number_input(
                        "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
                    )
                with col_f2:
                    mes_label = st.selectbox(
                        "Mes",
                        meses_es,
                        index=mes_emitidos - 1,
                        key="mes_emitidos",
                    )
                mes_emitidos = meses_es.index(mes_label) + 1
                with col_f3:
                    dia_emitidos = st.number_input(
                        "Día (0 = Todos)",
                        min_value=0,
                        max_value=31,
                        value=datetime.now().day,
                        step=1,
                        help="Ingresa 0 para descargar todos los días del mes.",
                    )
            estado_emitidos = st.selectbox(
                "Estado de autorización", ["Autorizados", "No autorizados"], index=0
            )
            col_e1, col_e2 = st.columns([1, 1])
            with col_e1:
                establecimiento_input = st.text_input(
                    "Establecimiento",
                    value="Todos",
                    help="Escribe 'Todos' o un número de 3 dígitos (ej. 001).",
                )
            with col_e2:
                punto_emision_input = st.text_input(
                    "Punto de emisión (opcional)",
                    value="",
                    max_chars=3,
                    key="punto_emision_input",
                    help="Hasta 3 dígitos numéricos (ej. 001).",
                )
                if punto_emision_input:
                    solo_digitos = "".join(ch for ch in punto_emision_input if ch.isdigit())
                    if solo_digitos != punto_emision_input:
                        st.session_state.punto_emision_input = solo_digitos
                        punto_emision_input = solo_digitos
            descargar_pdf_emitidos = st.checkbox(
                'Descargar PDFs individuales',
                value=False,
                help='Genera un PDF por cada comprobante emitido.',
            )
            descargar_xml_emitidos = st.checkbox(
                'Descargar XMLs individuales',
                value=False,
                help='Extrae el comprobante XML autorizado y lo organiza en la carpeta XML.',
            )
            formatos = ["Excel"]
            if descargar_pdf_emitidos:
                formatos.append("PDF")
            if descargar_xml_emitidos:
                formatos.append("XML")

    with _group_card(3, "Carpeta base", "Dónde se guardan las descargas"):
        current_dir = st.session_state.get("download_base_dir", str(DESC_DIR))
        st.text_input(
            "Ruta seleccionada",
            value=current_dir,
            help="Ej: C:\\\\RespaldosSRI o /home/usuario/SRI.",
            disabled=True,
        )
        if st.button("Seleccionar carpeta de descarga"):
            seleccionada, error = _select_directory_dialog(current_dir)
            if seleccionada:
                try:
                    nueva_ruta = Path(seleccionada).expanduser()
                    nueva_ruta.mkdir(parents=True, exist_ok=True)
                    st.session_state["download_base_dir"] = str(nueva_ruta)
                    _persist_user_preferences()
                    st.success(f" Carpeta configurada: {nueva_ruta}")
                except Exception as err:
                    st.error(f"No se pudo usar la carpeta indicada: {err}")
            elif error:
                st.session_state["show_manual_dir"] = True
                st.session_state["last_manual_dir_error"] = error

        if st.session_state.get("show_manual_dir"):
            manual_default = st.session_state.get("download_base_dir", str(DESC_DIR))
            manual_dir = st.text_input("Ruta de carpeta (manual)", value=manual_default)
            if st.button("Guardar carpeta"):
                try:
                    nueva_ruta = Path(manual_dir).expanduser()
                    nueva_ruta.mkdir(parents=True, exist_ok=True)
                    st.session_state["download_base_dir"] = str(nueva_ruta)
                    _persist_user_preferences()
                    st.success(f" Carpeta configurada: {nueva_ruta}")
                    st.session_state["show_manual_dir"] = False
                    st.session_state["last_manual_dir_error"] = None
                except Exception as err:
                    st.error(f"No se pudo usar la carpeta indicada: {err}")
            last_err = st.session_state.get("last_manual_dir_error")
            if last_err:
                if "tk" in str(last_err).lower() or "libtk" in str(last_err).lower():
                    st.info("El selector nativo no está disponible en este entorno. Usa la ruta manual.")
                else:
                    st.warning(last_err)
        st.caption(
            f"Carpeta activa: `{st.session_state.get('download_base_dir', str(DESC_DIR))}`. Dentro se almacenarán tus descargas."
        )

    with _group_card(4, "Ejecutar"):
        start_clicked = st.button("▶️  Iniciar proceso", use_container_width=True, type="primary", key="start_process")
        stop_clicked = st.button("⏹️  Detener proceso", use_container_width=True, key="stop_process")

    if stop_clicked and st.session_state.download_status in {"running", "cancelling"}:
        from robot.downloader import request_cancel
        request_cancel()
        st.session_state.download_status = "cancelling"
        st.session_state.stop_notice_ts = time.time()
        st.session_state.running_notice_ts = None

    if start_clicked and st.session_state.download_status != "running":
        if not ruc or not clave:
            st.warning(" Ingresa RUC y clave antes de continuar.")
        else:
            mes_fin_val = None
            if origen == "Recibidos":
                formatos_final = formatos
                if not formatos_final:
                    st.warning("Selecciona al menos un formato (XML o PDF).")
                    st.stop()
                anio_val = int(anio_recibidos)
                mes_val = int(mes_recibidos)
                dia_val = int(dia_recibidos)
                if modo_fechas_recibidos == "Rango de meses":
                    mes_fin_val = int(mes_fin_recibidos)
                    if mes_fin_val < mes_val:
                        st.error("El mes fin debe ser mayor o igual al mes inicio.")
                        st.stop()
                    dia_val = 0
                elif modo_fechas_recibidos == "Año completo":
                    mes_val = 1
                    mes_fin_val = 12
                    dia_val = 0
                fecha_emitidos_val = None
                estado_emitidos_val = None
                establecimiento_val = None
                punto_emision_val = None
            else:
                anio_val = int(anio_emitidos)
                mes_val = int(mes_emitidos)
                dia_val = int(dia_emitidos)
                if modo_fechas_emitidos == "Rango de meses":
                    mes_fin_val = int(mes_fin_emitidos)
                    if mes_fin_val < mes_val:
                        st.error("El mes fin debe ser mayor o igual al mes inicio.")
                        st.stop()
                    dia_val = 0
                elif modo_fechas_emitidos == "Año completo":
                    mes_val = 1
                    mes_fin_val = 12
                    dia_val = 0
                if modo_fechas_emitidos == "Mes y día":
                    dias_en_mes = calendar.monthrange(anio_val, mes_val)[1]
                    if dia_val > dias_en_mes:
                        st.error(f"El día debe estar entre 1 y {dias_en_mes}, o 0 para todos.")
                        st.stop()
                fecha_emitidos_val = None if dia_val == 0 else f"{dia_val:02d}/{mes_val:02d}/{anio_val}"
                estado_emitidos_val = estado_emitidos
                est_clean = (establecimiento_input or "").strip()
                if not est_clean or est_clean.lower() == "todos":
                    establecimiento_val = "Todos"
                else:
                    if not (est_clean.isdigit() and len(est_clean) == 3):
                        st.error("El establecimiento debe ser 'Todos' o un número de tres dígitos (ej. 001).")
                        st.stop()
                    establecimiento_val = est_clean
                punto_clean = (punto_emision_input or "").strip()
                if punto_clean:
                    punto_emision_val = punto_clean
                else:
                    punto_emision_val = ""
                formatos_final = formatos
                if not any(fmt in formatos_final for fmt in ("PDF", "XML")):
                    st.warning("Selecciona al menos un formato (PDF o XML).")
                    st.stop()
                estado_emitidos_norm = (
                    unicodedata.normalize("NFKD", estado_emitidos_val or "")
                    .encode("ascii", "ignore")
                    .decode("ascii")
                    .strip()
                    .lower()
                )
                es_emitidos_autorizados = estado_emitidos_norm == "autorizados"
                selecciono_xml_emitidos = "XML" in formatos_final
                selecciono_pdf_emitidos = "PDF" in formatos_final
                if es_emitidos_autorizados and selecciono_xml_emitidos:
                    if modo_fechas_emitidos == "Mes y día":
                        if dia_val in (0, None):
                            fecha_inicio_sel = date(anio_val, mes_val, 1)
                            fecha_fin_sel = date(anio_val, mes_val, calendar.monthrange(anio_val, mes_val)[1])
                            periodo_txt = f"{meses_es[mes_val - 1]} {anio_val}"
                        else:
                            fecha_inicio_sel = date(anio_val, mes_val, int(dia_val))
                            fecha_fin_sel = fecha_inicio_sel
                            periodo_txt = fecha_inicio_sel.strftime("%d/%m/%Y")
                    elif modo_fechas_emitidos == "Rango de meses":
                        mes_fin_real = int(mes_fin_val or mes_val)
                        fecha_inicio_sel = date(anio_val, mes_val, 1)
                        fecha_fin_sel = date(anio_val, mes_fin_real, calendar.monthrange(anio_val, mes_fin_real)[1])
                        periodo_txt = f"{meses_es[mes_val - 1]} {anio_val} a {meses_es[mes_fin_real - 1]} {anio_val}"
                    else:
                        fecha_inicio_sel = date(anio_val, 1, 1)
                        fecha_fin_sel = date(anio_val, 12, 31)
                        periodo_txt = f"Ano completo {anio_val}"

                    limite_xml_emitidos = date.today() - timedelta(days=30)
                    if fecha_fin_sel < limite_xml_emitidos:
                        aviso_fecha_xml = (
                            f"Advertencia de fecha: para XML Emitidos, el periodo {periodo_txt} esta fuera del limite de 30 dias "
                            f"(antes del {limite_xml_emitidos.strftime('%d/%m/%Y')})."
                        )
                        if selecciono_pdf_emitidos:
                            formatos_final = [fmt for fmt in formatos_final if fmt != "XML"]
                            st.warning(aviso_fecha_xml + " Se continuará automaticamente solo con PDF.")
                        else:
                            st.warning(aviso_fecha_xml + " Selecciona PDF o usa una fecha dentro de los ultimos 30 dias.")
                            st.stop()

            base_descargas = _get_download_base_dir()
            destino = base_descargas / ruc
            destino.mkdir(parents=True, exist_ok=True)

            params = {
                "ruc": ruc,
                "clave": clave,
                "anio": anio_val,
                "mes": mes_val,
                "mes_fin": mes_fin_val,
                "dia": int(dia_val),
                "tipo": tipo,
                "formatos": formatos_final,
                "destino": destino,
                "origen": origen,
                "ci_adicional": ci_adicional_input.strip() or None,
                "fecha_emitidos": fecha_emitidos_val,
                "estado_emitidos": estado_emitidos_val,
                "establecimiento": establecimiento_val,
                "punto_emision": punto_emision_val,
            }
            checkpoint_payload = build_checkpoint_payload(st.session_state.get("user_email"), params)
            _start_download_process(params, resume_download=False, checkpoint_payload=checkpoint_payload)
            st.rerun()

    if st.session_state.download_status not in {"running", "cancelling"}:
        st.session_state.running_notice_ts = None
        st.session_state.stop_notice_ts = None

    if st.session_state.download_status == "cancelling":
        ahora = time.time()
        if st.session_state.stop_notice_ts and (ahora - st.session_state.stop_notice_ts) <= 10:
            st.warning("Cancelando proceso. Espera a que se cierre el navegador...")
        # Si el hilo ya terminó y no llegó mensaje, marcar como cancelado
        hilo = st.session_state.download_thread
        if hilo and not hilo.is_alive() and not st.session_state.download_error and not st.session_state.download_result:
            st.session_state.download_error = "Proceso cancelado por el usuario."
            st.session_state.download_status = "error"
    elif st.session_state.download_status == "running":
        ahora = time.time()
        if st.session_state.stop_notice_ts and (ahora - st.session_state.stop_notice_ts) <= 10:
            st.warning("Solicitud de detener registrada. Esperando a que el proceso termine...")
        if st.session_state.running_notice_ts and (ahora - st.session_state.running_notice_ts) <= 10:
            st.info("Proceso en ejecución. Puedes detenerlo con el botón rojo.")
        if st.session_state.last_download_message:
            msg, ts = st.session_state.last_download_message
            if (ahora - ts) <= 10:
                st.warning(msg)
    _render_manual_consultar_modal()

    if st.session_state.download_status in {"done", "error"} and st.session_state.download_params:
        st.session_state.running_notice_ts = None
        st.session_state.stop_notice_ts = None
        resultado = st.session_state.download_result or {}
        params = st.session_state.download_params or {}
        if st.session_state.download_error:
            raw_error = str(st.session_state.download_error)
            level, user_msg = _friendly_download_error_message(raw_error, params.get("origen"))
            if level == "warning":
                st.warning(user_msg)
            else:
                st.error(user_msg)
            if user_msg.strip() != raw_error.strip():
                st.caption(f"Detalle tecnico: {raw_error}")
        if resultado and not st.session_state.download_registered:
            dia_registro = params.get("dia")
            registrar_descarga(
                params.get("ruc"),
                params.get("origen"),
                params.get("anio"),
                params.get("mes"),
                dia_registro,
                params.get("tipo"),
                resultado,
                device_id=DEVICE_FINGERPRINT,
            )
            st.session_state.download_registered = True

        if resultado:
            estado = resultado.get("estado", "")
            carpeta_tipo = Path(resultado.get("carpeta_tipo") or params.get("destino"))
            aviso_recorte = resultado.get("aviso_recorte")
            mensaje_verificacion = str(resultado.get("mensaje_verificacion") or "").strip()
            if aviso_recorte:
                st.warning(aviso_recorte)
            if mensaje_verificacion:
                if resultado.get("descarga_completa", True):
                    st.caption(mensaje_verificacion)
                else:
                    st.warning(f"Verificación de descarga: {mensaje_verificacion}")
            if estado in {"sin_descargas", "sin_resultados"}:
                st.warning(" No se encontraron comprobantes para el período seleccionado.")
            elif params.get("origen") == "Emitidos":
                n_regs = resultado.get("n_registros", 0)
                st.success(f' Reporte de emitidos generado con {n_regs} registros.')
                if "PDF" in (params.get("formatos") or []):
                    n_pdf_emitidos = resultado.get('n_pdf', 0)
                    pdf_dir_emitidos = resultado.get('pdf_dir')
                    st.caption(f'PDFs descargados: {n_pdf_emitidos}')
                    if pdf_dir_emitidos:
                        st.caption(f'Carpeta de PDFs: `{Path(pdf_dir_emitidos)}`')
                if "XML" in (params.get("formatos") or []):
                    n_xml_emitidos = resultado.get('n_xml', 0)
                    xml_dir_emitidos = resultado.get('xml_dir')
                    st.caption(f'XMLs descargados: {n_xml_emitidos}')
                    if xml_dir_emitidos:
                        st.caption(f'Carpeta de XMLs: `{Path(xml_dir_emitidos)}`')
                filtros = []
                if resultado.get("fecha_filtro"):
                    filtros.append(f"Fecha: {resultado['fecha_filtro']}")
                if resultado.get("estado_autorizacion"):
                    filtros.append(f"Estado: {resultado['estado_autorizacion']}")
                if resultado.get("establecimiento"):
                    filtros.append(f"Establecimiento: {resultado['establecimiento']}")
                if resultado.get("punto_emision"):
                    filtros.append(f"Punto: {resultado['punto_emision']}")
                if filtros:
                    st.caption(" | ".join(filtros))
                st.caption(f"Archivos organizados en: `{carpeta_tipo}`")
                reporte_path = resultado.get("reporte")
                if reporte_path and Path(reporte_path).exists():
                    with open(reporte_path, "rb") as f:
                        st.download_button(
                            " Descargar reporte Excel (Emitidos)",
                            f,
                            file_name=Path(reporte_path).name,
                            use_container_width=True,
                        )
                reporte_pdf_path = resultado.get("reporte_pdf")
                if reporte_pdf_path and Path(reporte_pdf_path).exists():
                    with open(reporte_pdf_path, "rb") as f:
                        st.download_button(
                            " Descargar reporte PDF (Emitidos)",
                            f,
                            file_name=Path(reporte_pdf_path).name,
                            use_container_width=True,
                        )
                reporte_xml_path = resultado.get("reporte_xml")
                if reporte_xml_path and Path(reporte_xml_path).exists():
                    with open(reporte_xml_path, "rb") as f:
                        st.download_button(
                            " Descargar reporte XML (Emitidos)",
                            f,
                            file_name=Path(reporte_xml_path).name,
                            use_container_width=True,
                        )
                if resultado.get("rango_meses"):
                    st.info("Se generaron reportes por cada mes en las carpetas correspondientes.")
                    reporte_xml_rango = resultado.get("reporte_xml_rango")
                    if reporte_xml_rango and Path(reporte_xml_rango).exists():
                        with open(reporte_xml_rango, "rb") as f:
                            st.download_button(
                                " Descargar reporte XML del rango (Emitidos)",
                                f,
                                file_name=Path(reporte_xml_rango).name,
                                use_container_width=True,
                            )
                    reporte_pdf_anual = resultado.get("reporte_pdf_anual")
                    if reporte_pdf_anual and Path(reporte_pdf_anual).exists():
                        with open(reporte_pdf_anual, "rb") as f:
                            st.download_button(
                                " Descargar reporte PDF anual (Emitidos)",
                                f,
                                file_name=Path(reporte_pdf_anual).name,
                                use_container_width=True,
                            )
                    reporte_xml_anual = resultado.get("reporte_xml_anual")
                    if reporte_xml_anual and Path(reporte_xml_anual).exists():
                        with open(reporte_xml_anual, "rb") as f:
                            st.download_button(
                                " Descargar reporte XML anual (Emitidos)",
                                f,
                                file_name=Path(reporte_xml_anual).name,
                                use_container_width=True,
                            )
            else:
                n_xml = resultado.get("n_xml", 0)
                n_pdf = resultado.get("n_pdf", 0)
                tipo_visible = resultado.get("tipo_visible", params.get("tipo"))
                st.success(f" Descarga completada ({tipo_visible}). XML: {n_xml} | PDF: {n_pdf}")
                st.caption(f"Archivos organizados en: `{carpeta_tipo}`")

                txt_path = resultado.get("txt")
                if txt_path and Path(txt_path).exists():
                    with open(txt_path, "rb") as f:
                        st.download_button(
                            " Descargar TXT semilla",
                            f,
                            file_name=Path(txt_path).name,
                            use_container_width=True,
                        )
                if resultado.get("rango_meses"):
                    st.info("Se generaron reportes por cada mes en las carpetas correspondientes.")
                    reporte_xml_rango = resultado.get("reporte_xml_rango")
                    if reporte_xml_rango and Path(reporte_xml_rango).exists():
                        with open(reporte_xml_rango, "rb") as f:
                            st.download_button(
                                " Descargar reporte XML del rango (Recibidos)",
                                f,
                                file_name=Path(reporte_xml_rango).name,
                                use_container_width=True,
                            )
                    reporte_pdf_anual = resultado.get("reporte_pdf_anual")
                    if reporte_pdf_anual and Path(reporte_pdf_anual).exists():
                        with open(reporte_pdf_anual, "rb") as f:
                            st.download_button(
                                " Descargar reporte PDF anual (Recibidos)",
                                f,
                                file_name=Path(reporte_pdf_anual).name,
                                use_container_width=True,
                            )
                    reporte_xml_anual = resultado.get("reporte_xml_anual")
                    if reporte_xml_anual and Path(reporte_xml_anual).exists():
                        with open(reporte_xml_anual, "rb") as f:
                            st.download_button(
                                " Descargar reporte XML anual (Recibidos)",
                                f,
                                file_name=Path(reporte_xml_anual).name,
                                use_container_width=True,
                            )
                else:
                    if n_xml > 0:
                        reporte_xml_path = resultado.get("reporte_xml")
                        if reporte_xml_path and Path(reporte_xml_path).exists():
                            with open(reporte_xml_path, "rb") as f:
                                st.download_button(
                                    " Descargar reporte Excel (Recibidos)",
                                    f,
                                    file_name=Path(reporte_xml_path).name,
                                    use_container_width=True,
                                )
                        else:
                            xml_folder = Path(resultado.get("xml_dir") or (carpeta_tipo / "XML"))
                            tipo_param = params.get("tipo") or ""
                            tipo_slug = resultado.get("tipo_slug", tipo_param.lower().replace(" ", "_"))
                            anio_param = params.get("anio") or 0
                            mes_param = params.get("mes") or 0
                            try:
                                mes_int = int(mes_param)
                            except Exception:
                                mes_int = 0
                            excel_path = carpeta_tipo / f"reporte_{tipo_slug}_{anio_param}_{mes_int:02d}.xlsx"
                            construir_reporte(xml_folder, excel_path)
                            if not excel_path.exists():
                                # Fallback para estructura por dias (base del mes).
                                xml_base_mes = Path(resultado.get("xml_dir") or carpeta_tipo)
                                tipo_visible_tmp = resultado.get("tipo_visible", tipo_param)
                                _, _, tipo_prefijo_tmp = _prefijo_tipo(tipo_visible_tmp)
                                xml_files_tmp = _xml_files_por_tipo(xml_base_mes, tipo_prefijo_tmp)
                                if xml_files_tmp:
                                    construir_reporte(xml_base_mes, excel_path, None, xml_files=xml_files_tmp)
                            if excel_path.exists():
                                with open(excel_path, "rb") as f:
                                    st.download_button(
                                        " Descargar reporte Excel (Recibidos)",
                                        f,
                                        file_name=excel_path.name,
                                        use_container_width=True,
                                    )

                    reporte_pdf_path = resultado.get("reporte_pdf")
                    if reporte_pdf_path and Path(reporte_pdf_path).exists():
                        with open(reporte_pdf_path, "rb") as f:
                            st.download_button(
                                " Descargar reporte PDF (Recibidos)",
                                f,
                                file_name=Path(reporte_pdf_path).name,
                                use_container_width=True,
                            )

            if estado not in {"sin_descargas", "sin_resultados"} and carpeta_tipo:
                zip_target = carpeta_tipo
                zip_path = zip_target.with_suffix(".zip")
                if zip_path.exists():
                    try:
                        zip_path.unlink()
                    except PermissionError:
                        sufijo = 1
                        while True:
                            zip_path = zip_target.with_name(f"{zip_target.name}_{sufijo}").with_suffix(".zip")
                            if not zip_path.exists():
                                break
                            sufijo += 1
                shutil.make_archive(str(zip_path.with_suffix("")), "zip", zip_target)
                if zip_path.exists():
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            " Descargar ZIP de la carpeta",
                            f,
                            file_name=zip_path.name,
                            use_container_width=True,
                        )

    if st.session_state.download_status in {"running", "cancelling"}:
        # Intervalo de polling del estado de descarga. Cada `st.rerun()`
        # produce un overlay oscuro momentáneo en la UI de Streamlit; a
        # 0.6s la pantalla parpadea constantemente mientras descarga. A
        # 1.5s es mucho menos perceptible y el progreso sigue siendo
        # cómodo. Configurable via env var STREAMLIT_REFRESH_SEC.
        try:
            _refresh_sec = max(0.3, float(os.getenv("STREAMLIT_REFRESH_SEC", "1.5")))
        except (TypeError, ValueError):
            _refresh_sec = 1.5
        time.sleep(_refresh_sec)
        try:
            st.rerun()
        except Exception:
            pass

with tab2:
    st.markdown(
        "<div class='pagehead'>"
        "<h1>Reporte e Historial</h1>"
        "<p>Genera reportes por fechas y revisa las ejecuciones recientes.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    with _group_card(1, "Reporte por fechas", "Excel y PDF"):

        if "custom_report_base_dir" not in st.session_state:
            st.session_state["custom_report_base_dir"] = st.session_state.get("download_base_dir", str(DESC_DIR))
        if "_custom_report_base_dir_pending" in st.session_state:
            pending_custom_dir = st.session_state.pop("_custom_report_base_dir_pending")
            st.session_state["custom_report_base_dir"] = pending_custom_dir
            st.session_state["custom_report_base_dir_input"] = pending_custom_dir
        if "custom_report_base_dir_input" not in st.session_state:
            st.session_state["custom_report_base_dir_input"] = st.session_state.get("custom_report_base_dir", str(DESC_DIR))
        custom_dir_value = st.session_state.get("custom_report_base_dir_input", str(DESC_DIR))
        st.text_input(
            "Carpeta fuente",
            key="custom_report_base_dir_input",
            help="Selecciona la carpeta donde ya tienes descargados los comprobantes.",
        )
        if st.button("Seleccionar carpeta fuente", key="btn_select_custom_report_dir"):
            seleccionada, error = _select_directory_dialog(custom_dir_value)
            if seleccionada:
                st.session_state["_custom_report_base_dir_pending"] = str(Path(seleccionada).expanduser())
                st.rerun()
            if error:
                st.warning(error)

        meses_es_report = [
            "Enero",
            "Febrero",
            "Marzo",
            "Abril",
            "Mayo",
            "Junio",
            "Julio",
            "Agosto",
            "Septiembre",
            "Octubre",
            "Noviembre",
            "Diciembre",
        ]
        cr1, cr2, cr3 = st.columns([1.2, 1.2, 1.2])
        with cr1:
            custom_origen = st.selectbox("Origen", ["Recibidos", "Emitidos"], key="custom_report_origen")
        with cr2:
            custom_tipo = st.selectbox(
                "Tipo de comprobante",
                [
                    "Facturas",
                    "Retenciones",
                    "Notas de crédito",
                    "Notas de débito",
                    "Liquidación de compra",
                    "Guía de remisión",
                ],
                key="custom_report_tipo",
            )
        with cr3:
            custom_estado_emitidos = None
            if custom_origen == "Emitidos":
                custom_estado_emitidos = st.selectbox(
                    "Estado de autorización",
                    ["Autorizados", "No autorizados"],
                    key="custom_report_estado_emitidos",
                )

        custom_mode = st.radio(
            "Modo de fecha",
            ["Día específico", "Mes completo", "Rango de fechas", "Rango de meses", "Año completo"],
            horizontal=True,
            key="custom_report_mode",
        )
        fecha_inicio_custom = None
        fecha_fin_custom = None
        today_local = date.today()
        if custom_mode == "Día específico":
            fecha_unica = st.date_input("Fecha", value=today_local, key="custom_report_single_date")
            fecha_inicio_custom = fecha_unica
            fecha_fin_custom = fecha_unica
        elif custom_mode == "Mes completo":
            cm1, cm2 = st.columns([1, 1])
            with cm1:
                custom_year = st.number_input(
                    "Año",
                    min_value=2015,
                    max_value=today_local.year,
                    value=today_local.year,
                    step=1,
                    key="custom_report_year_month",
                )
            with cm2:
                custom_month_label = st.selectbox(
                    "Mes",
                    meses_es_report,
                    index=max(0, today_local.month - 1),
                    key="custom_report_month_label",
                )
            custom_month = meses_es_report.index(custom_month_label) + 1
            fecha_inicio_custom = date(int(custom_year), custom_month, 1)
            fecha_fin_custom = date(int(custom_year), custom_month, calendar.monthrange(int(custom_year), custom_month)[1])
        elif custom_mode == "Rango de fechas":
            rf1, rf2 = st.columns([1, 1])
            with rf1:
                fecha_inicio_custom = st.date_input("Fecha inicio", value=today_local.replace(day=1), key="custom_report_start_date")
            with rf2:
                fecha_fin_custom = st.date_input("Fecha fin", value=today_local, key="custom_report_end_date")
        elif custom_mode == "Rango de meses":
            rm1, rm2, rm3 = st.columns([1, 1, 1])
            with rm1:
                custom_year = st.number_input(
                    "Año",
                    min_value=2015,
                    max_value=today_local.year,
                    value=today_local.year,
                    step=1,
                    key="custom_report_year_range",
                )
            with rm2:
                custom_month_start_label = st.selectbox(
                    "Mes inicio",
                    meses_es_report,
                    index=0,
                    key="custom_report_month_start",
                )
            with rm3:
                custom_month_end_label = st.selectbox(
                    "Mes fin",
                    meses_es_report,
                    index=max(0, today_local.month - 1),
                    key="custom_report_month_end",
                )
            custom_month_start = meses_es_report.index(custom_month_start_label) + 1
            custom_month_end = meses_es_report.index(custom_month_end_label) + 1
            fecha_inicio_custom = date(int(custom_year), custom_month_start, 1)
            fecha_fin_custom = date(int(custom_year), custom_month_end, calendar.monthrange(int(custom_year), custom_month_end)[1])
        else:
            custom_year = st.number_input(
                "Año",
                min_value=2015,
                max_value=today_local.year,
                value=today_local.year,
                step=1,
                key="custom_report_year_full",
            )
            fecha_inicio_custom = date(int(custom_year), 1, 1)
            fecha_fin_custom = date(int(custom_year), 12, 31)

        if st.button("Generar reporte por fechas", key="btn_generate_custom_report", use_container_width=True):
            source_dir = Path(st.session_state.get("custom_report_base_dir_input") or "").expanduser()
            if not source_dir.exists():
                st.error("La carpeta fuente no existe. Selecciona una ruta válida.")
            elif fecha_inicio_custom is None or fecha_fin_custom is None:
                st.error("Debes definir un rango de fechas válido.")
            elif fecha_inicio_custom > fecha_fin_custom:
                st.error("La fecha inicio no puede ser mayor que la fecha fin.")
            else:
                resultado_custom = _build_custom_report_from_folder(
                    source_dir,
                    origen=custom_origen,
                    tipo=custom_tipo,
                    fecha_inicio=fecha_inicio_custom,
                    fecha_fin=fecha_fin_custom,
                    estado_emitidos=custom_estado_emitidos,
                )
                st.session_state["custom_report_result"] = resultado_custom

        custom_report_result = st.session_state.get("custom_report_result")
        if isinstance(custom_report_result, dict):
            if custom_report_result.get("ok") and custom_report_result.get("path"):
                output_path = Path(custom_report_result["path"])
                st.success(
                    f"Reporte generado. Documentos incluidos: {custom_report_result.get('rows', 0)} "
                    f"(XML: {custom_report_result.get('xml_count', 0)} | PDF: {custom_report_result.get('pdf_count', 0)})."
                )
                if custom_report_result.get("errors"):
                    st.caption(f"PDF con advertencias omitidos: {len(custom_report_result.get('errors') or [])}")
                if output_path.exists():
                    with open(output_path, "rb") as custom_file:
                        st.download_button(
                            "Descargar reporte por fechas",
                            custom_file,
                            file_name=output_path.name,
                            use_container_width=True,
                            key="btn_download_custom_report",
                        )
            elif custom_report_result.get("message"):
                st.warning(custom_report_result.get("message"))

    st.markdown('<h3 class="historial-title">Historial de ejecuciones recientes</h3>', unsafe_allow_html=True)
    historial = obtener_historial(DEVICE_FINGERPRINT)
    historial_raw = historial.copy()

    #  Evitar error valor de verdad de un DataFrame es ambiguo
    if isinstance(historial, pd.DataFrame) and not historial.empty:
        st.markdown(
            """
            <style>
            /* Tabla de historial — usa las variables de tema (var --table-*)
               definidas en el :root global, asi acompana el toggle
               claro/oscuro sin CSS condicional. */
            .historial-table { width: 100%; overflow-x: auto; }
            .historial-table table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                border-radius: 12px;
                overflow: hidden;
                font-size: 0.88rem;
                border: 1px solid var(--border);
            }
            .historial-table th,
            .historial-table td {
                text-align: center;
                padding: 9px 12px;
                border-bottom: 1px solid var(--border);
                border-right: 1px solid var(--border);
                vertical-align: middle;
                white-space: nowrap;
                color: var(--text);
            }
            .historial-table thead th {
                font-weight: 700;
                letter-spacing: 0.2px;
                color: var(--text-strong);
                background: var(--table-head);
                border-bottom: 1px solid var(--border-strong);
            }
            .historial-table tbody td {
                background: var(--table-row);
            }
            .historial-table tbody tr:nth-child(even) td {
                background: var(--table-row-alt);
            }
            .historial-table tbody tr:hover td {
                background: var(--table-hover);
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        def _opciones_columna(df, col, ordenar_numerico=False):
            if col not in df.columns:
                return ["Todos"]
            valores = [
                str(v).strip()
                for v in df[col].dropna().unique().tolist()
                if str(v).strip() and str(v).lower() != "nan"
            ]
            if ordenar_numerico:
                try:
                    valores = sorted({int(v) for v in valores})
                    valores = [str(v) for v in valores]
                except Exception:
                    valores = sorted(set(valores))
            else:
                valores = sorted(set(valores))
            return ["Todos"] + valores

        def _selectbox_con_opciones(label, opciones, key):
            seleccion = st.session_state.get(key, "Todos")
            if seleccion not in opciones:
                seleccion = "Todos"
                st.session_state[key] = seleccion
            return st.selectbox(label, opciones, key=key, index=opciones.index(seleccion))

        def _aplicar_filtro(df, col, valor):
            if valor != "Todos" and col in df.columns:
                return df[df[col].astype(str) == str(valor)]
            return df

        st.markdown("##### Filtros")
        base_opciones = historial.copy()
        if "dia" in base_opciones.columns:
            base_opciones["dia"] = base_opciones["dia"].apply(
                lambda x: "Todos" if str(x) in {"0", "Todos", "None"} else x
            )
        f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
        with f1:
            filtro_busqueda = st.text_input(
                "Búsqueda",
                placeholder="RUC, descripcion, estado, etc.",
                key="historial_busqueda",
            ).strip()
        with f2:
            opciones_ruc = _opciones_columna(base_opciones, "ruc")
            filtro_ruc = _selectbox_con_opciones(
                "RUC",
                opciones_ruc,
                "historial_filtro_ruc",
            )
        base_opciones = _aplicar_filtro(base_opciones, "ruc", filtro_ruc)
        with f3:
            opciones_origen = _opciones_columna(base_opciones, "origen")
            filtro_origen = _selectbox_con_opciones(
                "Origen",
                opciones_origen,
                "historial_filtro_origen",
            )
        base_opciones = _aplicar_filtro(base_opciones, "origen", filtro_origen)
        with f4:
            opciones_tipo = _opciones_columna(base_opciones, "tipo")
            filtro_tipo = _selectbox_con_opciones(
                "Tipo",
                opciones_tipo,
                "historial_filtro_tipo",
            )
        base_opciones = _aplicar_filtro(base_opciones, "tipo", filtro_tipo)

        f5, f6, f7, f8 = st.columns([1, 1, 1, 1])
        with f5:
            opciones_estado_aut = _opciones_columna(base_opciones, "estado_autorizacion")
            filtro_estado_aut = _selectbox_con_opciones(
                "Estado autorizacion",
                opciones_estado_aut,
                "historial_filtro_estado_aut",
            )
        base_opciones = _aplicar_filtro(base_opciones, "estado_autorizacion", filtro_estado_aut)
        with f6:
            opciones_anio = _opciones_columna(base_opciones, "anio", ordenar_numerico=True)
            filtro_anio = _selectbox_con_opciones(
                "Año",
                opciones_anio,
                "historial_filtro_anio",
            )
        base_opciones = _aplicar_filtro(base_opciones, "anio", filtro_anio)
        with f7:
            opciones_mes = _opciones_columna(base_opciones, "mes", ordenar_numerico=True)
            filtro_mes = _selectbox_con_opciones(
                "Mes",
                opciones_mes,
                "historial_filtro_mes",
            )
        base_opciones = _aplicar_filtro(base_opciones, "mes", filtro_mes)
        with f8:
            if filtro_mes == "Todos":
                opciones_dia = ["Todos"]
            else:
                opciones_dia = _opciones_columna(base_opciones, "dia", ordenar_numerico=True)
            filtro_dia = _selectbox_con_opciones(
                "Dia",
                opciones_dia,
                "historial_filtro_dia",
            )
        columnas_amigables = {
            "timestamp": "Fecha y hora",
            "ruc": "RUC",
            "origen": "Origen",
            "anio": "A\u00f1o",
            "mes": "Mes",
            "dia": "D\u00eda",
            "tipo": "Tipo de comprobante",
            "estado": "Estado",
            "fecha_filtro": "Fecha filtro",
            "estado_autorizacion": "Estado autorizaci\u00f3n",
            "establecimiento": "Establecimiento",
            "punto_emision": "Punto de emisi\u00f3n",
            "tipo_visible": "Descripci\u00f3n",
        }
        if "dia" in historial.columns:
            historial["dia"] = historial["dia"].apply(lambda x: "Todos" if str(x) in {"0", "Todos", "None"} else x)
        if "fecha_filtro" in historial.columns:
            def _rango_mes(row):
                valor = str(row.get("fecha_filtro", "")).strip()
                if valor and valor.lower() != "nan":
                    return valor
                dia_val = str(row.get("dia", "")).strip()
                if dia_val != "Todos":
                    return valor
                try:
                    anio_val = int(row.get("anio"))
                    mes_val = int(row.get("mes"))
                except (TypeError, ValueError):
                    return valor
                ultimo = calendar.monthrange(anio_val, mes_val)[1]
                return f"01/{mes_val:02d}/{anio_val} - {ultimo:02d}/{mes_val:02d}/{anio_val}"
            historial["fecha_filtro"] = historial.apply(_rango_mes, axis=1)
        for columna in ("fecha_filtro", "estado_autorizacion", "establecimiento", "punto_emision"):
            if columna in historial.columns:
                historial[columna] = historial[columna].replace("", "N/A").fillna("N/A")

        filtro_idx = historial.index
        if filtro_ruc != "Todos" and "ruc" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["ruc"].astype(str) == filtro_ruc].index
            )
        if filtro_origen != "Todos" and "origen" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["origen"].astype(str) == filtro_origen].index
            )
        if filtro_tipo != "Todos" and "tipo" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["tipo"].astype(str) == filtro_tipo].index
            )
        if filtro_estado_aut != "Todos" and "estado_autorizacion" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["estado_autorizacion"].astype(str) == filtro_estado_aut].index
            )
        if filtro_anio != "Todos" and "anio" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["anio"].astype(str) == filtro_anio].index
            )
        if filtro_mes != "Todos" and "mes" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["mes"].astype(str) == filtro_mes].index
            )
        if filtro_dia != "Todos" and "dia" in historial.columns:
            filtro_idx = filtro_idx.intersection(
                historial[historial["dia"].astype(str) == filtro_dia].index
            )
        if filtro_busqueda:
            busqueda = filtro_busqueda.lower()
            columnas_busqueda = [
                col
                for col in [
                    "timestamp",
                    "ruc",
                    "origen",
                    "tipo",
                    "estado",
                    "tipo_visible",
                    "fecha_filtro",
                    "estado_autorizacion",
                    "establecimiento",
                    "punto_emision",
                ]
                if col in historial.columns
            ]
            if columnas_busqueda:
                mask = historial[columnas_busqueda].astype(str).apply(
                    lambda fila: busqueda in " ".join(fila.values).lower(), axis=1
                )
                filtro_idx = filtro_idx.intersection(historial[mask].index)

        historial = historial.loc[filtro_idx].copy()
        if isinstance(historial_raw, pd.DataFrame) and not historial_raw.empty:
            historial_raw = historial_raw.loc[filtro_idx].copy()

        historial = historial.reset_index(drop=True)
        historial = historial.drop(
            columns=[
                "n_registros",
                "n_xml",
                "n_pdf",
                "reporte",
                "reporte_xml",
                "reporte_pdf",
                "reporte_pdf_anual",
                "reporte_xml_anual",
                "reportes_xml",
                "reportes_pdf",
                "xml_dir",
                "pdf_dir",
            ],
            errors="ignore",
        )

        historial.insert(0, "No.", range(1, len(historial) + 1))
        historial = historial.rename(columns=columnas_amigables)


        # Pills coloreadas para ESTADO y AUTORIZACION (estilo mockup).
        def _pill_estado(val):
            v = str(val).strip().lower()
            if v in {'ok', 'exitoso', 'completado'}:
                cls = 'pill pill-ok'
            elif v in {'pendiente', 'en proceso', 'procesando'}:
                cls = 'pill pill-pendiente'
            elif v in {'error', 'fallido'}:
                cls = 'pill pill-error'
            else:
                cls = 'pill pill-other'
            return f"<span class='{cls}'>&#9679; {html.escape(str(val))}</span>"

        def _pill_autorizacion(val):
            v = str(val).strip().lower()
            if 'no autoriz' in v:
                cls = 'pill pill-auth-no'
            elif 'autoriz' in v:
                cls = 'pill pill-auth-yes'
            else:
                cls = 'pill pill-auth-none'
            return f"<span class='{cls}'>&#9679; {html.escape(str(val))}</span>"

        if 'Estado' in historial.columns:
            historial['Estado'] = historial['Estado'].apply(_pill_estado)
        if 'Estado autorización' in historial.columns:
            historial['Estado autorización'] = historial['Estado autorización'].apply(_pill_autorizacion)

        tabla_html = historial.to_html(
            index=False,
            escape=False,
            border=0,
            classes="historial-table-grid",
        )
        st.markdown(f"<div class='historial-table'>{tabla_html}</div>", unsafe_allow_html=True)
        _total_count = len(historial)
        st.markdown(
            f'<div class="historial-total">'
            f'<span class="historial-total-label">Total de operaciones registradas</span>'
            f'<span class="historial-total-badge">&#9679; {_total_count}</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        descargables = []
        if isinstance(historial_raw, pd.DataFrame):
            hist_reset = historial_raw.reset_index(drop=True)
            for idx, fila in hist_reset.iterrows():
                rutas = [
                    fila.get("reporte"),
                    fila.get("reporte_xml"),
                    fila.get("reporte_pdf"),
                    fila.get("reporte_xml_anual"),
                    fila.get("reporte_pdf_anual"),
                ]
                for ruta_val in rutas:
                    if not ruta_val or not isinstance(ruta_val, (str, Path)):
                        continue
                    ruta_path = Path(ruta_val).expanduser()
                    if not ruta_path.exists():
                        continue
                    etiqueta = f"{idx + 1}. {ruta_path.name} ({fila.get('timestamp', '')})"
                    descargables.append((etiqueta, ruta_path))

        if False and descargables:
            st.markdown("##### Descargar reporte generado anteriormente")
            opciones = ["Seleccione un reporte"] + [etq for etq, _ in descargables]
            seleccion = st.selectbox(
                "Selecciona un reporte disponible",
                opciones,
                key="historial_descarga_select",
                label_visibility="collapsed",
            )
            if seleccion != "Seleccione un reporte":
                ruta_sel = next(path for etq, path in descargables if etq == seleccion)
                try:
                    with open(ruta_sel, "rb") as archivo:
                        st.download_button(
                            f"Descargar {ruta_sel.name}",
                            archivo,
                            file_name=ruta_sel.name,
                            use_container_width=True,
                        )
                except Exception as err:
                    st.warning(f"No se pudo abrir el archivo seleccionado: {err}")
    else:
        st.info("A\u00fan no hay registros de descargas o reportes.")

with tab3:
    st.markdown(
        "<div class='pagehead'>"
        "<h1>Consolidar desde carpeta</h1>"
        "<p>Genera reportes consolidados desde documentos ya descargados.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    with _group_card(1, "Carpeta origen a consolidar", "Dónde buscar documentos"):
        consolidar_dir_actual = st.session_state.get(
            "consolidate_base_dir",
            st.session_state.get("download_base_dir", str(DESC_DIR)),
        )
        st.text_input(
            "Carpeta origen para consolidar",
            value=consolidar_dir_actual,
            help="Puedes elegir la carpeta base de descargas, una carpeta RUC o una subcarpeta de Recibidos/Emitidos.",
            disabled=True,
        )
        col_sel_c1, col_sel_c2 = st.columns([1, 1])
        with col_sel_c1:
            if st.button("Seleccionar carpeta para consolidar", key="btn_select_consolidate_dir"):
                seleccionada, error = _select_directory_dialog(consolidar_dir_actual)
                if seleccionada:
                    try:
                        ruta_cons = Path(seleccionada).expanduser()
                        ruta_cons.mkdir(parents=True, exist_ok=True)
                        st.session_state["consolidate_base_dir"] = str(ruta_cons)
                        _persist_user_preferences()
                        st.success(f" Carpeta de consolidacion: {ruta_cons}")
                    except Exception as err:
                        st.error(f"No se pudo usar la carpeta indicada: {err}")
                elif error:
                    st.session_state["show_manual_consolidate_dir"] = True
                    st.session_state["last_manual_consolidate_dir_error"] = error
        with col_sel_c2:
            if st.button("Usar carpeta de descargas activa", key="btn_use_download_dir_for_consolidate"):
                st.session_state["consolidate_base_dir"] = st.session_state.get("download_base_dir", str(DESC_DIR))
                _persist_user_preferences()
                st.success("Se usara la carpeta de descargas activa para consolidar.")

        if st.session_state.get("show_manual_consolidate_dir"):
            manual_default_cons = st.session_state.get(
                "consolidate_base_dir",
                st.session_state.get("download_base_dir", str(DESC_DIR)),
            )
            manual_cons = st.text_input(
                "Ruta de carpeta para consolidar (manual)",
                value=manual_default_cons,
                key="manual_consolidate_dir_input",
            )
            if st.button("Guardar carpeta de consolidacion", key="btn_save_manual_consolidate_dir"):
                try:
                    nueva_ruta = Path(manual_cons).expanduser()
                    nueva_ruta.mkdir(parents=True, exist_ok=True)
                    st.session_state["consolidate_base_dir"] = str(nueva_ruta)
                    _persist_user_preferences()
                    st.success(f" Carpeta de consolidacion: {nueva_ruta}")
                    st.session_state["show_manual_consolidate_dir"] = False
                    st.session_state["last_manual_consolidate_dir_error"] = None
                except Exception as err:
                    st.error(f"No se pudo usar la carpeta indicada: {err}")
            last_err_cons = st.session_state.get("last_manual_consolidate_dir_error")
            if last_err_cons:
                if "tk" in str(last_err_cons).lower() or "libtk" in str(last_err_cons).lower():
                    st.info("El selector nativo no esta disponible en este entorno. Usa la ruta manual.")
                else:
                    st.warning(last_err_cons)

        st.caption(
            f"Carpeta de busqueda activa para consolidacion: `{st.session_state.get('consolidate_base_dir', consolidar_dir_actual)}`"
        )

    with _group_card(2, "Filtros", "RUC, origen, tipo, año y periodo"):
        col_ruc_cons, col_origen_cons = st.columns([1.2, 1.2])
        with col_ruc_cons:
            ruc_consolidar = st.text_input(
                "RUC a buscar (opcional)",
                value=ruc,
                help="Si lo dejas vacio, se intentara consolidar desde la carpeta seleccionada.",
                key="consolidar_ruc_hint",
            )
        with col_origen_cons:
            origen_consolidar = st.selectbox(
                "Origen a consolidar",
                ["Recibidos", "Emitidos"],
                index=0 if origen == "Recibidos" else 1,
                key="consolidar_origen",
            )

        col_c1, col_c2, col_c3 = st.columns([1.2, 1, 1])
        with col_c1:
            if origen_consolidar == "Emitidos":
                tipos_disponibles = [
                    "Facturas",
                    "Liquidacion de compra",
                    "Guia de remision",
                    "Retencion",
                    "Notas de debito",
                    "Notas de credito",
                ]
            else:
                tipos_disponibles = [
                    "Retencion",
                    "Facturas",
                    "Notas de debito",
                    "Notas de credito",
                    "Liquidacion de compra",
                ]
            tipo_consolidar = st.selectbox(
                "Tipo de comprobante",
                tipos_disponibles,
                index=tipos_disponibles.index(tipo) if tipo in tipos_disponibles else 0,
                key="consolidar_tipo",
            )
        with col_c2:
            anio_consolidar = st.number_input(
                "Año a consolidar",
                min_value=2015,
                max_value=datetime.now().year,
                value=int(datetime.now().year),
                step=1,
                key="consolidar_anio",
            )
        with col_c3:
            estado_consolidar = None
            if origen_consolidar == "Emitidos":
                estado_default = (
                    st.session_state.get("estado_autorizacion")
                    or (estado_emitidos if "estado_emitidos" in locals() else None)
                    or "Autorizados"
                )
                estado_consolidar = st.selectbox(
                    "Estado autorizacion",
                    ["Autorizados", "No Autorizados"],
                    index=0 if estado_default == "Autorizados" else 1,
                    key="consolidar_estado",
                )
            else:
                st.write("")

        modo_fecha_consolidar = st.radio(
            "Modo de fecha",
            ["Mes y dia", "Rango de meses", "Año completo"],
            horizontal=True,
            key="consolidar_modo_fechas",
        )

        mes_inicio_consolidar = 1
        mes_fin_consolidar = 12
        dia_consolidar = 0
        mes_actual = int(datetime.now().month)
        if modo_fecha_consolidar == "Rango de meses":
            col_fr1, col_fr2 = st.columns([1, 1])
            with col_fr1:
                mes_inicio_label = st.selectbox(
                    "Mes inicio",
                    MESES_ES,
                    index=mes_actual - 1,
                    key="consolidar_mes_inicio",
                )
            with col_fr2:
                mes_fin_label = st.selectbox(
                    "Mes fin",
                    MESES_ES,
                    index=mes_actual - 1,
                    key="consolidar_mes_fin",
                )
            mes_inicio_consolidar = MESES_ES.index(mes_inicio_label) + 1
            mes_fin_consolidar = MESES_ES.index(mes_fin_label) + 1
        elif modo_fecha_consolidar == "Mes y dia":
            col_fd1, col_fd2 = st.columns([1, 1])
            with col_fd1:
                mes_label = st.selectbox(
                    "Mes",
                    MESES_ES,
                    index=mes_actual - 1,
                    key="consolidar_mes",
                )
                mes_inicio_consolidar = MESES_ES.index(mes_label) + 1
                mes_fin_consolidar = mes_inicio_consolidar
            with col_fd2:
                dia_consolidar = st.number_input(
                    "Dia (0 = Todos)",
                    min_value=0,
                    max_value=31,
                    value=0,
                    step=1,
                    key="consolidar_dia",
                )
        else:
            mes_inicio_consolidar = 1
            mes_fin_consolidar = 12
            dia_consolidar = 0

    with _group_card(3, "Salida", "Formatos y ejecución"):
        col_f1, col_f2 = st.columns([1, 1])
        with col_f1:
            incluir_xml = st.checkbox("Consolidar XML", value=True, key="consolidar_xml")
        with col_f2:
            incluir_pdf = st.checkbox("Consolidar PDF", value=True, key="consolidar_pdf")

        if st.button("Consolidar desde carpeta", use_container_width=True, key="btn_run_consolidation"):
            carpeta_base_cons = Path(
                st.session_state.get(
                    "consolidate_base_dir",
                    st.session_state.get("download_base_dir", str(DESC_DIR)),
                )
            ).expanduser()
            if not carpeta_base_cons.exists():
                st.error(f"La carpeta seleccionada no existe: {carpeta_base_cons}")
                st.stop()
            if not incluir_xml and not incluir_pdf:
                st.warning("Selecciona al menos una opcion: XML o PDF.")
                st.stop()
            if modo_fecha_consolidar == "Rango de meses" and int(mes_fin_consolidar) < int(mes_inicio_consolidar):
                st.error("El mes fin debe ser mayor o igual al mes inicio.")
                st.stop()

            tipo_visible = TIPOS_MAP.get(tipo_consolidar, tipo_consolidar)
            tipo_slug = _slug_tipo(tipo_visible or tipo_consolidar)
            _, _, tipo_prefijo = _prefijo_tipo(tipo_visible or tipo_consolidar)
            anio_int = int(anio_consolidar)
            estado_slug = _slug_estado_emitidos(estado_consolidar or "Sin Estado") if origen_consolidar == "Emitidos" else None
            periodo_suffix = _sufijo_periodo_consolidacion(
                modo_fecha_consolidar,
                anio_int,
                int(mes_inicio_consolidar),
                int(mes_fin_consolidar),
                int(dia_consolidar),
            )

            base_search = _resolver_busqueda_consolidacion(
                carpeta_base_cons,
                origen_consolidar,
                ruc_hint=ruc_consolidar,
                estado_slug=estado_slug,
            )
            if not base_search.exists():
                st.error(f"No se encontro una carpeta valida para buscar reportes: {base_search}")
                st.stop()

            prefix_base = "recibidos_reporte" if origen_consolidar == "Recibidos" else "emitidos_reporte"
            destino_anual_dir = carpeta_base_cons / "Consolidados" / origen_consolidar
            if origen_consolidar == "Emitidos" and estado_slug:
                destino_anual_dir = destino_anual_dir / estado_slug
            if modo_fecha_consolidar == "Ano completo":
                destino_anual_dir = destino_anual_dir / f"{anio_int:04d}"
            else:
                destino_anual_dir = destino_anual_dir / periodo_suffix
            destino_anual_dir.mkdir(parents=True, exist_ok=True)

            st.caption(f"Buscando reportes en: `{base_search}`")
            st.caption(f"Guardando consolidados en: `{destino_anual_dir}`")

            if incluir_xml:
                reportes_xml = _buscar_reportes_por_periodo(
                    base_search,
                    origen_consolidar,
                    tipo_slug,
                    "xml",
                    modo_fecha_consolidar,
                    anio_int,
                    int(mes_inicio_consolidar),
                    int(mes_fin_consolidar),
                    int(dia_consolidar),
                )
                xml_files = _colectar_documentos_por_periodo(
                    base_search,
                    tipo_prefijo,
                    "xml",
                    modo_fecha_consolidar,
                    anio_int,
                    int(mes_inicio_consolidar),
                    int(mes_fin_consolidar),
                    int(dia_consolidar),
                    tipo_slug_archivo=tipo_slug,
                )
                st.caption(f"Reportes XML base encontrados: {len(reportes_xml)}")
                st.caption(f"Documentos XML encontrados: {len(xml_files)}")
                destino_xml = destino_anual_dir / f"{prefix_base}_xml_{tipo_slug}_{periodo_suffix}.xlsx"
                anual_xml = None
                if xml_files:
                    estado_default_reporte = None
                    if origen_consolidar == "Emitidos":
                        estado_txt = (estado_consolidar or "").strip().lower()
                        if "no autoriz" in estado_txt:
                            estado_default_reporte = estado_consolidar
                    try:
                        construir_reporte(base_search, destino_xml, estado_default_reporte, xml_files=xml_files)
                    except Exception as err:
                        st.error(f"No se pudo generar el reporte XML desde los documentos: {err}")
                    if destino_xml.exists():
                        anual_xml = destino_xml
                elif reportes_xml:
                    anual_xml = _consolidar_reportes_xml_desde_excels(reportes_xml, destino_xml)
                if anual_xml:
                    st.success(f"Reporte XML consolidado: {anual_xml}")
                else:
                    st.info("No se encontraron insumos XML para consolidar.")

                destino_copia_xml = destino_anual_dir / "XML"
                copiados_xml = _copiar_documentos_unicos(xml_files, destino_copia_xml)
                if copiados_xml > 0:
                    st.success(f"XML copiados: {copiados_xml} en `{destino_copia_xml}`")
                else:
                    st.info("No se copiaron XML porque no hubo documentos para el periodo seleccionado.")

            if incluir_pdf:
                reportes_pdf = _buscar_reportes_por_periodo(
                    base_search,
                    origen_consolidar,
                    tipo_slug,
                    "pdf",
                    modo_fecha_consolidar,
                    anio_int,
                    int(mes_inicio_consolidar),
                    int(mes_fin_consolidar),
                    int(dia_consolidar),
                )
                pdf_files = _colectar_documentos_por_periodo(
                    base_search,
                    tipo_prefijo,
                    "pdf",
                    modo_fecha_consolidar,
                    anio_int,
                    int(mes_inicio_consolidar),
                    int(mes_fin_consolidar),
                    int(dia_consolidar),
                    tipo_slug_archivo=tipo_slug,
                )
                st.caption(f"Reportes PDF encontrados: {len(reportes_pdf)}")
                st.caption(f"Documentos PDF encontrados: {len(pdf_files)}")
                if reportes_pdf:
                    destino_pdf = destino_anual_dir / f"{prefix_base}_pdf_{tipo_slug}_{periodo_suffix}.xlsx"
                    anual_pdf = _consolidar_reportes_excel(
                        [str(p) for p in reportes_pdf], destino_pdf
                    )
                    if anual_pdf:
                        st.success(f"Reporte PDF consolidado: {anual_pdf}")
                    else:
                        st.error("No se pudo generar el reporte PDF consolidado.")
                else:
                    st.info("No se encontraron reportes PDF para consolidar.")

                destino_copia_pdf = destino_anual_dir / "PDF"
                copiados_pdf = _copiar_documentos_unicos(pdf_files, destino_copia_pdf)
                if copiados_pdf > 0:
                    st.success(f"PDF copiados: {copiados_pdf} en `{destino_copia_pdf}`")
                else:
                    st.info("No se copiaron PDF porque no hubo documentos para el periodo seleccionado.")

with tab4:
    # ===== Encabezado =====
    st.markdown(
        "<div class='pagehead'>"
        "<h1>Centro de ayuda</h1>"
        "<p>Guía rápida para usar el sistema paso a paso.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # ===== Quickstart cards (numeradas 01-04) =====
    st.markdown(
        """
        <div class="quickstart">
          <div class="qs"><div class="n">01</div><h4>Inicia sesión</h4>
            <p>Verifica tu correo, contraseña y licencia activa.</p></div>
          <div class="qs"><div class="n">02</div><h4>Descarga comprobantes</h4>
            <p>Ingresa RUC, clave del SRI y filtros.</p></div>
          <div class="qs"><div class="n">03</div><h4>Genera o consolida</h4>
            <p>Elige formato (XML/PDF) según lo que necesites.</p></div>
          <div class="qs"><div class="n">04</div><h4>Revisa el historial</h4>
            <p>Consulta el estado de cada ejecución.</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ===== Acciones del tour =====
    col_help_1, col_help_2 = st.columns([1, 1])
    with col_help_1:
        if st.button("Activar tour de primer uso", key="help_start_tour", use_container_width=True):
            _start_first_use_tour(reset_step=True)
            st.success("Tour activado. Ve a la pestana 'Descarga de Comprobantes' para verlo.")
    with col_help_2:
        if st.button("Marcar tour como no visto", key="help_reset_tour", use_container_width=True):
            st.session_state["first_use_tour_completed"] = False
            st.session_state["first_use_tour_active"] = True
            st.session_state["first_use_tour_step"] = 0
            _persist_user_preferences()
            st.success("Listo. El tour volvera a mostrarse.")

    # ===== Accordion de ayuda (usando <details> nativo, estilizado como el mockup) =====
    st.markdown(
        """
        <div class="acc">
          <details open>
            <summary><span class="ic">●</span> Inicio de sesión y licencias <span class="pm">+</span></summary>
            <div class="body">Si no puedes entrar, verifica correo, contraseña y estado de licencia.
              El acceso depende de la base en Render; si el usuario no existe o no tiene licencia activa,
              no podrá ingresar. La sesión queda vinculada a este dispositivo.</div>
          </details>
          <details>
            <summary><span class="ic">●</span> Descarga de Recibidos <span class="pm">+</span></summary>
            <div class="body">Usa Recibidos para descargar comprobantes por mes, día, rango de meses
              o año completo. Puedes combinar XML y PDF en la misma ejecución.</div>
          </details>
          <details>
            <summary><span class="ic">●</span> Descarga de Emitidos <span class="pm">+</span></summary>
            <div class="body">En Emitidos define estado de autorización, establecimiento y punto de
              emisión si aplica. Para XML de Emitidos Autorizados, el sistema valida automáticamente
              el límite operativo de 30 días.</div>
          </details>
          <details>
            <summary><span class="ic">●</span> Consolidación desde carpeta <span class="pm">+</span></summary>
            <div class="body">Permite generar reportes consolidados desde documentos ya descargados.
              Puedes consolidar XML, PDF o ambos y copiar todos los archivos encontrados al
              directorio final.</div>
          </details>
          <details>
            <summary><span class="ic">●</span> Errores frecuentes y qué hacer <span class="pm">+</span></summary>
            <div class="body">
              · Timeout o portal lento: vuelve a intentar en unos segundos.<br>
              · Captcha incorrecta: espera 1-2 minutos y repite la consulta.<br>
              · Sin resultados: valida rango de fechas, tipo y estado seleccionado.<br>
              · No descarga archivos: revisa permisos de carpeta y espacio disponible.
            </div>
          </details>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ===== Acerca de la aplicación (versión + buscar actualizaciones) =====
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    with _group_card("i", "Acerca de la aplicación", "Información y actualizaciones"):
        col_about_1, col_about_2 = st.columns([3, 2])
        with col_about_1:
            st.markdown(
                f"""
                <div class="about-card" style="background:transparent;border:0;padding:0">
                  <div class="row">
                    <span class="label">Versión actual</span>
                    <span class="value">v{html.escape(_app_version_display)}</span>
                  </div>
                  <div class="row" style="margin-top:.55rem">
                    <span class="label">Modo</span>
                    <span class="value">{'compilado (.exe)' if getattr(sys, 'frozen', False) else 'desarrollo'}</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_about_2:
            if st.button(
                "🔄 Buscar actualizaciones",
                help="Verifica si hay una versión más reciente disponible",
                use_container_width=True,
                key="btn_buscar_update",
            ):
                # Reset del estado para forzar un nuevo chequeo.
                st.session_state.pop("_update_checked", None)
                st.session_state.pop("_update_message", None)
                st.session_state["_manual_update_check"] = True
                st.rerun()

        # Resultado del chequeo manual (solo se ejecuta UNA vez tras el click).
        # MISMA logica que vivia antes en el sidebar — solo se traslada de
        # ubicacion visual, sin cambiar el mecanismo interno.
        if st.session_state.pop("_manual_update_check", False):
            if not getattr(sys, "frozen", False):
                st.info(
                    "La auto-actualización solo funciona en el ejecutable .exe "
                    "compilado, no en modo desarrollo."
                )
            elif _desktop_launcher is None:
                st.warning("Módulo de actualización no disponible en este build.")
            else:
                with st.spinner("Verificando actualizaciones..."):
                    try:
                        _update_payload = _get_update_payload()
                    except Exception as _upd_err:
                        _update_payload = None
                        st.error(f"Error al verificar: {_upd_err}")
                if _update_payload:
                    _new_ver = _update_payload.get("version", "?")
                    _msg = (
                        f"Nueva versión **{_new_ver}** disponible. "
                        f"Descargando y aplicando... la app se reiniciará en breve."
                    )
                    st.session_state["_update_message"] = _msg
                    st.session_state["_update_checked"] = True
                    st.success(_msg)

                    def _manual_update_worker():
                        try:
                            _start_update(_update_payload)
                        except Exception:
                            pass

                    threading.Thread(target=_manual_update_worker, daemon=True).start()
                else:
                    st.session_state["_update_checked"] = True
                    st.success(
                        f"Ya estás en la última versión ({_app_version_display}). "
                        "No hay actualizaciones disponibles."
                    )

        # Mensaje persistente de actualización en curso (si quedó del auto-check).
        _persistent_msg = st.session_state.get("_update_message")
        if _persistent_msg and not st.session_state.pop("_manual_update_check", False):
            st.caption(_persistent_msg)