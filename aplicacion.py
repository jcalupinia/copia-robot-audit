# =====================================================
#  SRI ROBOT AUDIT  APLICACIN PRINCIPAL STREAMLIT
# Versin estable para Render.com / Octubre 2025
# =====================================================

import os
import shutil
import hashlib
import platform
import time
import uuid
import json
import calendar
import secrets
import smtplib
import ssl
import base64
from datetime import datetime, timedelta, date
from pathlib import Path
from email.message import EmailMessage

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from robot.downloader import descargar_sri, set_user_notifier, MANUAL_CONSULTA_RECIBIDOS
from robot.parser import construir_reporte
from robot.historial import registrar_descarga, obtener_historial   #  FIX import correcto
from licensing_client import LicensingClient
# Para restablecer contraseñas directamente en la base local
try:
    from licensing_api.database import SessionLocal
    from licensing_api import crud as lic_crud, security as lic_security
except Exception:
    SessionLocal = None
    lic_crud = None
    lic_security = None

import asyncio
import threading
import queue

if os.name == "nt":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())


@st.cache_data(show_spinner=False)
def _get_logo_data_uri():
    logo_path = Path(__file__).parent / "LogoAUDIT.png"
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

def _drain_download_queue():
    q = st.session_state.download_queue
    while True:
        try:
            kind, payload = q.get_nowait()
        except queue.Empty:
            break
        if kind == "msg":
            st.session_state.download_messages.append(str(payload))
            st.session_state.last_download_message = (str(payload), time.time())
        elif kind == "done":
            st.session_state.download_result = payload
            st.session_state.download_status = "done"
        elif kind == "error":
            st.session_state.download_error = str(payload)
            st.session_state.download_status = "error"

def _download_worker(params: dict, q: "queue.Queue"):
    from robot.downloader import descargar_sri, set_user_notifier, clear_cancel

    def _notify(msg: str):
        try:
            q.put(("msg", msg))
        except Exception:
            pass

    clear_cancel()
    set_user_notifier(_notify)
    try:
        resultado = descargar_sri(**params)
        q.put(("done", resultado))
    except Exception as err:
        q.put(("error", str(err)))
    finally:
        set_user_notifier(None)

# ==============================
# CONFIGURACIN GENERAL
# ==============================
st.set_page_config(
    page_title="SRI Robot Audit",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(
    """
    <style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&display=swap');
div[data-testid="stToolbarActions"],
button[title="Deploy"] {
    display:none !important;
}
header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
}
div[data-testid="stDecoration"] {
    display: none !important;
}
button[aria-label=" Iniciar proceso"],
button[aria-label="Iniciar proceso"]{
    background-color:#16a34a !important;
    border-color:#16a34a !important;
    color:#ffffff !important;
}
button[aria-label=" Detener proceso"],
button[aria-label="Detener proceso"]{
    background-color:#dc2626 !important;
    border-color:#dc2626 !important;
    color:#ffffff !important;
}
    :root{
        --auth-card-bg: var(--secondary-background-color, #10131a);
        --auth-card-text: var(--text-color, #f5f5f5);
        --auth-card-muted: rgba(255,255,255,0.65);
    }
    .stApp {
        background: radial-gradient(120% 120% at 10% 10%, rgba(0,128,255,0.35), transparent 50%),
                    radial-gradient(120% 120% at 90% 20%, rgba(0,255,170,0.25), transparent 55%),
                    radial-gradient(120% 120% at 30% 80%, rgba(0,80,160,0.35), transparent 55%),
                    linear-gradient(135deg, #0b0f1a 0%, #0b1f2a 45%, #020508 100%);
        background-size: 200% 200%;
        animation: liquidShift 18s ease-in-out infinite;
    }
    .stApp, .stApp p, .stApp span, .stApp label,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 {
        color: #e9eef5 !important;
    }
    .stApp [data-testid="stMarkdownContainer"] {
        color: #e9eef5 !important;
    }
    .stApp [data-baseweb="tab"] {
        color: #d7e3f4 !important;
    }
    .stApp [data-baseweb="tab"][aria-selected="true"] {
        color: #ffffff !important;
    }
    body[data-theme="dark"] [data-testid="stSidebar"] * {
        color: #e9eef5 !important;
    }
    body[data-theme="light"] [data-testid="stSidebar"] * {
        color: #1f2937 !important;
    }
    @keyframes liquidShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .auth-title{
        text-align:center;
        font-size:2.1rem;
        font-weight:800;
        font-family:'Manrope', sans-serif;
        letter-spacing:0.02em;
        color:#ffffff !important;
        margin-bottom:1.4rem;
    }
    div[data-testid="stForm"]{
        background:var(--auth-card-bg);
        border-radius:24px;
        padding:2.2rem 2.6rem;
        box-shadow:0 28px 55px rgba(0,0,0,0.45);
        border:1px solid rgba(15,18,26,0.18);
        color:var(--auth-card-text);
    }
    div[data-testid="stForm"] label,
    div[data-testid="stForm"] p,
    div[data-testid="stForm"] span{
        color:var(--auth-card-text);
    }
    div[data-testid="stForm"] div[data-baseweb="input"]{
        background:var(--secondary-background-color, #151621);
        border-color:rgba(255,255,255,0.15);
    }
    div[data-testid="stForm"] input,
    div[data-testid="stForm"] textarea{
        color:var(--auth-card-text) !important;
        background:var(--secondary-background-color, #151621);
    }
    div[data-testid="stForm"] input::placeholder,
    div[data-testid="stForm"] textarea::placeholder{
        color:var(--auth-card-muted);
        opacity:0.9;
    }
    div[data-testid="stForm"] div[data-baseweb="input"] button{
        background:transparent !important;
        box-shadow:none !important;
    }
    div[data-testid="stForm"] div[data-baseweb="input"] div{
        background:transparent !important;
    }
    div[data-testid="stForm"] div[data-baseweb="input"] > div:last-child{
        background:var(--secondary-background-color, #151621) !important;
        border-left:1px solid rgba(255,255,255,0.12) !important;
    }
    div[data-testid="stForm"] div[data-baseweb="input"] svg{
        color:var(--auth-card-text) !important;
        fill:var(--auth-card-text) !important;
    }
    div[data-testid="stForm"] button[kind="primaryFormSubmit"],
    div[data-testid="stForm"] button[data-testid="baseButton-primary"]{
        background-color:#101936;
        color:#f5f5f5;
        border:none;
        border-radius:12px;
        font-weight:600;
        box-shadow:0 12px 25px rgba(0,0,0,0.25);
    }
    div[data-testid="stForm"] button[kind="primaryFormSubmit"]:disabled,
    div[data-testid="stForm"] button[data-testid="baseButton-primary"]:disabled{
        background:rgba(16,25,54,0.3);
        color:rgba(255,255,255,0.6);
        border:1px solid rgba(16,25,54,0.4);
        box-shadow:none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

_init_download_state()
_drain_download_queue()

# Variables para Playwright (Render / Docker)
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "/ms-playwright")
os.environ.setdefault("PYPPETEER_HOME", "/ms-playwright")

BASE_DIR = Path(__file__).parent
DESC_DIR = BASE_DIR / "descargas"
DESC_DIR.mkdir(exist_ok=True, parents=True)
LICENSE_CLIENT = LicensingClient()
SESSION_CACHE = BASE_DIR / "session_cache.json"
ENABLE_SESSION_CACHE = os.getenv("ENABLE_SESSION_CACHE", "1").strip().lower() not in {"0", "false", "no"}
PREFERENCES_FILE = BASE_DIR / "user_prefs.json"
RESET_REQUESTS_FILE = BASE_DIR / "password_reset_requests.json"
RESET_TOKEN_TTL = 3600

def _generate_device_fingerprint() -> str:
    raw = f"{platform.node()}|{platform.system()}|{platform.release()}|{uuid.getnode()}"
    return hashlib.sha256(raw.encode()).hexdigest()


def _get_or_init_client_device_id() -> str | None:
    cached_id = st.session_state.get("_device_id")
    if cached_id:
        return str(cached_id)
    device_id = st.query_params.get("device_id")
    if isinstance(device_id, list):
        device_id = device_id[0] if device_id else None
    if device_id:
        st.session_state["_device_id"] = device_id
        return str(device_id)
    # Fallback server-side generation (no JS needed)
    new_id = uuid.uuid4().hex
    st.session_state["_device_id"] = new_id
    st.query_params["device_id"] = new_id
    st.rerun()
    return None


def _require_client_device_id() -> str | None:
    device_id = _get_or_init_client_device_id()
    if device_id:
        return device_id
    st.markdown("<div style='height:120px'></div>", unsafe_allow_html=True)
    st.info("Preparando tu sesión en este equipo...")
    if st.button("Reintentar"):
        st.rerun()
    return None


def _persist_session_state():
    if not ENABLE_SESSION_CACHE:
        return
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
        SESSION_CACHE.write_text(json.dumps(payload))


def _load_cached_session():
    if not ENABLE_SESSION_CACHE:
        return
    if "auth_token" in st.session_state:
        return
    if not SESSION_CACHE.exists():
        return
    try:
        data = json.loads(SESSION_CACHE.read_text())
    except Exception:
        return
    for key, value in data.items():
        st.session_state[key] = value


def _clear_cached_session():
    if not ENABLE_SESSION_CACHE:
        return
    try:
        SESSION_CACHE.unlink(missing_ok=True)
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
    st.session_state["_prefs_loaded"] = True


def _persist_user_preferences():
    data = {
        "download_base_dir": st.session_state.get("download_base_dir", str(DESC_DIR)),
    }
    try:
        PREFERENCES_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception as err:
        st.warning(f"No se pudo guardar la configuracin local: {err}")


def _get_download_base_dir() -> Path:
    base = Path(st.session_state.get("download_base_dir") or str(DESC_DIR)).expanduser()
    base.mkdir(parents=True, exist_ok=True)
    return base


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


def _load_reset_requests() -> dict:
    if RESET_REQUESTS_FILE.exists():
        try:
            with open(RESET_REQUESTS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_reset_requests(data: dict) -> None:
    RESET_REQUESTS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _create_reset_token(email: str) -> str:
    data = _load_reset_requests()
    token = secrets.token_urlsafe(32)
    now = time.time()
    data[token] = {
        "email": email,
        "created": now,
        "expires_at": now + RESET_TOKEN_TTL,
    }
    _save_reset_requests(data)
    return token


def _validate_reset_token(token: str) -> str | None:
    data = _load_reset_requests()
    info = data.get(token)
    if not info:
        return None
    # Compatibilidad con tokens antiguos sin expires_at
    if "expires_at" not in info:
        info["expires_at"] = info.get("created", 0) + RESET_TOKEN_TTL
        data[token] = info
        _save_reset_requests(data)

    if time.time() > info.get("expires_at", 0):
        data.pop(token, None)
        _save_reset_requests(data)
        return None
    return info.get("email")


def _update_local_password(email: str, new_password: str) -> None:
    """
    Actualiza la contraseña en la base local licensing_api.db.
    """
    if not SessionLocal or not lic_crud or not lic_security:
        raise RuntimeError("No se puede acceder a la base de licencias para actualizar la contraseña.")
    db = SessionLocal()
    try:
        user = lic_crud.get_user_by_email(db, email=email)
        if not user:
            raise ValueError("Usuario no encontrado.")
        user.password_hash = lic_security.get_password_hash(new_password)
        db.add(user)
        db.commit()
    finally:
        db.close()


def _discard_reset_token(token: str) -> None:
    data = _load_reset_requests()
    if token in data:
        data.pop(token, None)
        _save_reset_requests(data)


def _send_reset_email_message(email: str, token: str) -> None:
    base_url = os.getenv("APP_BASE_URL", "http://localhost:8501").rstrip("/")
    link = f"{base_url}/?reset_token={token}"
    sender = os.getenv("SMTP_FROM") or os.getenv("SMTP_USER") or "no-reply@example.com"

    msg = EmailMessage()
    msg["Subject"] = "Recupera tu contraseña - SRI Robot"
    msg["From"] = sender
    msg["To"] = email
    msg.set_content(
        f"""Hola,

Hemos recibido una solicitud para restablecer tu contraseña en SRI Robot.
Haz clic en el siguiente enlace para crear una nueva contraseña:
{link}

Si no solicitaste este cambio, ignora este mensaje.
"""
    )
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    if not all([host, user, password]):
        raise RuntimeError("Configura SMTP_HOST, SMTP_PORT, SMTP_USER y SMTP_PASSWORD para enviar correos.")
    use_tls = os.getenv("SMTP_USE_TLS", "1").lower() not in {"0", "false", "no"}
    context = ssl.create_default_context()
    with smtplib.SMTP(host, port, timeout=10) as server:
        if use_tls:
            server.starttls(context=context)
        server.login(user, password)
        server.send_message(msg)


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
    email = _validate_reset_token(token)
    if email:
        st.session_state["password_recovery_mode"] = True
        st.session_state["recovery_email"] = email
        st.session_state["active_reset_token"] = token
    else:
        st.warning("El enlace de recuperación no es válido o ya expiró.")
    st.query_params.clear()


def _render_reset_request():
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 0.9, 1])
    with col:
        logo_html = _logo_html(140)
        if logo_html:
            st.markdown(logo_html, unsafe_allow_html=True)
        st.markdown("<div class='auth-title'>Recuperar contraseña</div>", unsafe_allow_html=True)
        st.info("Ingresa tu correo registrado y te enviaremos un enlace para restablecer tu contraseña.")
        with st.form("password_request_form"):
            email = st.text_input("Correo electrónico", value=st.session_state.get("recovery_email", ""))
            col_send, col_spacer, col_cancel = st.columns([1, 1.4, 1])
            with col_send:
                send = st.form_submit_button("Enviar enlace", type="primary")
            with col_spacer:
                st.write("")
            with col_cancel:
                cancel = st.form_submit_button("Volver al inicio de sesión", type="secondary")
            if cancel:
                st.session_state["reset_request_mode"] = False
                st.session_state.pop("recovery_email", None)
                st.query_params.clear()
                st.rerun()
            if send:
                if not email:
                    st.error("Ingresa el correo registrado.")
                else:
                    try:
                        token = _create_reset_token(email.strip())
                        _send_reset_email_message(email.strip(), token)
                        st.success("Hemos enviado un enlace de recuperación a tu correo.")
                        st.session_state["reset_request_mode"] = False
                        st.session_state["password_recovery_mode"] = False
                        st.query_params.clear()
                        st.rerun()
                    except Exception as err:
                        st.error(f"No se pudo enviar el correo: {err}")


def _render_password_recovery():
    st.session_state.setdefault("password_recovery_mode", False)
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 0.9, 1])
    with col:
        logo_html = _logo_html(140)
        if logo_html:
            st.markdown(logo_html, unsafe_allow_html=True)
        st.markdown("<div class='auth-title'>Reestablecer contraseña</div>", unsafe_allow_html=True)
        st.info("Introduce tu nueva contraseña y confírmala para finalizar el proceso.")
        preset_email = st.session_state.get("recovery_email", "")
        active_token = st.session_state.get("active_reset_token")
        with st.form("password_recovery_form"):
            email = st.text_input("Correo electrónico", value=preset_email, disabled=bool(preset_email))
            new_password = st.text_input("Nueva contraseña", type="password")
            confirm_password = st.text_input("Confirmar contraseña", type="password")
            col_submit, col_back = st.columns([1, 1])
            with col_submit:
                submitted = st.form_submit_button("Guardar contraseña", type="primary")
            with col_back:
                back_to_login = st.form_submit_button("Volver al inicio de sesión", type="secondary")
            if back_to_login:
                st.session_state["password_recovery_mode"] = False
                st.session_state.pop("recovery_email", None)
                st.session_state.pop("active_reset_token", None)
                st.rerun()
            if submitted:
                if not email or not new_password or not confirm_password:
                    st.error("Completa todos los campos.")
                elif new_password != confirm_password:
                    st.error("Las contrase?as no coinciden.")
                else:
                    try:
                        _update_local_password(email.strip(), new_password)
                        st.success("Tu contrase?a se actualiz? correctamente.")
                        if active_token:
                            _discard_reset_token(active_token)
                        st.session_state["password_recovery_mode"] = False
                        st.session_state.pop("recovery_email", None)
                        st.session_state.pop("active_reset_token", None)
                        st.rerun()
                    except Exception as err:
                        st.error(f"No se pudo actualizar la contrase?a: {err}")


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
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 0.9, 1])
    with col:
        logo_html = _logo_html(160)
        if logo_html:
            st.markdown(logo_html, unsafe_allow_html=True)
        st.markdown("<div class='auth-title'>Iniciar sesión</div>", unsafe_allow_html=True)
        with st.form("login_form"):
            email = st.text_input("Correo electrónico")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar sesión", type="primary")
        st.markdown(
            "<div style='display:flex; justify-content:flex-end; margin-top:8px;'>"
            "<a href='?reset_request=1' style='color:#77aaff;text-decoration:underline;font-size:0.9rem;'>¿Olvidaste tu contraseña?</a>"
            "</div>",
            unsafe_allow_html=True,
        )
    if submitted:
        if not email or not password:
            st.error("Completa todos los campos.")
        else:
            try:
                token = LICENSE_CLIENT.login(email.strip(), password)
                st.session_state["auth_token"] = token
                st.session_state["user_email"] = email.strip()
                st.session_state["license_validated"] = False
                _persist_session_state()
                st.success("Inicio de sesión exitoso. Continúa con la activación.")
            except Exception as err:
                st.error(f"Error al autenticar: {err}")


def _render_activation():
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    _, center_col, _ = st.columns([1, 0.9, 1])
    with center_col:
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
                    except Exception as err:
                        st.error(f"No se pudo activar la licencia: {err}")
def _ensure_access():
    _load_cached_session()
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

    if not st.session_state.get("license_validated"):
        try:
            LICENSE_CLIENT.validate_license(st.session_state["auth_token"], fingerprint)
            st.session_state["license_validated"] = True
            st.session_state["license_last_check"] = time.time()
            _persist_session_state()
        except Exception:
            _render_activation()
            st.stop()

    # Desactivamos la validación periódica para que la licencia siga activa
    # hasta que se elimine manualmente el registro del usuario.


_ensure_access()
_load_user_preferences()
DEVICE_FINGERPRINT = st.session_state.get("device_fingerprint") or st.session_state.get("user_email")

# ==============================
# SIDEBAR CORPORATIVO
# ==============================
with st.sidebar:
    user_email = st.session_state.get("user_email") or "No disponible"
    with st.expander("Perfil", expanded=False):
        st.markdown("**Usuario conectado**")
        st.caption(user_email)
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            _clear_cached_session()

    logo_path = Path(__file__).parent / "LogoAUDIT.png"
    if logo_path.exists():
        st.image(str(logo_path), width=180)
    st.markdown("###  Auditora Web SRI Robot")
    st.write("Automatiza descargas, valida comprobantes y genera reportes tributarios.")
    st.markdown("---")
    st.markdown("**Versión:** 2.0  \n**Actualizado:** Noviembre 2025")

# ==============================
# INTERFAZ PRINCIPAL
# ==============================
st.title(" SRI Robot Audit  Descarga y Reporte Automático")

tab1, tab2 = st.tabs([" Descarga de Comprobantes", " Reportes e Historial"])

# =====================================================
# TAB 1  DESCARGA Y PROCESAMIENTO AUTOMTICO
# =====================================================
with tab1:
    st.markdown("#### Ingreso de Credenciales y Filtros")

    col_base1, col_base2 = st.columns([2, 2])
    with col_base1:
        ruc = st.text_input("RUC", placeholder="Ejemplo: 0999999001")
        ci_adicional_input = ""
        clave = st.text_input("Clave del SRI", type="password", placeholder="********")

    with col_base2:
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
    anio_recibidos = datetime.now().year
    mes_recibidos = datetime.now().month
    dia_recibidos = 0

    if origen == "Recibidos":
        col_r1, col_r2, col_r3 = st.columns([1, 1, 1])
        with col_r1:
            anio_recibidos = st.number_input(
                "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
            )
        with col_r2:
            mes_recibidos = st.number_input(
                "Mes (1-12)", min_value=1, max_value=12, value=datetime.now().month, step=1
            )
        with col_r3:
            dia_recibidos = st.number_input(
                "Día (0 = Todos)", min_value=0, max_value=31, value=0, step=1,
                help="Elige 0 para descargar todo el mes o un da específico (1-31).",
            )
        formatos = st.multiselect("Formatos a descargar", ["XML", "PDF"], default=["XML", "PDF"])
    else:
        col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
        col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
        with col_f1:
            anio_emitidos = st.number_input(
                "Año", min_value=2015, max_value=datetime.now().year, value=datetime.now().year, step=1
            )
        with col_f2:
            mes_emitidos = st.number_input(
                "Mes (1-12)", min_value=1, max_value=12, value=datetime.now().month, step=1
            )
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
            "Estado de autorización", ["Autorizados", "No autorizados", "Por procesar"], index=0
        )
        col_e1, col_e2 = st.columns([1, 1])
        with col_e1:
            establecimiento_input = st.text_input(
                "Establecimiento",
                value="Todos",
                help="Escribe 'Todos' o un número de 3 dgitos (ej. 001).",
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
    st.markdown("---")
    st.markdown("#### Carpeta base donde se guardarán las descargas")
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
            st.error(error)
    st.caption(
        f"Carpeta activa: `{st.session_state.get('download_base_dir', str(DESC_DIR))}`. Dentro se almacenarán tus descargas."
    )

    start_clicked = st.button(" Iniciar proceso", use_container_width=True, type="primary", key="start_process")
    stop_clicked = st.button(" Detener proceso", use_container_width=True, key="stop_process")

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
            if origen == "Recibidos":
                formatos_final = formatos
                if not formatos_final:
                    st.warning("Selecciona al menos un formato (XML o PDF).")
                    st.stop()
                anio_val = int(anio_recibidos)
                mes_val = int(mes_recibidos)
                dia_val = int(dia_recibidos)
                fecha_emitidos_val = None
                estado_emitidos_val = None
                establecimiento_val = None
                punto_emision_val = None
            else:
                anio_val = int(anio_emitidos)
                mes_val = int(mes_emitidos)
                dia_val = int(dia_emitidos)
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
                        st.error("El establecimiento debe ser 'Todos' o un número de tres dgitos (ej. 001).")
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

            base_descargas = _get_download_base_dir()
            destino = base_descargas / ruc
            destino.mkdir(parents=True, exist_ok=True)

            st.session_state.download_messages = []
            st.session_state.download_result = None
            st.session_state.download_error = None
            st.session_state.download_registered = False
            st.session_state.download_status = "running"
            st.session_state.running_notice_ts = time.time()
            params = {
                "ruc": ruc,
                "clave": clave,
                "anio": anio_val,
                "mes": mes_val,
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
            st.session_state.download_params = params
            worker = threading.Thread(
                target=_download_worker,
                args=(params, st.session_state.download_queue),
                daemon=True,
            )
            st.session_state.download_thread = worker
            worker.start()
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

    if st.session_state.download_status in {"done", "error"} and st.session_state.download_params:
        st.session_state.running_notice_ts = None
        st.session_state.stop_notice_ts = None
        if st.session_state.download_error:
            if "Proceso cancelado por el usuario" in st.session_state.download_error:
                st.warning("Proceso cancelado por el usuario.")
            else:
                st.error(f"Ocurrió un error inesperado: {st.session_state.download_error}")
        resultado = st.session_state.download_result or {}
        params = st.session_state.download_params
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
                carpeta_tipo = Path(resultado.get("carpeta_tipo") or params.get("destino"))
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
            else:
                n_xml = resultado.get("n_xml", 0)
                n_pdf = resultado.get("n_pdf", 0)
                tipo_visible = resultado.get("tipo_visible", params.get("tipo"))
                carpeta_tipo = Path(resultado.get("carpeta_tipo") or params.get("destino"))
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
                if n_xml > 0:
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

                zip_target = carpeta_tipo
                zip_path = zip_target.with_suffix(".zip")
                if zip_path.exists():
                    zip_path.unlink()
                shutil.make_archive(str(zip_target), "zip", zip_target)
                if zip_path.exists():
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            " Descargar ZIP de la carpeta",
                            f,
                            file_name=zip_path.name,
                            use_container_width=True,
                        )

    if st.session_state.download_status in {"running", "cancelling"}:
        time.sleep(0.6)
        try:
            st.experimental_rerun()
        except Exception:
            pass

with tab2:
    st.markdown("####  Historial de ejecuciones recientes")
    historial = obtener_historial(DEVICE_FINGERPRINT)
    historial_raw = historial.copy()

    #  Evitar error valor de verdad de un DataFrame es ambiguo
    if isinstance(historial, pd.DataFrame) and not historial.empty:
        st.markdown(
            """
            <style>
            .historial-table { width: 100%; overflow-x: auto; }
            .historial-table table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                border-radius: 10px;
                overflow: hidden;
                font-size: 0.88rem;
            }
            .historial-table th,
            .historial-table td {
                text-align: center;
                padding: 8px 10px;
                border: 1px solid rgba(120, 120, 120, 0.35);
                vertical-align: middle;
                white-space: nowrap;
            }
            .historial-table thead th {
                font-weight: 600;
                text-transform: none;
                letter-spacing: 0.2px;
            }
            .historial-table tbody tr:hover td {
                background: rgba(30, 74, 168, 0.10);
            }
            @media (prefers-color-scheme: dark) {
                .historial-table thead th {
                    background: rgba(30, 74, 168, 0.25);
                }
                .historial-table td {
                    background: rgba(12, 15, 22, 0.35);
                }
            }
            @media (prefers-color-scheme: light) {
                .historial-table thead th {
                    background: rgba(30, 74, 168, 0.12);
                }
                .historial-table td {
                    background: rgba(255, 255, 255, 0.9);
                }
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
            columns=["n_registros", "n_xml", "n_pdf", "reporte", "reporte_xml", "reporte_pdf", "xml_dir", "pdf_dir"],
            errors="ignore",
        )

        historial.insert(0, "No.", range(1, len(historial) + 1))
        historial = historial.rename(columns=columnas_amigables)

        tabla_html = historial.to_html(
            index=False,
            escape=False,
            border=0,
            classes="historial-table-grid",
        )
        st.markdown(f"<div class='historial-table'>{tabla_html}</div>", unsafe_allow_html=True)
        st.success(f" Total de operaciones registradas: {len(historial)}")

        descargables = []
        if isinstance(historial_raw, pd.DataFrame):
            hist_reset = historial_raw.reset_index(drop=True)
            for idx, fila in hist_reset.iterrows():
                ruta_val = fila.get("reporte") or fila.get("reporte_xml") or fila.get("reporte_pdf")
                if not ruta_val or not isinstance(ruta_val, (str, Path)):
                    continue
                ruta_path = Path(ruta_val).expanduser()
                if not ruta_path.exists():
                    continue
                etiqueta = f"{idx + 1}. {ruta_path.name} ({fila.get('timestamp', '')})"
                descargables.append((etiqueta, ruta_path))

        if descargables:
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
