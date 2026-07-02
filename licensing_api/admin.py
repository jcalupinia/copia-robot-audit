"""Panel administrativo web para la API de licencias.

Vive en /admin dentro de la misma FastAPI app y permite administrar
usuarios y licencias desde una interfaz HTML simple (sin Shell de Render).

Auth (rework 2026-06-21):
  - Login contra la BD: buscar user por email + verificar password_hash.
  - Requisitos para entrar: user.is_active=True Y user.role=="admin".
  - Ya NO se usan ADMIN_EMAIL/ADMIN_PASSWORD del env — el acceso es 100% DB.
  - Sesion en cookie httponly + samesite=lax firmada con el mismo
    JWT_SECRET_KEY que el resto de la API (reusa security.create_access_token).
  - El "subject" del JWT es `admin:<user_id>` (int estable, no email que
    podria cambiar).
  - TTL: 4 horas.
  - En cada request se re-valida contra la DB (por si cambio el rol o se
    desactivo el usuario).

Funcionalidad:
  - Dashboard con tabla de usuarios + rol + licencias con estado y expiracion.
  - Crear usuario nuevo con rol seleccionable (admin/operador/cliente).
  - Agregar licencia con duracion (1/3/6/12 meses, fecha manual, sin venc.).
  - Editar fecha de expiracion de una licencia existente.
  - Renovar licencia (extender expires_at).
  - Cambiar rol de un usuario.
  - Desactivar / reactivar licencia (sin borrar).
  - Borrar usuario (cascada a licencias y dispositivos).
  - Resetear contrasena de un usuario.

Reglas de seguridad:
  - Un admin NO puede cambiar su propio rol (evita bloqueo accidental).
  - Un admin NO puede borrarse a si mismo.
  - No se puede degradar/borrar/desactivar al ULTIMO admin activo.
  - Todas las rutas /admin/* salvo /admin/login chequean sesion + rol='admin'.

Diseno:
  - HTML/CSS inline en strings de Python (sin jinja2, sin assets externos).
  - Paleta similar a la landing: dark `#070a12` con acentos verde y azul.
  - Responsive simple (form-row con grid auto-fit).
"""

from __future__ import annotations

import html
from datetime import datetime, timedelta
from typing import Optional
from urllib.parse import quote

from fastapi import APIRouter, Depends, Form, Request, Response, status
from fastapi.responses import HTMLResponse, RedirectResponse
from sqlalchemy.orm import Session

from licensing_api import crud, models
from licensing_api.database import get_db
from licensing_api.security import (
    create_access_token,
    decode_access_token,
    get_password_hash,
    verify_password,
)


router = APIRouter(prefix="/admin", tags=["admin"])

ADMIN_COOKIE_NAME = "admin_session"
ADMIN_TOKEN_TTL = timedelta(hours=4)


# ============================================================
# AUTH HELPERS (DB-based)
# ============================================================

def _get_authenticated_admin(
    request: Request, db: Session
) -> Optional[models.User]:
    """Devuelve el user admin de la sesion si es valida, o None.

    Chequea:
      1) cookie admin_session presente
      2) JWT valido con sub="admin:<user_id>"
      3) user existe en DB
      4) user.is_active
      5) user.role == "admin"

    En cada request se re-valida contra la DB (por si el rol cambio o el
    user fue desactivado desde otro admin).
    """
    token = request.cookies.get(ADMIN_COOKIE_NAME, "") or ""
    if not token:
        return None
    sub = decode_access_token(token)
    if not sub or not sub.startswith("admin:"):
        return None
    try:
        user_id = int(sub.split(":", 1)[1])
    except (ValueError, IndexError):
        return None
    user = crud.get_user_by_id(db, user_id)
    if not user or not user.is_active or user.role != "admin":
        return None
    return user


def _redirect_to_login(msg: Optional[str] = None) -> RedirectResponse:
    url = "/admin/login"
    if msg:
        url = f"{url}?error={quote(msg)}"
    return RedirectResponse(url=url, status_code=status.HTTP_303_SEE_OTHER)


def _is_secure_request(request: Request) -> bool:
    proto = (request.headers.get("x-forwarded-proto", "") or "").strip().lower()
    if proto == "https":
        return True
    return request.url.scheme == "https"


def _set_admin_cookie(
    response: Response, request: Request, user_id: int
) -> None:
    token = create_access_token(
        f"admin:{user_id}",
        expires_delta=ADMIN_TOKEN_TTL,
    )
    response.set_cookie(
        key=ADMIN_COOKIE_NAME,
        value=token,
        max_age=int(ADMIN_TOKEN_TTL.total_seconds()),
        httponly=True,
        secure=_is_secure_request(request),
        samesite="lax",
        path="/",
    )


def _clear_admin_cookie(response: Response) -> None:
    response.delete_cookie(ADMIN_COOKIE_NAME, path="/")


def _redirect_dashboard(
    success: Optional[str] = None,
    error: Optional[str] = None,
) -> RedirectResponse:
    qs = ""
    if success:
        qs = f"?success={quote(success)}"
    elif error:
        qs = f"?error={quote(error)}"
    return RedirectResponse(
        url=f"/admin{qs}",
        status_code=status.HTTP_303_SEE_OTHER,
    )


# ============================================================
# HTML LAYOUT
# ============================================================

_CSS = """
:root {
    --bg: #070a12;
    --panel: #0f1a2e;
    --panel-strong: #15243d;
    --border: #1e2b45;
    --text: #e7eaf0;
    --text-muted: #94a1bd;
    --primary: #10b981;
    --primary-strong: #34d399;
    --accent: #60a5fa;
    --danger: #ef4444;
    --warn: #f59e0b;
}
* { box-sizing: border-box; }
body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background: var(--bg);
    background-image:
        radial-gradient(60rem 30rem at 90% -10%, rgba(91,140,255,0.10), transparent 60%),
        radial-gradient(50rem 30rem at -10% 10%, rgba(34,197,94,0.10), transparent 60%);
    color: var(--text);
    min-height: 100vh;
    line-height: 1.5;
}
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }
.container { max-width: 1300px; margin: 0 auto; padding: 24px; }
.header {
    display: flex; align-items: center; justify-content: space-between;
    border-bottom: 1px solid var(--border); padding: 16px 24px;
    background: rgba(15,26,46,0.5); backdrop-filter: blur(6px);
}
.brand { display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 18px; color: var(--text); }
.brand .accent { color: var(--primary); }
.nav-actions { display: flex; gap: 4px; align-items: center; }
.nav-actions .who { color: var(--text-muted); font-size: 13px; margin-right: 8px; }
.nav-actions a, .nav-actions button {
    color: var(--text-muted); background: transparent; border: 1px solid transparent;
    text-decoration: none; padding: 8px 14px; border-radius: 8px; font: inherit; cursor: pointer;
}
.nav-actions a:hover, .nav-actions button:hover {
    color: var(--text); border-color: var(--border); text-decoration: none;
}
h1 { font-size: 28px; margin: 0 0 8px; font-weight: 600; }
h2 { font-size: 18px; margin: 32px 0 16px; color: var(--text); font-weight: 600; }
.muted { color: var(--text-muted); font-size: 14px; }
.card {
    background: var(--panel); border: 1px solid var(--border); border-radius: 14px;
    padding: 20px; margin-bottom: 20px;
}
.alert {
    border-radius: 10px; padding: 12px 16px; margin-bottom: 16px;
    border: 1px solid var(--border); font-size: 14px;
}
.alert.success {
    background: rgba(16,185,129,0.10); border-color: rgba(16,185,129,0.4);
    color: var(--primary-strong);
}
.alert.error {
    background: rgba(239,68,68,0.10); border-color: rgba(239,68,68,0.4);
    color: #fca5a5;
}
.alert.info {
    background: rgba(96,165,250,0.10); border-color: rgba(96,165,250,0.4);
    color: #93c5fd;
}
.form-row {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 14px; margin-bottom: 16px;
}
label {
    display: block; font-size: 13px; color: var(--text-muted);
    margin-bottom: 6px; font-weight: 500;
}
input[type=text], input[type=email], input[type=password], input[type=date], input[type=number], select {
    width: 100%; padding: 10px 12px; background: rgba(7,10,18,0.7);
    border: 1px solid var(--border); border-radius: 8px;
    color: var(--text); font: inherit;
}
select { cursor: pointer; }
input:focus, select:focus {
    outline: none; border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(16,185,129,0.15);
}
.btn {
    display: inline-flex; align-items: center; gap: 8px; padding: 10px 18px;
    border-radius: 10px; border: 1px solid transparent; font: inherit; cursor: pointer;
    text-decoration: none; transition: filter 0.15s ease, transform 0.05s ease;
}
.btn:active { transform: translateY(1px); }
.btn-primary {
    background: linear-gradient(135deg, #10b981 0%, #34d399 100%);
    color: #052e1e; font-weight: 600;
}
.btn-primary:hover { filter: brightness(1.08); }
.btn-ghost {
    background: transparent; border-color: var(--border); color: var(--text-muted);
}
.btn-ghost:hover { color: var(--text); border-color: var(--primary); }
.btn-danger {
    background: rgba(239,68,68,0.15); border-color: rgba(239,68,68,0.4); color: #fca5a5;
}
.btn-danger:hover { background: rgba(239,68,68,0.25); }
.btn-warn {
    background: rgba(245,158,11,0.15); border-color: rgba(245,158,11,0.4); color: #fcd34d;
}
.btn-warn:hover { background: rgba(245,158,11,0.25); }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td {
    padding: 12px 10px; text-align: left;
    border-bottom: 1px solid var(--border); vertical-align: top;
}
th {
    font-weight: 600; color: var(--text-muted); font-size: 12px;
    text-transform: uppercase; letter-spacing: 0.04em;
    background: rgba(7,10,18,0.3);
}
tbody tr:hover { background: rgba(16,185,129,0.03); }
.pill {
    display: inline-block; padding: 2px 10px; border-radius: 999px;
    font-size: 12px; font-weight: 500; border: 1px solid var(--border);
    white-space: nowrap;
}
.pill.ok { background: rgba(16,185,129,0.12); border-color: rgba(16,185,129,0.4); color: var(--primary-strong); }
.pill.bad { background: rgba(239,68,68,0.12); border-color: rgba(239,68,68,0.4); color: #fca5a5; }
.pill.warn { background: rgba(245,158,11,0.12); border-color: rgba(245,158,11,0.4); color: #fcd34d; }
.pill.info { background: rgba(96,165,250,0.12); border-color: rgba(96,165,250,0.4); color: #93c5fd; }
.pill.role-admin { background: rgba(96,165,250,0.15); border-color: rgba(96,165,250,0.5); color: #93c5fd; }
.pill.role-operador { background: rgba(16,185,129,0.12); border-color: rgba(16,185,129,0.4); color: var(--primary-strong); }
.pill.role-cliente { background: rgba(148,163,189,0.12); border-color: rgba(148,163,189,0.4); color: var(--text-muted); }
.licencias-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 8px; }
.licencias-list li {
    background: rgba(7,10,18,0.5); border: 1px solid var(--border);
    border-radius: 8px; padding: 10px 12px;
}
.licencias-list .code {
    font-family: "SFMono-Regular", Consolas, Menlo, monospace;
    font-size: 11px; color: var(--accent); word-break: break-all;
}
.inline-form { display: inline-block; margin: 0; }
.actions { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }
.actions .btn { padding: 4px 8px; font-size: 11px; }
.actions select { width: auto; padding: 4px 8px; font-size: 12px; }
.actions input[type=number], .actions input[type=date] { width: 110px; padding: 4px 8px; font-size: 12px; }
.center-page { min-height: 80vh; display: grid; place-items: center; }
.login-card { max-width: 400px; width: 100%; padding: 32px; }
.empty { padding: 32px; text-align: center; color: var(--text-muted); }
.days-left { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
.days-left.warn { color: #fcd34d; }
.days-left.bad { color: #fca5a5; }
@media (max-width: 720px) {
    .header { padding: 12px 16px; }
    .container { padding: 16px; }
    table { font-size: 11px; }
    th, td { padding: 8px 6px; }
}
"""


def _layout(title: str, body: str, *, show_header: bool = True, current_admin_email: str = "") -> str:
    title_safe = html.escape(title)
    header_html = ""
    if show_header:
        who = ""
        if current_admin_email:
            who = f'<span class="who">👤 {html.escape(current_admin_email)}</span>'
        header_html = f"""
        <div class="header">
            <div class="brand">ROBOT&nbsp;AUDIT<span class="accent">&nbsp;SRI</span>&nbsp;· Admin</div>
            <div class="nav-actions">
                {who}
                <a href="/admin">Dashboard</a>
                <a href="/admin/logout">Cerrar sesi&oacute;n</a>
            </div>
        </div>
        """
    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title_safe} — Admin · ROBOT AUDIT SRI</title>
<meta name="robots" content="noindex, nofollow">
<style>{_CSS}</style>
</head>
<body>
{header_html}
{body}
</body>
</html>"""


# ============================================================
# ROUTES — login / logout
# ============================================================

@router.get("/login", response_class=HTMLResponse)
def admin_login_page(
    request: Request,
    db: Session = Depends(get_db),
    error: Optional[str] = None,
):
    # Si ya hay sesion admin valida, ir directo al dashboard.
    if _get_authenticated_admin(request, db):
        return RedirectResponse(
            url="/admin", status_code=status.HTTP_303_SEE_OTHER,
        )

    error_html = ""
    if error:
        error_html = f'<div class="alert error">{html.escape(error)}</div>'

    body = f"""
    <div class="container center-page">
        <div class="card login-card">
            <h1 style="font-size:22px; margin-bottom:4px">Panel administrativo</h1>
            <p class="muted" style="margin-bottom:20px">
                Solo usuarios con rol de administrador pueden ingresar.
            </p>
            {error_html}
            <form method="post" action="/admin/login">
                <div style="margin-bottom:14px">
                    <label for="email">Email</label>
                    <input type="email" name="email" id="email" required autocomplete="username">
                </div>
                <div style="margin-bottom:20px">
                    <label for="password">Contraseña</label>
                    <input type="password" name="password" id="password" required autocomplete="current-password">
                </div>
                <button type="submit" class="btn btn-primary" style="width:100%; justify-content:center">
                    Ingresar
                </button>
            </form>
        </div>
    </div>
    """
    return HTMLResponse(content=_layout("Login", body, show_header=False))


@router.post("/login")
def admin_login(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    password: str = Form(...),
):
    email_n = (email or "").strip().lower()
    password_n = password or ""

    user = crud.get_user_by_email(db, email=email_n)
    if not user or not verify_password(password_n, user.password_hash):
        return _redirect_to_login("Credenciales incorrectas.")
    if not user.is_active:
        return _redirect_to_login("Usuario inactivo.")
    if user.role != "admin":
        return _redirect_to_login(
            "Este usuario no tiene permisos de administrador."
        )

    response = RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)
    _set_admin_cookie(response, request, user.id)
    return response


@router.get("/logout")
def admin_logout():
    response = RedirectResponse(
        url="/admin/login", status_code=status.HTTP_303_SEE_OTHER,
    )
    _clear_admin_cookie(response)
    return response


# ============================================================
# ROUTES — dashboard
# ============================================================

_DURACION_PRESETS = {
    "1m": ("1 mes", 30),
    "3m": ("3 meses", 90),
    "6m": ("6 meses", 180),
    "12m": ("1 año", 365),
    "none": ("Sin vencimiento", None),
    "manual": ("Fecha manual", None),
}


@router.get("/", response_class=HTMLResponse)
def admin_dashboard(
    request: Request,
    db: Session = Depends(get_db),
    success: Optional[str] = None,
    error: Optional[str] = None,
):
    current_admin = _get_authenticated_admin(request, db)
    if not current_admin:
        return _redirect_to_login()

    users = (
        db.query(models.User)
        .order_by(models.User.created_at.desc())
        .all()
    )

    success_html = (
        f'<div class="alert success">{html.escape(success)}</div>' if success else ""
    )
    error_html = (
        f'<div class="alert error">{html.escape(error)}</div>' if error else ""
    )

    # === Tabla de usuarios ===
    rows_html_parts: list[str] = []
    for u in users:
        is_self = (u.id == current_admin.id)

        # Licencias del usuario
        lic_items: list[str] = []
        for lic in (u.licenses or []):
            estado = crud.get_license_status(lic)
            if estado == "activa":
                estado_cls = "ok"
                estado_txt = "activa"
            elif estado == "vencida":
                estado_cls = "bad"
                estado_txt = "vencida"
            else:
                estado_cls = "warn"
                estado_txt = "inactiva"

            act_text = (
                f"Activada: {lic.activated_at.strftime('%Y-%m-%d %H:%M')}"
                if lic.activated_at
                else "<em>Sin activar</em>"
            )
            exp_text = ""
            days_left_html = ""
            if lic.expires_at:
                exp_text = f" · Expira: {lic.expires_at.strftime('%Y-%m-%d')}"
                dr = crud.days_remaining(lic)
                if dr is not None:
                    if dr < 0:
                        days_left_html = f'<div class="days-left bad">Vencida hace {-dr} d\xEDa(s)</div>'
                    elif dr <= 15:
                        days_left_html = f'<div class="days-left warn">Quedan {dr} d\xEDa(s)</div>'
                    else:
                        days_left_html = f'<div class="days-left">Quedan {dr} d\xEDa(s)</div>'
            else:
                exp_text = " · Sin vencimiento"

            toggle_label = "Desactivar" if lic.is_active else "Reactivar"
            toggle_cls = "btn-warn" if lic.is_active else "btn-primary"

            lic_items.append(f"""
            <li>
                <div class="code">{html.escape(lic.code or '')}</div>
                <div style="margin-top:6px; display:flex; gap:8px; align-items:center; flex-wrap:wrap">
                    <span class="pill {estado_cls}">{estado_txt}</span>
                    <span class="muted">{act_text}{exp_text}</span>
                </div>
                {days_left_html}
                <div class="actions" style="margin-top:8px">
                    <form method="post" action="/admin/licenses/{lic.id}/toggle" class="inline-form">
                        <button type="submit" class="btn {toggle_cls}">{toggle_label}</button>
                    </form>
                    <form method="post" action="/admin/licenses/{lic.id}/set-expiry" class="inline-form">
                        <input type="date" name="expires_at" value="{lic.expires_at.strftime('%Y-%m-%d') if lic.expires_at else ''}" title="Nueva fecha de expiración (vacío = sin vencimiento)">
                        <button type="submit" class="btn btn-ghost">Editar exp.</button>
                    </form>
                    <form method="post" action="/admin/licenses/{lic.id}/renew" class="inline-form">
                        <input type="number" name="additional_days" min="1" max="3650" value="365" title="Días a agregar">
                        <button type="submit" class="btn btn-primary">Renovar</button>
                    </form>
                </div>
            </li>
            """)

        licencias_html = (
            f'<ul class="licencias-list">{"".join(lic_items)}</ul>'
            if lic_items
            else '<span class="muted">Sin licencias</span>'
        )

        estado_user_pill = "ok" if u.is_active else "bad"
        estado_user_txt = "activo" if u.is_active else "inactivo"
        role_cls = f"role-{u.role}"
        fecha = u.created_at.strftime("%Y-%m-%d %H:%M") if u.created_at else "-"
        email_safe = html.escape(u.email or "")
        email_attr = html.escape(u.email or "", quote=True)

        # Selector de rol (disable si es self)
        role_options = ""
        for r in ("admin", "operador", "cliente"):
            sel = " selected" if u.role == r else ""
            role_options += f'<option value="{r}"{sel}>{r}</option>'
        role_form = f"""
        <form method="post" action="/admin/users/{u.id}/set-role" class="inline-form">
            <select name="role" {'disabled title="No podés cambiar tu propio rol"' if is_self else ''}>{role_options}</select>
            <button type="submit" class="btn btn-ghost" {'disabled' if is_self else ''}>Cambiar rol</button>
        </form>
        """

        delete_form = ""
        if not is_self:
            delete_form = f"""
            <form method="post" action="/admin/users/{u.id}/delete" class="inline-form"
                  onsubmit="return confirm('¿Borrar al usuario {email_attr} y todas sus licencias? Esta accion es IRREVERSIBLE.');">
                <button type="submit" class="btn btn-danger">Borrar</button>
            </form>
            """

        rows_html_parts.append(f"""
        <tr>
            <td>
                <strong>{email_safe}</strong>
                {' <span class="pill info" style="margin-left:6px">Vos</span>' if is_self else ''}
                <br><span class="muted">ID {u.id}</span>
            </td>
            <td><span class="pill {role_cls}">{html.escape(u.role or 'operador')}</span></td>
            <td><span class="pill {estado_user_pill}">{estado_user_txt}</span></td>
            <td>{fecha}</td>
            <td>{licencias_html}</td>
            <td>
                <div class="actions">
                    {role_form}
                    {delete_form}
                </div>
            </td>
        </tr>
        """)

    table_body = (
        "".join(rows_html_parts)
        if rows_html_parts
        else '<tr><td colspan="6" class="empty">No hay usuarios todavía.</td></tr>'
    )

    # Presets de duracion para el form de crear licencia
    dur_options = "".join(
        f'<option value="{k}">{label}</option>'
        for k, (label, _days) in _DURACION_PRESETS.items()
    )

    body = f"""
    <div class="container">
        <h1>Dashboard</h1>
        <p class="muted">Gestión de usuarios, licencias y roles.</p>

        {success_html}
        {error_html}

        <div class="card">
            <h2 style="margin-top:0">Crear usuario nuevo</h2>
            <p class="muted">Opcional: incluí un código de licencia y su duración para crearla en el mismo paso.</p>
            <form method="post" action="/admin/users/create">
                <div class="form-row">
                    <div>
                        <label for="new_email">Email</label>
                        <input type="email" name="email" id="new_email" required>
                    </div>
                    <div>
                        <label for="new_password">Contraseña</label>
                        <input type="text" name="password" id="new_password" required>
                    </div>
                    <div>
                        <label for="new_role">Rol</label>
                        <select name="role" id="new_role">
                            <option value="cliente" selected>Cliente</option>
                            <option value="operador">Operador</option>
                            <option value="admin">Admin</option>
                        </select>
                    </div>
                </div>
                <div class="form-row">
                    <div>
                        <label for="new_license">Código de licencia (opcional)</label>
                        <input type="text" name="license_code" id="new_license" placeholder="AUDIT-XXXX-YYYY">
                    </div>
                    <div>
                        <label for="new_license_dur">Duración de la licencia</label>
                        <select name="license_duration" id="new_license_dur">
                            {dur_options}
                        </select>
                    </div>
                    <div>
                        <label for="new_license_manual">Fecha manual (si eligió "Fecha manual")</label>
                        <input type="date" name="license_manual_date" id="new_license_manual">
                    </div>
                </div>
                <button type="submit" class="btn btn-primary">Crear usuario</button>
            </form>
        </div>

        <div class="card">
            <h2 style="margin-top:0">Agregar licencia a usuario existente</h2>
            <form method="post" action="/admin/licenses/create">
                <div class="form-row">
                    <div>
                        <label for="lic_user_email">Email del usuario</label>
                        <input type="email" name="email" id="lic_user_email" required>
                    </div>
                    <div>
                        <label for="lic_code">Código de licencia</label>
                        <input type="text" name="code" id="lic_code" required placeholder="AUDIT-XXXX-YYYY">
                    </div>
                    <div>
                        <label for="lic_dur">Duración</label>
                        <select name="license_duration" id="lic_dur">
                            {dur_options}
                        </select>
                    </div>
                    <div>
                        <label for="lic_manual">Fecha manual (si aplica)</label>
                        <input type="date" name="license_manual_date" id="lic_manual">
                    </div>
                </div>
                <button type="submit" class="btn btn-primary">Agregar licencia</button>
            </form>
        </div>

        <div class="card">
            <h2 style="margin-top:0">Resetear contraseña</h2>
            <p class="muted">El hash se guarda con <code>get_password_hash</code> — la contraseña nunca se guarda en claro.</p>
            <form method="post" action="/admin/users/reset-password">
                <div class="form-row">
                    <div>
                        <label for="reset_email">Email del usuario</label>
                        <input type="email" name="email" id="reset_email" required>
                    </div>
                    <div>
                        <label for="reset_password">Nueva contraseña</label>
                        <input type="text" name="new_password" id="reset_password" required>
                    </div>
                </div>
                <button type="submit" class="btn btn-warn">Actualizar contraseña</button>
            </form>
        </div>

        <h2>Usuarios y licencias</h2>
        <div class="card" style="padding:0; overflow:auto">
            <table>
                <thead>
                    <tr>
                        <th>Usuario</th>
                        <th>Rol</th>
                        <th>Estado</th>
                        <th>Creado</th>
                        <th>Licencias</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    {table_body}
                </tbody>
            </table>
        </div>
    </div>
    """
    return HTMLResponse(content=_layout("Dashboard", body, current_admin_email=current_admin.email))


# ============================================================
# HELPER: parseo de duracion desde form
# ============================================================

def _parse_expiration_from_form(
    duration_key: str,
    manual_date_str: str,
) -> tuple[Optional[datetime], Optional[int]]:
    """Convierte los campos del form (duration_key + manual_date) en
    (expires_at, validity_days) para pasarle a crud.create_license.

    Retorna (None, None) si "sin vencimiento".
    """
    duration_key = (duration_key or "none").strip()
    manual_date_str = (manual_date_str or "").strip()

    if duration_key == "manual":
        if not manual_date_str:
            raise ValueError("Elegiste 'Fecha manual' pero no ingresaste ninguna fecha.")
        try:
            expires_at = datetime.strptime(manual_date_str, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"Fecha manual inválida: {manual_date_str!r} (formato esperado YYYY-MM-DD).")
        # Setear a fin del dia para que el user tenga hasta las 23:59:59
        expires_at = expires_at.replace(hour=23, minute=59, second=59)
        return expires_at, None

    if duration_key == "none":
        return None, None

    preset = _DURACION_PRESETS.get(duration_key)
    if not preset:
        raise ValueError(f"Duración inválida: {duration_key!r}.")
    _, days = preset
    if days is None:
        return None, None
    return None, int(days)


# ============================================================
# ROUTES — acciones
# ============================================================

@router.post("/users/create")
def admin_users_create(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    password: str = Form(...),
    role: str = Form("cliente"),
    license_code: str = Form(""),
    license_duration: str = Form("none"),
    license_manual_date: str = Form(""),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    email_n = (email or "").strip().lower()
    password_n = password or ""
    code_n = (license_code or "").strip()
    role_n = (role or "cliente").strip().lower()

    if not email_n or not password_n:
        return _redirect_dashboard(error="Email y contraseña son obligatorios.")
    if role_n not in crud.VALID_ROLES:
        return _redirect_dashboard(error=f"Rol inválido: {role_n!r}.")

    if crud.get_user_by_email(db, email=email_n):
        return _redirect_dashboard(error=f"El usuario {email_n} ya existe.")

    try:
        user = crud.create_user(db, email=email_n, password=password_n, role=role_n)
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(error=f"No se pudo crear el usuario: {exc}")

    if code_n:
        if db.query(models.License).filter(models.License.code == code_n).first():
            try:
                db.delete(user)
                db.commit()
            except Exception:
                db.rollback()
            return _redirect_dashboard(
                error=f"El código '{code_n}' ya está en uso. Usuario NO creado."
            )
        try:
            expires_at, validity_days = _parse_expiration_from_form(
                license_duration, license_manual_date,
            )
        except ValueError as exc:
            db.delete(user)
            db.commit()
            return _redirect_dashboard(error=str(exc))
        try:
            crud.create_license(
                db, user=user, code=code_n,
                expires_at=expires_at, validity_days=validity_days,
            )
        except Exception as exc:
            db.rollback()
            return _redirect_dashboard(
                error=f"Usuario {email_n} creado pero la licencia falló: {exc}"
            )
        return _redirect_dashboard(
            success=f"Usuario {email_n} (rol {role_n}) creado y licencia '{code_n}' asignada."
        )
    return _redirect_dashboard(
        success=f"Usuario {email_n} (rol {role_n}) creado correctamente."
    )


@router.post("/users/{user_id}/delete")
def admin_users_delete(
    request: Request,
    user_id: int,
    db: Session = Depends(get_db),
):
    current_admin = _get_authenticated_admin(request, db)
    if not current_admin:
        return _redirect_to_login()

    if user_id == current_admin.id:
        return _redirect_dashboard(error="No podés borrarte a vos mismo.")

    user = crud.get_user_by_id(db, user_id)
    if not user:
        return _redirect_dashboard(error="Usuario no encontrado.")

    # No dejar borrar el ultimo admin activo
    if user.role == "admin" and user.is_active and crud.count_active_admins(db) <= 1:
        return _redirect_dashboard(
            error="No podés borrar al último administrador activo del sistema."
        )

    email = user.email
    try:
        db.delete(user)
        db.commit()
        return _redirect_dashboard(
            success=f"Usuario {email} y sus licencias/dispositivos borrados."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(error=f"No se pudo borrar el usuario: {exc}")


@router.post("/users/{user_id}/set-role")
def admin_users_set_role(
    request: Request,
    user_id: int,
    role: str = Form(...),
    db: Session = Depends(get_db),
):
    current_admin = _get_authenticated_admin(request, db)
    if not current_admin:
        return _redirect_to_login()

    if user_id == current_admin.id:
        return _redirect_dashboard(error="No podés cambiar tu propio rol.")

    user = crud.get_user_by_id(db, user_id)
    if not user:
        return _redirect_dashboard(error="Usuario no encontrado.")

    role_n = (role or "").strip().lower()
    if role_n not in crud.VALID_ROLES:
        return _redirect_dashboard(error=f"Rol inválido: {role_n!r}.")

    # No permitir degradar al ultimo admin activo
    if (
        user.role == "admin"
        and role_n != "admin"
        and crud.count_active_admins(db) <= 1
    ):
        return _redirect_dashboard(
            error="No podés degradar al último administrador activo del sistema."
        )

    try:
        crud.set_user_role(db, user, role_n)
        return _redirect_dashboard(
            success=f"Rol de {user.email} cambiado a '{role_n}'."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(error=f"No se pudo cambiar el rol: {exc}")


@router.post("/users/reset-password")
def admin_users_reset_password(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    new_password: str = Form(...),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    email_n = (email or "").strip().lower()
    password_n = new_password or ""
    if not email_n or not password_n:
        return _redirect_dashboard(
            error="Email y nueva contraseña son obligatorios."
        )

    user = crud.get_user_by_email(db, email=email_n)
    if not user:
        return _redirect_dashboard(error=f"No existe el usuario {email_n}.")

    try:
        user.password_hash = get_password_hash(password_n)
        db.add(user)
        # Invalidar tokens de reset pendientes
        db.query(models.PasswordResetToken).filter(
            models.PasswordResetToken.user_id == user.id
        ).delete(synchronize_session=False)
        db.commit()
        return _redirect_dashboard(
            success=f"Contraseña de {email_n} actualizada."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(
            error=f"No se pudo actualizar la contraseña: {exc}"
        )


@router.post("/licenses/create")
def admin_licenses_create(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    code: str = Form(...),
    license_duration: str = Form("none"),
    license_manual_date: str = Form(""),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    email_n = (email or "").strip().lower()
    code_n = (code or "").strip()
    if not email_n or not code_n:
        return _redirect_dashboard(error="Email y código son obligatorios.")

    user = crud.get_user_by_email(db, email=email_n)
    if not user:
        return _redirect_dashboard(error=f"No existe el usuario {email_n}.")

    if db.query(models.License).filter(models.License.code == code_n).first():
        return _redirect_dashboard(error=f"El código '{code_n}' ya existe.")

    try:
        expires_at, validity_days = _parse_expiration_from_form(
            license_duration, license_manual_date,
        )
    except ValueError as exc:
        return _redirect_dashboard(error=str(exc))

    try:
        crud.create_license(
            db, user=user, code=code_n,
            expires_at=expires_at, validity_days=validity_days,
        )
        return _redirect_dashboard(
            success=f"Licencia '{code_n}' agregada a {email_n}."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(
            error=f"No se pudo crear la licencia: {exc}"
        )


@router.post("/licenses/{license_id}/toggle")
def admin_licenses_toggle(
    request: Request,
    license_id: int,
    db: Session = Depends(get_db),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    lic = (
        db.query(models.License).filter(models.License.id == license_id).first()
    )
    if not lic:
        return _redirect_dashboard(error="Licencia no encontrada.")

    try:
        lic.is_active = not bool(lic.is_active)
        db.add(lic)
        db.commit()
        nuevo = "activada" if lic.is_active else "desactivada"
        return _redirect_dashboard(
            success=f"Licencia '{lic.code}' {nuevo}."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(
            error=f"No se pudo cambiar el estado de la licencia: {exc}"
        )


@router.post("/licenses/{license_id}/set-expiry")
def admin_licenses_set_expiry(
    request: Request,
    license_id: int,
    expires_at: str = Form(""),
    db: Session = Depends(get_db),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    lic = (
        db.query(models.License).filter(models.License.id == license_id).first()
    )
    if not lic:
        return _redirect_dashboard(error="Licencia no encontrada.")

    expires_str = (expires_at or "").strip()
    new_exp: Optional[datetime] = None
    if expires_str:
        try:
            new_exp = datetime.strptime(expires_str, "%Y-%m-%d")
            new_exp = new_exp.replace(hour=23, minute=59, second=59)
        except ValueError:
            return _redirect_dashboard(
                error=f"Fecha inválida: {expires_str!r} (formato YYYY-MM-DD)."
            )

    try:
        crud.set_license_expires_at(db, lic, new_exp)
        msg = (
            f"Expiración de '{lic.code}' actualizada a {new_exp.strftime('%Y-%m-%d')}."
            if new_exp else
            f"Licencia '{lic.code}' quedó sin vencimiento."
        )
        return _redirect_dashboard(success=msg)
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(
            error=f"No se pudo actualizar la expiración: {exc}"
        )


@router.post("/licenses/{license_id}/renew")
def admin_licenses_renew(
    request: Request,
    license_id: int,
    additional_days: int = Form(...),
    db: Session = Depends(get_db),
):
    if not _get_authenticated_admin(request, db):
        return _redirect_to_login()

    lic = (
        db.query(models.License).filter(models.License.id == license_id).first()
    )
    if not lic:
        return _redirect_dashboard(error="Licencia no encontrada.")

    if additional_days <= 0 or additional_days > 3650:
        return _redirect_dashboard(
            error="Los días adicionales deben estar entre 1 y 3650."
        )

    try:
        # from_now=True: renueva desde HOY, no desde la fecha vieja.
        # Esto evita que renovar una licencia vencida hace tiempo la deje
        # todavia vencida (o cerca).
        crud.renew_license(db, lic, additional_days, from_now=True)
        return _redirect_dashboard(
            success=(
                f"Licencia '{lic.code}' renovada por {additional_days} día(s). "
                f"Nueva expiración: {lic.expires_at.strftime('%Y-%m-%d')}."
            )
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(
            error=f"No se pudo renovar la licencia: {exc}"
        )
