"""Panel administrativo web para la API de licencias.

Vive en /admin dentro de la misma FastAPI app y permite administrar
usuarios y licencias desde una interfaz HTML simple (sin Shell de Render).

Auth:
  - Login con ADMIN_EMAIL + ADMIN_PASSWORD (env vars).
  - Sesion en cookie httponly + samesite=lax firmada con el mismo
    JWT_SECRET_KEY que el resto de la API (reusa security.create_access_token).
  - El "subject" del JWT es `admin:<email>` para no colisionar con tokens
    de usuarios normales del cliente.
  - TTL: 4 horas.

Funcionalidad:
  - Dashboard con tabla de usuarios + sus licencias (estado, fecha activacion).
  - Crear usuario nuevo (opcional con licencia asociada en el mismo paso).
  - Agregar licencia a usuario existente.
  - Borrar usuario (cascada a licencias y dispositivos).
  - Desactivar / reactivar licencia (sin borrar).
  - Resetear contrasena de un usuario.

Diseno:
  - HTML/CSS inline en strings de Python (no requiere jinja2 ni assets).
  - Paleta similar a la landing: dark `#070a12` con acentos verde `#10b981`
    y azul `#60a5fa`.
  - Responsive simple (form-row con grid auto-fit).

Seguridad:
  - Todas las rutas /admin/* salvo /admin/login chequean cookie -> si no
    hay sesion valida, redirigen a /admin/login (303).
  - El login compara contra ADMIN_EMAIL/ADMIN_PASSWORD del env. Las
    contrasenas de USUARIOS DEL CLIENTE siempre se hashean con
    get_password_hash() — nunca se guardan en claro.
  - Cookie con httponly, samesite=lax y secure cuando la request entra
    por HTTPS (detectado por x-forwarded-proto si esta detras de proxy).
  - Pagina marcada con `robots: noindex, nofollow`.
"""

from __future__ import annotations

import html
import os
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
)


router = APIRouter(prefix="/admin", tags=["admin"])

ADMIN_COOKIE_NAME = "admin_session"
ADMIN_TOKEN_TTL = timedelta(hours=4)


# ============================================================
# AUTH HELPERS
# ============================================================

def _admin_email_env() -> str:
    return (os.getenv("ADMIN_EMAIL", "") or "").strip().lower()


def _admin_password_env() -> str:
    return os.getenv("ADMIN_PASSWORD", "") or ""


def _admin_creds_configured() -> bool:
    return bool(_admin_email_env()) and bool(_admin_password_env())


def _is_admin_authenticated(request: Request) -> bool:
    """True si la cookie de sesion es valida y el sub coincide con el
    ADMIN_EMAIL actual del env. Si las env vars cambian, las sesiones
    viejas se invalidan automaticamente.
    """
    token = request.cookies.get(ADMIN_COOKIE_NAME, "") or ""
    if not token or not _admin_creds_configured():
        return False
    sub = decode_access_token(token)
    if not sub:
        return False
    expected = f"admin:{_admin_email_env()}"
    return sub == expected


def _redirect_to_login() -> RedirectResponse:
    return RedirectResponse(url="/admin/login", status_code=status.HTTP_303_SEE_OTHER)


def _is_secure_request(request: Request) -> bool:
    proto = (request.headers.get("x-forwarded-proto", "") or "").strip().lower()
    if proto == "https":
        return True
    return request.url.scheme == "https"


def _set_admin_cookie(response: Response, request: Request) -> None:
    token = create_access_token(
        f"admin:{_admin_email_env()}",
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
.container { max-width: 1200px; margin: 0 auto; padding: 24px; }
.header {
    display: flex; align-items: center; justify-content: space-between;
    border-bottom: 1px solid var(--border); padding: 16px 24px;
    background: rgba(15,26,46,0.5); backdrop-filter: blur(6px);
}
.brand { display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 18px; color: var(--text); }
.brand .accent { color: var(--primary); }
.nav-actions { display: flex; gap: 4px; }
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
input[type=text], input[type=email], input[type=password] {
    width: 100%; padding: 10px 12px; background: rgba(7,10,18,0.7);
    border: 1px solid var(--border); border-radius: 8px;
    color: var(--text); font: inherit;
}
input[type=text]:focus, input[type=email]:focus, input[type=password]:focus {
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
table { width: 100%; border-collapse: collapse; font-size: 14px; }
th, td {
    padding: 12px 14px; text-align: left;
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
}
.pill.ok {
    background: rgba(16,185,129,0.12); border-color: rgba(16,185,129,0.4);
    color: var(--primary-strong);
}
.pill.bad {
    background: rgba(239,68,68,0.12); border-color: rgba(239,68,68,0.4);
    color: #fca5a5;
}
.pill.pending {
    background: rgba(245,158,11,0.12); border-color: rgba(245,158,11,0.4);
    color: #fcd34d;
}
.licencias-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 8px; }
.licencias-list li {
    background: rgba(7,10,18,0.5); border: 1px solid var(--border);
    border-radius: 8px; padding: 10px 12px;
}
.licencias-list .code {
    font-family: "SFMono-Regular", Consolas, Menlo, monospace;
    font-size: 12px; color: var(--accent); word-break: break-all;
}
.inline-form { display: inline-block; margin: 0; }
.actions { display: flex; gap: 6px; flex-wrap: wrap; }
.center-page { min-height: 80vh; display: grid; place-items: center; }
.login-card { max-width: 400px; width: 100%; padding: 32px; }
.empty {
    padding: 32px; text-align: center; color: var(--text-muted);
}
@media (max-width: 720px) {
    .header { padding: 12px 16px; }
    .container { padding: 16px; }
    table { font-size: 12px; }
    th, td { padding: 8px 6px; }
}
"""


def _layout(title: str, body: str, *, show_header: bool = True) -> str:
    title_safe = html.escape(title)
    header_html = ""
    if show_header:
        header_html = """
        <div class="header">
            <div class="brand">ROBOT&nbsp;AUDIT<span class="accent">&nbsp;SRI</span>&nbsp;· Admin</div>
            <div class="nav-actions">
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
    error: Optional[str] = None,
):
    if _is_admin_authenticated(request):
        return RedirectResponse(
            url="/admin", status_code=status.HTTP_303_SEE_OTHER,
        )

    error_html = ""
    if error:
        error_html = f'<div class="alert error">{html.escape(error)}</div>'

    creds_warn = ""
    if not _admin_creds_configured():
        creds_warn = (
            '<div class="alert info">'
            'Las variables <code>ADMIN_EMAIL</code> y <code>ADMIN_PASSWORD</code> '
            'no estan configuradas en el servidor. Sin ellas no se puede iniciar sesi&oacute;n.'
            '</div>'
        )

    body = f"""
    <div class="container center-page">
        <div class="card login-card">
            <h1 style="font-size:22px; margin-bottom:4px">Panel administrativo</h1>
            <p class="muted" style="margin-bottom:20px">
                Accedé con tus credenciales de administrador.
            </p>
            {error_html}
            {creds_warn}
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
    email: str = Form(...),
    password: str = Form(...),
):
    if not _admin_creds_configured():
        return RedirectResponse(
            url=(
                "/admin/login?error="
                + quote("No hay credenciales de admin configuradas en el servidor.")
            ),
            status_code=status.HTTP_303_SEE_OTHER,
        )

    email_n = (email or "").strip().lower()
    password_n = password or ""

    if email_n != _admin_email_env() or password_n != _admin_password_env():
        return RedirectResponse(
            url="/admin/login?error=" + quote("Credenciales incorrectas."),
            status_code=status.HTTP_303_SEE_OTHER,
        )

    response = RedirectResponse(
        url="/admin", status_code=status.HTTP_303_SEE_OTHER,
    )
    _set_admin_cookie(response, request)
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

@router.get("/", response_class=HTMLResponse)
def admin_dashboard(
    request: Request,
    db: Session = Depends(get_db),
    success: Optional[str] = None,
    error: Optional[str] = None,
):
    if not _is_admin_authenticated(request):
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
        # Licencias del usuario
        lic_items: list[str] = []
        for lic in (u.licenses or []):
            lic_estado = "activa" if lic.is_active else "inactiva"
            estado_pill_cls = "ok" if lic.is_active else "bad"
            act_text = (
                f"Activada: {lic.activated_at.strftime('%Y-%m-%d %H:%M')}"
                if lic.activated_at
                else "<em>Sin activar</em>"
            )
            exp_text = (
                f" · Expira: {lic.expires_at.strftime('%Y-%m-%d')}"
                if lic.expires_at else ""
            )
            toggle_label = "Desactivar" if lic.is_active else "Reactivar"
            toggle_cls = "btn-warn" if lic.is_active else "btn-primary"

            lic_items.append(f"""
            <li>
                <div class="code">{html.escape(lic.code or '')}</div>
                <div style="margin-top:6px; display:flex; gap:8px; align-items:center; flex-wrap:wrap">
                    <span class="pill {estado_pill_cls}">{lic_estado}</span>
                    <span class="muted">{act_text}{exp_text}</span>
                </div>
                <div class="actions" style="margin-top:8px">
                    <form method="post" action="/admin/licenses/{lic.id}/toggle" class="inline-form">
                        <button type="submit" class="btn {toggle_cls}" style="padding:4px 10px; font-size:12px">{toggle_label}</button>
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
        fecha = u.created_at.strftime("%Y-%m-%d %H:%M") if u.created_at else "-"
        email_safe = html.escape(u.email or "")
        email_attr = html.escape(u.email or "", quote=True)

        rows_html_parts.append(f"""
        <tr>
            <td><strong>{email_safe}</strong><br><span class="muted">ID {u.id}</span></td>
            <td><span class="pill {estado_user_pill}">{estado_user_txt}</span></td>
            <td>{fecha}</td>
            <td>{licencias_html}</td>
            <td>
                <div class="actions">
                    <form method="post" action="/admin/users/{u.id}/delete" class="inline-form"
                          onsubmit="return confirm('¿Borrar al usuario {email_attr} y todas sus licencias? Esta accion es IRREVERSIBLE.');">
                        <button type="submit" class="btn btn-danger" style="padding:4px 10px; font-size:12px">Borrar</button>
                    </form>
                </div>
            </td>
        </tr>
        """)

    table_body = (
        "".join(rows_html_parts)
        if rows_html_parts
        else '<tr><td colspan="5" class="empty">No hay usuarios todavía.</td></tr>'
    )

    body = f"""
    <div class="container">
        <h1>Dashboard</h1>
        <p class="muted">Gestión de usuarios y licencias.</p>

        {success_html}
        {error_html}

        <div class="card">
            <h2 style="margin-top:0">Crear usuario nuevo</h2>
            <p class="muted">Opcional: incluí un código de licencia para crearla en el mismo paso.</p>
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
                        <label for="new_license">Código de licencia (opcional)</label>
                        <input type="text" name="license_code" id="new_license" placeholder="AUDIT-XXXX-YYYY">
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
    return HTMLResponse(content=_layout("Dashboard", body))


# ============================================================
# ROUTES — acciones
# ============================================================

@router.post("/users/create")
def admin_users_create(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    password: str = Form(...),
    license_code: str = Form(""),
):
    if not _is_admin_authenticated(request):
        return _redirect_to_login()

    email_n = (email or "").strip().lower()
    password_n = password or ""
    code_n = (license_code or "").strip()

    if not email_n or not password_n:
        return _redirect_dashboard(error="Email y contraseña son obligatorios.")

    existing = crud.get_user_by_email(db, email=email_n)
    if existing:
        return _redirect_dashboard(error=f"El usuario {email_n} ya existe.")

    try:
        user = crud.create_user(db, email=email_n, password=password_n)
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(error=f"No se pudo crear el usuario: {exc}")

    if code_n:
        existing_lic = (
            db.query(models.License).filter(models.License.code == code_n).first()
        )
        if existing_lic:
            # No queremos un usuario sin licencia cuando el admin pidio una.
            try:
                db.delete(user)
                db.commit()
            except Exception:
                db.rollback()
            return _redirect_dashboard(
                error=(
                    f"El código de licencia '{code_n}' ya está en uso. "
                    f"Usuario {email_n} NO creado."
                )
            )
        try:
            crud.create_license(db, user=user, code=code_n)
        except Exception as exc:
            db.rollback()
            return _redirect_dashboard(
                error=(
                    f"Usuario {email_n} creado pero la licencia falló: {exc}"
                )
            )
        return _redirect_dashboard(
            success=f"Usuario {email_n} creado y licencia '{code_n}' asignada."
        )
    return _redirect_dashboard(success=f"Usuario {email_n} creado correctamente.")


@router.post("/users/{user_id}/delete")
def admin_users_delete(
    request: Request,
    user_id: int,
    db: Session = Depends(get_db),
):
    if not _is_admin_authenticated(request):
        return _redirect_to_login()

    user = db.query(models.User).filter(models.User.id == user_id).first()
    if not user:
        return _redirect_dashboard(error="Usuario no encontrado.")

    email = user.email
    try:
        # El cascade="all,delete" del modelo se encarga de licencias,
        # y a su vez License.devices tiene cascade="all,delete-orphan"
        # para borrar los dispositivos asociados. PasswordResetToken
        # tambien tiene cascade="all,delete-orphan".
        db.delete(user)
        db.commit()
        return _redirect_dashboard(
            success=f"Usuario {email} y sus licencias/dispositivos borrados."
        )
    except Exception as exc:
        db.rollback()
        return _redirect_dashboard(error=f"No se pudo borrar el usuario: {exc}")


@router.post("/users/reset-password")
def admin_users_reset_password(
    request: Request,
    db: Session = Depends(get_db),
    email: str = Form(...),
    new_password: str = Form(...),
):
    if not _is_admin_authenticated(request):
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
        # Invalidar tokens de reset pendientes — la contrasena ya cambio,
        # los links viejos no deben servir.
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
):
    if not _is_admin_authenticated(request):
        return _redirect_to_login()

    email_n = (email or "").strip().lower()
    code_n = (code or "").strip()
    if not email_n or not code_n:
        return _redirect_dashboard(error="Email y código son obligatorios.")

    user = crud.get_user_by_email(db, email=email_n)
    if not user:
        return _redirect_dashboard(error=f"No existe el usuario {email_n}.")

    existing = (
        db.query(models.License).filter(models.License.code == code_n).first()
    )
    if existing:
        return _redirect_dashboard(error=f"El código '{code_n}' ya existe.")

    try:
        crud.create_license(db, user=user, code=code_n)
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
    if not _is_admin_authenticated(request):
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
