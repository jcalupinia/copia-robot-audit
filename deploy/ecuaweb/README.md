# Despliegue en audit-ia.ec (Ecuaweb) + Render

Guia paso a paso para que la landing y el panel admin vivan en
`https://audit-ia.ec/sri_robot_audit/*` pero la API real corra en Render.

## Arquitectura

```
                              (proxy interno)
audit-ia.ec/sri_robot_audit/*  ───►  sri-robot-audit-ik01.onrender.com/*
   (Apache + .htaccess)              (FastAPI + admin + landing + DB)

Estado de los archivos:
   /sri_robot_audit/.htaccess        ← este reverse proxy
   /sri_robot_audit/proxy.php        ← fallback si mod_proxy no anda

Estado del .exe:
   GitHub Releases (CDN gratis)      ← donde subis cada version del .exe
```

## Configuracion paso a paso

### Paso 1 - Subir el .exe a GitHub Releases

1. En GitHub, ir a [Releases](https://github.com/jcalupinia/copia-robot-audit/releases) del repo.
2. **Create a new release**:
   - Tag: e.g. `v1.0.0`
   - Title: `ROBOT AUDIT SRI v1.0.0`
   - **Adjuntar el `.exe`** arrastrandolo a "Attach binaries".
3. Publicar.
4. Copiar la URL del `.exe`. Tiene la forma:

   ```
   https://github.com/jcalupinia/copia-robot-audit/releases/download/v1.0.0/ROBOT_AUDIT_SRI.exe
   ```

> Para futuras versiones repetis el proceso con un tag nuevo (`v1.0.1`, etc).
> GitHub mantiene todas las versiones disponibles para rollback.

### Paso 2 - Configurar variables de entorno en Render

En el dashboard de Render → tu servicio → **Settings → Environment**:

| Variable | Valor | Por que |
|---|---|---|
| `UPDATE_FILE_URL` | URL del `.exe` en GitHub Releases | El endpoint `/updates/download` ahora redirige a esta URL (302). El cliente baja directo del CDN de GitHub. |
| `UPDATE_TOKEN` | **(BORRAR)** | Hace publica la descarga. El bot desktop sigue funcionando (manda token pero el servidor lo ignora cuando esta vacio). |
| `UPDATE_VERSION` | e.g. `1.0.0` | Lo que el `/updates/latest` reporta como version disponible. |
| `ADMIN_EMAIL` | e.g. `admin@audit-ia.ec` | Para el login del panel admin. |
| `ADMIN_PASSWORD` | password fuerte 16+ chars | idem |
| `ALLOWED_ORIGINS` | `https://audit-ia.ec,https://sri-robot-audit-ik01.onrender.com` | CORS para que el frontend de audit-ia.ec pueda hablar con la API si en algun momento se necesita. |
| `APP_BASE_URL` | `https://audit-ia.ec/sri_robot_audit` | Base URL publica que la API usa para construir links (e.g. en mails de password reset). |
| (Opcional) podes BORRAR `R2_*` | | Si ya no usas Cloudflare R2. La logica R2 queda como fallback en el codigo pero no se ejecuta si tenes `UPDATE_FILE_URL` set. |

**Redeploy** despues de cambiar las vars.

### Paso 3 - Subir .htaccess a Ecuaweb

Por FTP/cPanel:

1. Crear la carpeta `audit-ia.ec/sri_robot_audit/` si no existe.
2. Subir el archivo `.htaccess` que esta en este repo bajo
   `deploy/ecuaweb/.htaccess` a esa carpeta.
3. (Si ya tenias archivos HTML estaticos de la landing alli, BORRARLOS — ahora
   la landing va a venir de Render via proxy.)

### Paso 4 - Verificar que el proxy funciona

Abrir en el navegador:

| URL | Debe mostrar |
|---|---|
| `https://audit-ia.ec/sri_robot_audit/landing` | La landing nueva con boton de descarga |
| `https://audit-ia.ec/sri_robot_audit/admin/login` | El form de login del admin |
| `https://audit-ia.ec/sri_robot_audit/updates/latest` | JSON con la version actual |

**Si las URLs dan 500 / 403 / pagina en blanco**: Ecuaweb no soporta `mod_proxy`.
Saltar al paso 5 (fallback con PHP).

**Si el download falla por CORS o redirect no funciona**: revisar que la cookie del
admin no tenga `Domain=` (el `.htaccess` ya lo limpia).

### Paso 5 (Opcional) - Fallback con PHP si mod_proxy no anda

Si el `.htaccess` no funciona porque Ecuaweb tiene `mod_proxy` capado:

1. Borrar el `.htaccess` que subiste.
2. Subir `deploy/ecuaweb/proxy.php` a `audit-ia.ec/sri_robot_audit/proxy.php`.
3. Crear un `.htaccess` NUEVO con este contenido:

   ```apache
   RewriteEngine On
   RewriteCond %{REQUEST_FILENAME} !-f
   RewriteCond %{REQUEST_FILENAME} !-d
   RewriteRule ^(.*)$ proxy.php?p=$1 [QSA,L]
   ```

   Esto manda cada peticion al `proxy.php` que la reenvia a Render con cURL.

### Paso 6 - Probar la descarga del .exe

1. Abrir `https://audit-ia.ec/sri_robot_audit/landing`
2. Click en "Descargar ROBOT_AUDIT_SRI.exe"
3. Comportamiento esperado:
   - Render responde 302 redirect a `UPDATE_FILE_URL`
   - El navegador sigue el redirect a GitHub Releases
   - GitHub sirve el `.exe` desde su CDN

### Paso 7 - Probar el panel admin

1. Abrir `https://audit-ia.ec/sri_robot_audit/admin/login`
2. Ingresar con `ADMIN_EMAIL` y `ADMIN_PASSWORD` que pusiste en Render.
3. Probar:
   - Crear usuario nuevo
   - Agregar licencia
   - Desactivar licencia
   - Borrar usuario
   - Resetear contrasena

## Por que Github Releases en vez de Cloudflare R2

| Tema | GitHub Releases | Cloudflare R2 |
|---|---|---|
| Costo | Gratis para repos publicos/privados | Gratis hasta 10GB + sin egress |
| Setup | Drag & drop en la UI de releases | Crear bucket, generar keys, configurar env vars |
| Versionado | Cada release es un tag, rollback trivial | Tenes que mantenerlo tu (mismo nombre = sobrescribe) |
| CDN global | Si (Fastly) | Si (Cloudflare) |
| URLs | Permanentes | Presigned con expiracion (si las usas) o publicas |
| Dependencias | Solo git/GitHub | `boto3` + config R2 + token + secret |

Para distribucion de un `.exe` versionado, GitHub Releases es estrictamente mejor.

## Como subir una version nueva del .exe

Una vez que esta todo configurado:

1. Compilar el `.exe` localmente (tu proceso actual).
2. En GitHub → New release → tag `v1.0.1` → adjuntar `ROBOT_AUDIT_SRI.exe`.
3. Copiar la URL del nuevo `.exe`.
4. En Render → Environment → actualizar `UPDATE_FILE_URL` con la URL nueva.
5. (Opcional) actualizar `UPDATE_VERSION` a `1.0.1`.
6. Redeploy.

A partir de ese momento:
- La landing en audit-ia.ec/sri_robot_audit/landing baja la version nueva.
- El bot desktop ya instalado en clientes ve que hay version nueva via
  `/updates/latest` y se auto-actualiza.

## Que pasa si rompo Ecuaweb (rollback de emergencia)

El admin y la landing tambien funcionan directo desde Render — sin
proxy de Ecuaweb. Usa estas URLs:

- Admin: `https://sri-robot-audit-ik01.onrender.com/admin/login`
- Landing: `https://sri-robot-audit-ik01.onrender.com/`

Y para distribuir el `.exe`:
- `https://sri-robot-audit-ik01.onrender.com/updates/download` (redirige a
  GitHub Releases)
