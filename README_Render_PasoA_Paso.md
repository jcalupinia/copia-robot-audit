# Robot SRI-Audit — Descarga y Reporte (Versión Web)

## Descripción
Aplicación web sencilla que permite descargar comprobantes del SRI, generar un Excel y un ZIP con los archivos.

## Pasos rápidos para Render
1. Sube esta carpeta a un repositorio en GitHub.
2. En https://render.com → New + → Web Service → Conecta tu repo.
3. Render detectará `render.yaml` automáticamente.
4. Espera unos minutos: tendrás tu URL pública (ej. https://robot-sri-audit-web.onrender.com).

## Uso
- Ingresa RUC, clave SRI, año y mes.
- Elige tipo de documento y formato.
- Pulsa “Iniciar descarga” y luego baja tu Excel o ZIP.

## Licenciamiento y control de acceso

El robot ahora requiere autenticación previa para cada usuario:

1. **API de licencias**  
   - Ejecuta `uvicorn licensing_api.main:app --reload` (o despliega en Render como *Web Service*).  
   - Configura `LICENSE_DB_URL` y `LICENSE_API_SECRET` si deseas usar otra base de datos o clave JWT.
2. **Creación de cuentas**  
   - Usa `python scripts/manage_licenses.py create --email usuario@correo.com --password Secreta123 --code CODIGO-LICENCIA` para dar de alta a un cliente.
3. **Variables en la app Streamlit**  
   - Define `LICENSE_API_URL` apuntando al endpoint público del servicio FastAPI.
4. **Flujo del usuario**  
   - Inicia sesión con su correo/contraseña.  
   - Ingresa el código de licencia y el identificador del equipo para activarla.  
   - El sistema valida la licencia periódicamente durante la sesión.

> Si el SRI presenta indisponibilidad temporal, la app mostrará un mensaje para reintentar más tarde.
