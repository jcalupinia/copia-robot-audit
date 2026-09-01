# Manual de Usuario — SRI Robot Audit

**Sistema de descarga y reporte automático del SRI Ecuador**

| Dato | Valor |
|------|-------|
| Versión del sistema | 2026.05.05.88 |
| Última actualización del manual | Módulo rápido de reportes y Retenciones vs Facturas |
| Empresa | Audit Consulting Group |
| Modalidad | Web (Render) y Escritorio (.exe) |
| Fecha de edición | _A completar_ |
| Autor del manual | _A completar_ |

---

## ÍNDICE GENERAL

1. [Introducción](#1-introducción)
2. [Objetivo del sistema](#2-objetivo-del-sistema)
3. [Requisitos mínimos](#3-requisitos-mínimos)
4. [Descarga e instalación](#4-descarga-e-instalación)
5. [Acceso al sistema](#5-acceso-al-sistema)
6. [Descripción del menú principal](#6-descripción-del-menú-principal)
7. [Descripción detallada de cada módulo](#7-descripción-detallada-de-cada-módulo)
8. [Procedimientos paso a paso](#8-procedimientos-paso-a-paso)
9. [Tabla de mensajes de error y solución](#9-tabla-de-mensajes-de-error-y-solución)
10. [Preguntas frecuentes](#10-preguntas-frecuentes)
11. [Soporte](#11-soporte)
12. [Anexos](#12-anexos)

---

# 1. INTRODUCCIÓN

## 1.1 Propósito del manual

Este manual describe paso a paso cómo utilizar **SRI Robot Audit**, una aplicación que automatiza la descarga de comprobantes electrónicos (XML y RIDE/PDF) desde el portal del Servicio de Rentas Internas (SRI) del Ecuador y genera reportes consolidados en formato Excel.

Está dirigido al usuario final del sistema: contadores, auditores, asistentes contables y personal administrativo que necesita obtener masivamente comprobantes electrónicos para procesos de auditoría, conciliación tributaria o respaldo documental.

## 1.2 Alcance y audiencia

- **Audiencia primaria**: usuario final con conocimientos básicos de Windows y del portal del SRI.
- **Alcance**: cubre el 100 % de las pantallas y funcionalidades visibles en la aplicación (no incluye instalación del servidor de licencias ni administración interna).

## 1.3 Convenciones tipográficas

| Elemento | Convención |
|---|---|
| Nombres de botones | **Negrita** |
| Nombres de campos | _Cursiva_ |
| Mensajes del sistema | "entre comillas" |
| Rutas y archivos | `formato monoespaciado` |
| Capturas de pantalla | `[CAPTURA NN — Nombre]` con descripción y flechas |
| Recuadros de advertencia | ⚠️ Atención: … |
| Recuadros informativos | ℹ️ Nota: … |

## 1.4 Glosario rápido

| Término | Significado |
|---|---|
| **RUC** | Registro Único de Contribuyentes (13 dígitos). |
| **Clave del SRI** | Contraseña que el contribuyente usa en el portal SRI. |
| **Comprobante** | Documento electrónico tributario (factura, retención, etc.). |
| **XML** | Archivo electrónico oficial firmado del comprobante. |
| **RIDE / PDF** | Representación Impresa del Documento Electrónico (PDF). |
| **Recibidos** | Comprobantes emitidos a favor del contribuyente por terceros. |
| **Emitidos** | Comprobantes que el contribuyente emite hacia terceros. |
| **Licencia** | Código que habilita el uso del software en un equipo. |
| **Fingerprint** | Huella digital única generada para el equipo donde corre la app. |
| **Consolidación** | Proceso de unir varios reportes mensuales en uno solo. |
| **Tour** | Recorrido guiado de introducción a la app. |
| **Nota de Crédito (NC)** | Comprobante que modifica (anula parcial o totalmente) una Factura previa. |
| **Valor Neto** | Importe Total Factura − Importe Total Nota de Crédito. |
| **Clave de acceso** | Identificador único de 49 dígitos de cada comprobante en el SRI. |
| **Código numérico** | 8 dígitos dentro del clave de acceso, asignados por el facturador del emisor. |

---

# 2. OBJETIVO DEL SISTEMA

## 2.1 Qué hace SRI Robot Audit

SRI Robot Audit automatiza tres procesos que tradicionalmente se realizan de forma manual y consumen muchas horas:

1. **Descarga masiva** de comprobantes (XML y PDF) desde el portal oficial del SRI.
2. **Generación de reportes Excel** consolidados con la información extraída de los comprobantes descargados.
3. **Consolidación** de múltiples reportes mensuales en un solo reporte anual o por periodo.

El sistema utiliza un robot interno (basado en navegación automatizada con Playwright) que ingresa al portal del SRI con las credenciales del contribuyente, ejecuta las consultas y descarga los archivos respetando los filtros indicados por el usuario.

## 2.2 Beneficios para el auditor / contador

- Elimina la descarga manual archivo por archivo.
- Procesa lotes grandes de comprobantes sin intervención del usuario.
- Genera reportes Excel listos para conciliación tributaria.
- Reanuda automáticamente descargas interrumpidas por desconexión o cierre accidental.
- Mantiene un historial trazable de todas las ejecuciones por equipo.

## 2.3 Alcance funcional

El sistema cuenta con cuatro módulos principales (pestañas):

| Pestaña | Función |
|---|---|
| Descarga de Comprobantes | Descarga directa desde el portal del SRI. |
| Reportes e Historial | Genera reportes Excel desde archivos ya descargados y muestra el historial de ejecuciones. |
| Consolidación de documentos | Une múltiples reportes mensuales en un consolidado por periodo. |
| Ayuda | Centro de ayuda con guía rápida, FAQ y actualizaciones. |

## 2.4 Fuera de alcance

- El sistema **no** firma comprobantes ni emite documentos al SRI.
- **No** realiza declaraciones tributarias.
- **No** sustituye al portal del SRI: lo consulta y descarga, pero todas las validaciones legales son responsabilidad del contribuyente.

---

# 3. REQUISITOS MÍNIMOS

## 3.1 Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i3 o AMD equivalente | Intel Core i5 / Ryzen 5 |
| Memoria RAM | 4 GB | 8 GB o más |
| Almacenamiento libre | 1 GB | 5 GB (para descargas masivas) |
| Resolución de pantalla | 1280 × 720 | 1920 × 1080 |

## 3.2 Sistema operativo soportado

- **Windows 10** (64 bits) — soportado.
- **Windows 11** (64 bits) — recomendado.

La versión .exe está compilada para arquitectura Windows. La versión web puede usarse desde cualquier navegador moderno (Chrome, Edge, Firefox).

## 3.3 Conexión a internet

- Banda ancha estable (mínimo 5 Mbps recomendado).
- Acceso permitido a:
  - `https://srienlinea.sri.gob.ec` (portal del SRI)
  - `https://sri-robot-audit-ik01.onrender.com` (servidor de licencias y actualizaciones)

## 3.4 Permisos de carpeta

El usuario debe tener permisos de **escritura** sobre la carpeta donde se almacenarán las descargas (por defecto `C:\AuditSRI\Descargas` o la que el usuario elija).

## 3.5 Navegador Chromium embebido

La aplicación incluye un navegador interno (Playwright) que se utiliza para acceder al portal del SRI. No se requiere instalarlo aparte: viene empaquetado en el .exe.

## 3.6 Modalidades de uso

| Modalidad | Cómo se distribuye | Ventaja |
|---|---|---|
| **Web** | Se accede vía URL desde el navegador. | No requiere instalación. |
| **Escritorio (.exe)** | Se descarga e instala `ROBOT_AUDIT_SRI.exe`. | Funciona sin depender del servidor web; auto-actualizable. |

---

# 4. DESCARGA E INSTALACIÓN

## 4.1 Descarga del instalador

El software se distribuye como un ejecutable de Windows (`ROBOT_AUDIT_SRI.exe`). Para descargarlo:

1. Abra el portal web de descarga en su navegador:
   - URL: `https://sri-robot-audit-ik01.onrender.com`
2. En la sección **Descarga ahora**, presione el botón **Descargar ROBOT_AUDIT_SRI.exe**.
3. El navegador comenzará la descarga.

### 4.1.1 Advertencias del navegador y de Windows

Como el ejecutable proviene de un servidor propio (no del Microsoft Store), tanto el navegador como Windows muestran advertencias estándar al descargarlo y al abrirlo por primera vez. **Estas advertencias son esperadas y NO indican que el archivo sea malicioso** — se resuelven confirmando cada paso.

| Sistema | Mensaje típico | Acción para continuar |
|---|---|---|
| Chrome / Edge (barra inferior) | "Este archivo puede dañar tu computadora..." | Presione el menú **⋁** o **⋯** junto al archivo descargado → **Conservar** / **Mantener** / **Mantener de todos modos**. |
| Edge (segunda confirmación) | "¿Mantener este archivo?" | Seleccione **Mostrar más** → **Conservar de todos modos**. |
| Windows SmartScreen (al abrir) | "Windows protegió tu PC..." | Presione **Más información** → **Ejecutar de todas formas**. |
| Antivirus de terceros | "Archivo inusual detectado" | Marque el archivo como **Confiable** y permita la ejecución. |

ℹ️ **Por qué aparecen estas advertencias**: el ejecutable es nuevo en el ecosistema de Microsoft y aún no tiene una firma digital corporativa propagada masivamente. El archivo es seguro y proviene del servidor oficial de Audit Consulting (`sri-robot-audit-ik01.onrender.com`).

### 4.1.2 Opción "Elegir dónde guardar"

Junto al botón principal de descarga hay una opción **Elegir dónde guardar** que permite seleccionar la carpeta destino antes de descargar:

- Funciona en navegadores modernos (Chrome, Edge) sobre conexión HTTPS.
- Si el navegador no la soporta, use el botón principal y luego mueva el archivo desde la carpeta de **Descargas** a la ubicación que prefiera.

## 4.2 Primer arranque del .exe

Al hacer doble clic sobre `ROBOT_AUDIT_SRI.exe` por primera vez:

1. Se abrirá una **ventana negra de consola (CMD)** con los siguientes mensajes:
   ```
   Iniciando software, espere un momento...
   Preparando entorno y cargando la aplicacion.
   ```

⚠️ **Atención — Tiempo de carga del primer arranque**

La primera vez que se ejecuta, la ventana puede demorar entre **1 y 2 minutos** en abrir la interfaz gráfica. **Esto es completamente normal**: depende de las características del equipo (procesador, memoria RAM, presencia de antivirus, tipo de disco).

Los siguientes arranques son más rápidos (entre 15 y 30 segundos) porque el sistema ya tiene preparado el entorno.

**¿Por qué tarda?**
- El ejecutable empaqueta internamente Python, Playwright y todas las librerías necesarias. En el primer inicio descomprime esos archivos a una carpeta temporal del sistema.
- Si tiene antivirus activo, este analiza el ejecutable y los archivos descomprimidos antes de permitir la ejecución (segundos adicionales).
- En equipos con disco mecánico (HDD), el proceso es notablemente más lento que en SSD.
- En equipos con poca RAM (4 GB) puede llegar al límite alto de 2 minutos.

2. La aplicación se **autoinstala automáticamente** en:
   ```
   C:\Users\<usuario>\AppData\Local\ROBOT_AUDIT_SRI\
   ```
   No requiere permisos de administrador; se instala con los permisos del usuario actual.

3. Una vez completada la carga, se abrirá automáticamente la interfaz gráfica en una pestaña nueva de su navegador predeterminado, mostrando la pantalla de **Inicio de sesión** (ver sección 5.1).

⚠️ **No cierre la ventana CMD mientras use la aplicación**: esa ventana mantiene activo el servidor interno que comunica el navegador con el robot. Si la cierra, la aplicación dejará de responder y deberá reiniciarla.

### 4.2.1 Si la ventana CMD se cierra sola sin abrir el navegador

Indica un error durante el arranque. Para diagnosticarlo:

1. Vaya a la carpeta:
   ```
   C:\Users\<usuario>\AppData\Local\ROBOT_AUDIT_SRI\
   ```
2. Abra el archivo `desktop_launcher.log` con el Bloc de notas.
3. Incluya su contenido al contactar a soporte (sección 11).

### 4.2.2 Configuración inicial (`desktop_config.json`)

En el primer arranque, el sistema crea junto al ejecutable un archivo `desktop_config.json` con la configuración por defecto:

```json
{
  "LICENSE_API_URL": "https://sri-robot-audit-ik01.onrender.com",
  "SESSION_CACHE_DIR": ".session_cache"
}
```

ℹ️ **No es necesario modificar este archivo** para el uso normal. Si su organización utiliza un servidor de licencias propio, su administrador le indicará los valores correctos.

## 4.3 Actualizaciones automáticas

Cada vez que se inicia la aplicación, el launcher consulta al servidor si existe una versión más reciente:

- Si hay actualización disponible, la descarga, reemplaza el ejecutable actual y la deja lista en el siguiente arranque.
- El proceso es transparente: no requiere acción del usuario.
- También puede forzarse manualmente desde la pestaña **Ayuda → Acerca de la aplicación → Buscar actualizaciones** (ver sección 8.13).

## 4.4 Desinstalación

Para desinstalar el software:

1. Cierre la aplicación si está abierta (incluida la ventana CMD).
2. Elimine la carpeta:
   ```
   C:\Users\<usuario>\AppData\Local\ROBOT_AUDIT_SRI\
   ```
3. Opcionalmente, elimine los archivos generados por la aplicación (descargas, reportes, historial) en la carpeta base que haya configurado.

---

# 5. ACCESO AL SISTEMA

## 5.1 Inicio de sesión

Al abrir la aplicación, lo primero que aparece es la pantalla de inicio de sesión.

**Propósito de la pantalla**: autenticar al usuario contra el servidor de licencias.

### Campos

| Campo | Tipo | Obligatorio | Descripción |
|---|---|---|---|
| _Correo electrónico_ | Texto | Sí | Correo registrado por el administrador. |
| _Contraseña_ | Password | Sí | Contraseña personal del usuario. |

### Botones y acciones

| Botón | Acción |
|---|---|
| **Iniciar sesión** | Envía las credenciales y abre la sesión. |
| **¿Olvidaste tu contraseña?** | Navega al flujo de recuperación. |

### Flujo

1. El usuario abre la aplicación.
2. Ingresa el correo electrónico y la contraseña.
3. Presiona **Iniciar sesión**.
4. Si las credenciales son correctas:
   - Si la licencia ya está activada en este equipo → entra al menú principal.
   - Si la licencia aún no está activada → continúa a la pantalla de activación.
5. Si las credenciales son incorrectas, aparece el mensaje "Credenciales incorrectas".

```
[CAPTURA 01 — Pantalla de inicio de sesión]
Descripción de la captura:
- Flecha 1: Encabezado "Iniciar sesión" + subtítulo.
- Flecha 2: Campo correo electrónico.
- Flecha 3: Campo contraseña.
- Flecha 4: Botón Iniciar sesión.
- Flecha 5: Enlace "¿Olvidaste tu contraseña?".
- Flecha 6: Insignia "Sesión vinculada al dispositivo".
```

## 5.2 Recuperación de contraseña

### 5.2.1 Solicitar enlace por correo

**Propósito de la pantalla**: enviar al correo del usuario un enlace seguro para crear una nueva contraseña.

**Cómo se accede**: desde la pantalla de login, presione el enlace "¿Olvidaste tu contraseña?".

#### Campos

| Campo | Tipo | Obligatorio | Descripción |
|---|---|---|---|
| _Correo electrónico_ | Texto | Sí | Correo asociado a la cuenta. |

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Enviar enlace** | Envía un correo con un enlace de recuperación válido. |
| **← Volver a iniciar sesión** | Regresa a la pantalla de login. |

#### Flujo

1. Ingrese el correo electrónico registrado.
2. Presione **Enviar enlace**.
3. El sistema responde: "Si el correo existe, enviaremos un enlace de recuperación." (mensaje uniforme por seguridad).
4. Revise su bandeja de entrada (y carpeta de spam).
5. Abra el enlace que le llegó por correo para continuar con el restablecimiento.

```
[CAPTURA 02 — Solicitar recuperación de contraseña]
Descripción de la captura:
- Flecha 1: Título "Recuperar contraseña".
- Flecha 2: Campo correo electrónico.
- Flecha 3: Botón Enviar enlace.
- Flecha 4: Botón Volver a iniciar sesión.
```

### 5.2.2 Restablecer contraseña desde el enlace

**Propósito de la pantalla**: definir una nueva contraseña usando el token recibido por correo.

**Cómo se accede**: al hacer clic en el enlace recibido por correo (URL del tipo `?reset_token=...`).

#### Campos

| Campo | Tipo | Obligatorio | Descripción |
|---|---|---|---|
| _Correo electrónico_ | Texto (deshabilitado) | — | Se muestra automáticamente desde el token. |
| _Nueva contraseña_ | Password | Sí | Mínimo 8 caracteres. |
| _Confirmar contraseña_ | Password | Sí | Debe coincidir con la nueva contraseña. |

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Guardar contraseña** | Confirma la nueva contraseña. |
| **← Volver a iniciar sesión** | Cancela y vuelve al login. |

#### Flujo

1. Abra el enlace de recuperación recibido por correo.
2. El correo se mostrará automáticamente (no se puede modificar).
3. Ingrese la nueva contraseña.
4. Repítala en el campo de confirmación.
5. Presione **Guardar contraseña**.
6. Si todo es correcto, aparece "Tu contraseña se actualizó correctamente" y vuelve al login.

```
[CAPTURA 03 — Restablecer contraseña]
Descripción de la captura:
- Flecha 1: Título "Restablecer contraseña".
- Flecha 2: Campo correo (deshabilitado).
- Flecha 3: Campo nueva contraseña.
- Flecha 4: Campo confirmar contraseña.
- Flecha 5: Botón Guardar contraseña.
- Flecha 6: Botón Volver a iniciar sesión.
```

## 5.3 Activación de licencia

**Propósito de la pantalla**: vincular el equipo del usuario con un código de licencia válido. Cada licencia puede activarse en un solo equipo a la vez.

**Cómo se accede**: aparece automáticamente tras el primer inicio de sesión si la licencia aún no fue activada en este equipo.

### Campos

| Campo | Tipo | Obligatorio | Descripción |
|---|---|---|---|
| _Código de licencia_ | Texto | Sí | Código entregado por el administrador del sistema. |
| _Identificador del equipo_ | Texto (deshabilitado) | — | Se genera automáticamente (fingerprint). |

### Botones y acciones

| Botón | Acción |
|---|---|
| **Activar licencia** | Valida el código y vincula el equipo. |
| **Volver a inicio de sesión** | Sale de la activación. |

### Flujo

1. Ingrese el código de licencia (lo entrega el administrador).
2. El identificador del equipo se muestra ya rellenado (no se modifica).
3. Presione **Activar licencia**.
4. Si todo es correcto, aparece "Licencia activada correctamente" y entra al menú principal.

⚠️ **Atención**: una licencia ya activada en otro equipo no podrá usarse en este. En ese caso aparece "Esta licencia ya está activada en otro equipo".

```
[CAPTURA 04 — Activación de licencia]
Descripción de la captura:
- Flecha 1: Logo de Audit Consulting Group.
- Flecha 2: Título "Activación de licencia".
- Flecha 3: Mensaje de advertencia.
- Flecha 4: Campo código de licencia.
- Flecha 5: Identificador del equipo (deshabilitado).
- Flecha 6: Botón Activar licencia.
- Flecha 7: Botón Volver a inicio de sesión (esquina superior izquierda).
```

## 5.4 Cerrar sesión / cerrar app (popover Perfil)

**Propósito**: cerrar la sesión activa o cerrar la aplicación (solo en versión .exe).

**Cómo se accede**: clic en el botón **👤 Perfil** ubicado en el extremo derecho del topbar.

### Acciones disponibles dentro del popover

| Elemento | Función |
|---|---|
| Sesión activa + correo | Información: muestra el correo del usuario activo. |
| **🚪 Cerrar sesión** | Cierra la sesión y vuelve al login. |
| **⏻ Cerrar app** | (Solo en .exe) Cierra la aplicación tras confirmación. |

```
[CAPTURA 05 — Popover de perfil]
Descripción de la captura:
- Flecha 1: Botón 👤 Perfil en el topbar.
- Flecha 2: Indicador "Sesión activa".
- Flecha 3: Correo del usuario.
- Flecha 4: Botón Cerrar sesión.
- Flecha 5: Botón Cerrar app (solo visible en versión .exe).
```

---

# 6. DESCRIPCIÓN DEL MENÚ PRINCIPAL

## 6.1 Topbar (encabezado fijo)

El topbar se mantiene siempre visible en la parte superior de la pantalla. Está dividido en tres zonas:

### Zona izquierda
- **Logo Audit Consulting Group** (versión clara u oscura según el tema activo).

### Zona central
- Título: **"SRI Robot Audit | Descarga y Reporte Automático"**.

### Zona derecha
- **Botón de tema** (☀️ Claro / 🌙 Oscuro): alterna entre tema claro y oscuro.
- **Botón 👤 Perfil**: abre el popover con datos de sesión y acciones.

## 6.2 Pestañas (tabs) principales

Debajo del topbar aparece la barra de pestañas. Las pestañas disponibles son:

| Pestaña | Propósito |
|---|---|
| **Descarga de Comprobantes** | Descargar archivos XML/PDF desde el portal del SRI. |
| **Reportes e Historial** | Generar reportes Excel desde archivos ya descargados y revisar ejecuciones previas. |
| **Consolidación de documentos** | Unir reportes mensuales en un consolidado anual o por periodo. |
| **Ayuda** | Centro de ayuda, FAQ, tour de primer uso y actualizaciones. |

## 6.3 Botón "Primera vez? Ver tour"

En la parte superior derecha de la pestaña **Descarga de Comprobantes** aparece un botón en forma de píldora con un punto verde:

- **Botón**: **● Primera vez? Ver tour**
- **Acción**: inicia un recorrido guiado paso a paso por las funciones principales del sistema.

```
[CAPTURA 06 — Vista general del topbar y tabs]
Descripción de la captura:
- Flecha 1: Logo de Audit Consulting Group.
- Flecha 2: Título central.
- Flecha 3: Botón de tema claro/oscuro.
- Flecha 4: Botón 👤 Perfil.
- Flecha 5: Pestañas principales.
- Flecha 6: Botón "Primera vez? Ver tour".
```

---

# 7. DESCRIPCIÓN DETALLADA DE CADA MÓDULO

## 7.1 MÓDULO 1 — DESCARGA DE COMPROBANTES (Tab 1)

### 7.1.1 Propósito del módulo

Descargar masivamente comprobantes electrónicos (XML y/o PDF) desde el portal del SRI según los filtros indicados por el usuario.

### 7.1.2 Estructura por pasos (1 a 4)

El módulo está dividido en cuatro tarjetas numeradas:

1. **Credenciales** — datos para autenticarse en el SRI.
2. **Filtros** — qué descargar.
3. **Carpeta base** — dónde guardar.
4. **Ejecutar** — iniciar o detener el proceso.

### 7.1.3 Paso 1 — Credenciales

**Propósito**: ingresar las credenciales del contribuyente para que el robot se autentique automáticamente en el portal del SRI.

#### Campos

| Campo | Tipo | Obligatorio | Descripción |
|---|---|---|---|
| _RUC_ | Texto | Sí | Número de RUC del contribuyente (13 dígitos). Ejemplo: `0999999001`. |
| _Clave del SRI_ | Password | Sí | Contraseña del portal del SRI. No se guarda fuera de esta sesión. |

ℹ️ **Nota**: la clave del SRI **no se almacena en disco**. Se utiliza únicamente durante la sesión actual.

```
[CAPTURA 07 — Card Credenciales]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 1.
- Flecha 2: Título "Credenciales".
- Flecha 3: Subtítulo "Datos del SRI".
- Flecha 4: Campo RUC.
- Flecha 5: Campo Clave del SRI.
- Flecha 6: Icono de ojo para mostrar/ocultar la clave.
```

### 7.1.4 Paso 2 — Filtros

**Propósito**: definir qué tipo de comprobantes descargar, en qué rango de fechas y en qué formato.

Los filtros varían ligeramente según el origen elegido (Recibidos o Emitidos).

#### Campos comunes

| Campo | Tipo | Descripción |
|---|---|---|
| _Origen de comprobantes_ | Lista desplegable | "Recibidos" o "Emitidos". |
| _Tipo de comprobante_ | Lista desplegable | Tipos disponibles según origen (ver más abajo). |
| _Modo de fecha_ | Radio | "Mes y día", "Rango de meses", "Año completo". |
| _Año_ | Numérico | Año a consultar. |
| _Mes_ | Lista desplegable | Mes a consultar (cuando el modo lo requiere). |
| _Día (0 = Todos)_ | Numérico | Día específico o `0` para todo el mes. |
| _Modo rápido: solo reporte_ | Checkbox | Genera el Excel del período sin descargar ningún comprobante. |
| _Formatos a descargar_ | Multiselect | "XML", "PDF" o ambos. |

#### Modo rápido — solo reporte (sin PDF ni XML)

**Propósito**: obtener el listado completo del período en un Excel, en segundos, sin descargar un archivo por cada comprobante.

En vez de abrir cada fila para bajar su XML y su PDF, el sistema toma los datos que el propio portal ya publica en pantalla. Una consulta de 300 comprobantes pasa de varios minutos a unos pocos segundos.

| Origen | De dónde toma los datos | Consultas al portal |
|---|---|---|
| **Recibidos** | Del archivo que ofrece el enlace "Descargar reporte" del portal | 1 por mes |
| **Emitidos** | De la tabla en pantalla | 1 por día (el portal filtra por día) |

El checkbox está ubicado **entre el modo de fecha y los formatos**, porque es la decisión de *qué querés obtener* antes que la de *en qué formato*:

- Al marcarlo, la selección de formatos queda deshabilitada y vacía.
- Al elegir XML o PDF, el modo rápido se desmarca solo.

⚠️ **Qué NO hace**: no descarga comprobantes. Si necesitás los XML o los PDF —para respaldo, para el módulo de Consolidación o para el reporte de Retenciones vs Facturas— usá los formatos normales.

**Dónde queda el Excel**:

```
[Carpeta base]/[RUC]/[Origen]/[Tipo]/[Año]/[Mes]/TXT/
    recibidos_reporte_txt_[tipo]_[AAAAMM].xlsx
```

Si elegiste un rango de meses o el año completo, además se genera un **Excel consolidado** con todo el período en `[Año]/TXT/`.

Las columnas son las mismas que publica el portal: comprobante, serie, RUC y razón social del emisor, clave de acceso, fechas de emisión y autorización, valor sin impuestos, IVA e importe total.

```
[CAPTURA 08b — Checkbox Modo rápido activo]
Descripción de la captura:
- Flecha 1: Checkbox "Modo rápido: solo reporte (sin descargar PDF ni XML)" marcado.
- Flecha 2: Multiselect "Formatos a descargar" deshabilitado y vacío.
- Flecha 3: Leyenda "Se generará un Excel con los comprobantes del período elegido".
```

#### Tipos de comprobante disponibles

**En Recibidos**:
- Retención
- Facturas
- Notas de débito
- Notas de crédito
- Liquidación de compra

**En Emitidos**:
- Facturas
- Liquidación de compra
- Guía de remisión
- Retención
- Notas de débito
- Notas de crédito

#### Campos adicionales — solo en Emitidos

| Campo | Tipo | Descripción |
|---|---|---|
| _Estado autorización_ | Lista desplegable | "Autorizados" o "No Autorizados". |
| _Establecimiento_ | Texto | (Opcional) filtrar por número de establecimiento (ej. `001`). |
| _Punto de emisión_ | Texto | (Opcional) filtrar por punto de emisión (ej. `001`). |
| _Consolidar PDF / XML_ | Checkbox | Indica qué formato descargar al ejecutar. |

⚠️ **Atención (Emitidos Autorizados — XML)**: el portal del SRI tiene un **límite operativo de 30 días** para la descarga de XML de Emitidos Autorizados. El sistema valida automáticamente este límite.

```
[CAPTURA 08 — Card Filtros — modo Recibidos]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 2.
- Flecha 2: Título "Filtros" + subtítulo "Qué descargar".
- Flecha 3: Lista Origen de comprobantes.
- Flecha 4: Lista Tipo de comprobante.
- Flecha 5: Radio Modo de fecha.
- Flecha 6: Selectores Año, Mes, Día.
- Flecha 7: Multiselect Formatos a descargar (chips XML / PDF).
```

```
[CAPTURA 09 — Card Filtros — modo Emitidos con estado]
Descripción de la captura:
- Flecha 1: Origen "Emitidos" seleccionado.
- Flecha 2: Lista Estado de autorización.
- Flecha 3: Campos Establecimiento / Punto de emisión.
- Flecha 4: Checkbox Consolidar XML y Consolidar PDF.
```

### 7.1.5 Paso 3 — Carpeta base

**Propósito**: definir la carpeta donde se almacenarán los archivos descargados y los reportes generados.

#### Campos

| Campo | Tipo | Descripción |
|---|---|---|
| _Ruta seleccionada_ | Texto (deshabilitado) | Muestra la carpeta activa. |

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Seleccionar carpeta de descarga** | Abre un selector nativo de carpetas. |
| **Guardar carpeta** | (Fallback) confirma una ruta escrita manualmente. |

### Estructura de carpetas generada

Dentro de la carpeta base se crearán subcarpetas según el origen y los criterios elegidos:

```
[Carpeta base]
├── [RUC]
│   ├── Recibidos
│   │   └── [Tipo]
│   │       └── [Año]
│   │           └── [Mes]
│   │               ├── XML
│   │               └── PDF
│   └── Emitidos
│       └── [Estado]
│           └── [Tipo]
│               └── [Año]
│                   └── [Mes]
│                       ├── XML
│                       └── PDF
└── Reportes
```

```
[CAPTURA 10 — Card Carpeta base]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 3.
- Flecha 2: Título "Carpeta base" + subtítulo "Dónde se guardan las descargas".
- Flecha 3: Campo Ruta seleccionada.
- Flecha 4: Botón Seleccionar carpeta de descarga.
- Flecha 5: Mensaje "Carpeta activa: …. Dentro se almacenarán tus descargas".
```

### 7.1.6 Paso 4 — Ejecutar

**Propósito**: iniciar o detener el proceso del robot.

#### Botones y acciones

| Botón | Color | Acción |
|---|---|---|
| **▶️ Iniciar proceso** | Verde | Comienza la descarga según los filtros definidos. |
| **⏹️ Detener proceso** | Rojo | Solicita la cancelación segura del proceso en curso. |
| **Reanudar descarga** | (visible solo si hay checkpoint) | Continúa una descarga previamente interrumpida. |
| **Descartar** | (visible solo si hay checkpoint) | Borra el checkpoint pendiente. |

#### Flujo de ejecución

1. Tras presionar **▶️ Iniciar proceso**, el robot:
   - Se autentica en el portal del SRI con el RUC y la clave.
   - Consulta los comprobantes según los filtros.
   - Descarga los archivos en la carpeta base.
   - Registra cada operación en el historial.
2. Mientras el proceso corre, aparece una barra de progreso.
3. Al finalizar, se abre el modal **"Proceso terminado"** con un resumen (totales y enlaces de descarga).
4. Si el usuario presiona **⏹️ Detener proceso**, el sistema entra en estado "cancelling" y termina la descarga en curso de forma segura.

```
[CAPTURA 11 — Card Ejecutar con botones de control]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 4.
- Flecha 2: Título "Ejecutar".
- Flecha 3: Botón Iniciar proceso (verde).
- Flecha 4: Botón Detener proceso (rojo).
- Flecha 5: Barra de progreso (durante la ejecución).
```

### 7.1.7 Modal "Proceso terminado"

Una vez que la descarga finaliza, aparece un cuadro de diálogo con el resumen.

```
[CAPTURA 12 — Modal de resumen final]
Descripción de la captura:
- Flecha 1: Título "Proceso terminado".
- Flecha 2: Conteo de XML descargados.
- Flecha 3: Conteo de PDF descargados.
- Flecha 4: Botón Descargar reporte (Excel) generado.
- Flecha 5: Botón Cerrar.
```

---

## 7.2 MÓDULO 2 — REPORTES E HISTORIAL (Tab 2)

### 7.2.1 Propósito del módulo

Generar reportes Excel a partir de comprobantes que ya están descargados en una carpeta local, y consultar el historial de ejecuciones recientes con sus resultados.

El módulo tiene tres reportes, cada uno en su propia tarjeta:

| # | Reporte | Qué cruza |
|---|---|---|
| 1 | **Reporte por fechas** | Los comprobantes de una carpeta, filtrados por rango de fechas. |
| 2 | **Notas de Crédito vs Facturas** | Cada NC contra la factura que modifica, con el valor neto. |
| 3 | **Retenciones vs Facturas** | Cada retención contra la factura que le sirve de sustento, con los días transcurridos. |

### 7.2.2 Card "Reporte por fechas"

**Propósito**: crear un Excel con los datos de los XML y PDF disponibles en la carpeta fuente, según el rango de fechas indicado.

#### Campos

| Campo | Tipo | Descripción |
|---|---|---|
| _Carpeta fuente_ | Texto | Carpeta donde ya están descargados los comprobantes. |
| _Origen_ | Lista | "Recibidos" o "Emitidos". |
| _Tipo de comprobante_ | Lista | Facturas, Retenciones, Notas de crédito, Notas de débito, Liquidación de compra, Guía de remisión. |
| _Estado de autorización_ | Lista | Solo si Origen = Emitidos ("Autorizados" / "No autorizados"). |
| _Modo de fecha_ | Radio | "Día específico", "Mes completo", "Rango de fechas", "Rango de meses", "Año completo". |
| _Fecha / Año / Mes_ | Variable | Cambia según el modo de fecha elegido. |

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Seleccionar carpeta fuente** | Abre selector de carpetas. |
| **Generar reporte por fechas** | Crea el Excel a partir de los archivos. |
| **Descargar reporte por fechas** | (Aparece al finalizar) descarga el Excel generado. |

#### Flujo

1. Indique la carpeta fuente donde ya están los XML/PDF.
2. Elija origen, tipo y estado.
3. Defina el modo de fecha y el periodo.
4. Presione **Generar reporte por fechas**.
5. Espere el mensaje de éxito.
6. Presione **Descargar reporte por fechas** para guardar el archivo Excel.

```
[CAPTURA 13 — Card Reporte por fechas]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 1 ("Reporte por fechas / Excel y PDF").
- Flecha 2: Campo Carpeta fuente.
- Flecha 3: Botón Seleccionar carpeta fuente.
- Flecha 4: Listas Origen y Tipo de comprobante.
- Flecha 5: Radio Modo de fecha.
- Flecha 6: Campo Fecha.
- Flecha 7: Botón Generar reporte por fechas.
```

### 7.2.3 Sección "Historial de ejecuciones recientes"

**Propósito**: revisar las ejecuciones previas del robot con sus filtros, estado y resultados.

#### Filtros disponibles

| Filtro | Tipo | Descripción |
|---|---|---|
| _Búsqueda_ | Texto | Búsqueda libre por RUC, descripción, estado, etc. |
| _RUC_ | Lista | Filtra por RUC. |
| _Origen_ | Lista | Recibidos / Emitidos / Todos. |
| _Tipo_ | Lista | Tipo de comprobante. |
| _Estado autorización_ | Lista | Solo aplica para Emitidos. |
| _Año / Mes / Día_ | Listas | Filtros adicionales. |

#### Columnas de la tabla

- **Fecha y hora** de la ejecución.
- **RUC** consultado.
- **Tipo de comprobante**.
- **Estado** — con pill coloreada:
  - 🟢 Verde: "ok" / "exitoso" / "completado".
  - 🟡 Ámbar: "pendiente" / "en proceso".
  - 🔴 Rojo: "error" / "fallido".
- **Autorización** — pill azul ("Autorizados") o naranja ("No autorizados").
- **Periodo** — mes/año del periodo descargado.

#### Pie de la tabla

- **Total de operaciones registradas** con un badge verde mostrando el conteo.

```
[CAPTURA 14 — Tabla de historial con filtros]
Descripción de la captura:
- Flecha 1: Encabezado "Historial de ejecuciones recientes".
- Flecha 2: Campo Búsqueda.
- Flecha 3: Filtro Origen.
- Flecha 4: Filtro Estado autorización.
- Flecha 5: Tabla con columnas Fecha, RUC, Tipo, Estado, Autorización, Periodo.
- Flecha 6: Pie "Total de operaciones registradas" con badge verde.
```

```
[CAPTURA 15 — Pills de estado en detalle]
Descripción de la captura:
- Flecha 1: Pill verde "ok".
- Flecha 2: Pill ámbar "pendiente".
- Flecha 3: Pill azul "Autorizados".
```

### 7.2.4 Card "Reporte Notas de Crédito vs Facturas (Valor Neto)"

**Propósito**: para cada Nota de Crédito emitida, identificar la Factura original que modifica y generar un Excel con el cálculo del **Valor Neto** = Importe Total Factura − Importe Total Nota de Crédito. Útil para conciliaciones contables y para detectar diferencias entre lo facturado bruto y lo realmente cobrado después de devoluciones/descuentos.

#### Cómo encuentra cada factura

El sistema busca la factura modificada en dos lugares, en orden:

1. **Localmente** — primero busca el XML/PDF de la factura en la carpeta de Notas de Crédito que indicaste o en carpetas hermanas (mismo RUC + mismo año/mes). Esto es instantáneo.
2. **Portal del SRI** — si no encuentra la factura localmente, calcula su clave de acceso (49 dígitos) a partir de los datos de la propia NC y consulta el portal de Emitidos del SRI usando la opción "Clave de acceso / Nro. autorización". Solo aparece **1 fila** por factura buscada → match infalible, sin paginación ni filtros frágiles.

⚠️ **Importante sobre la búsqueda remota**:
- Solo se ejecuta si ingresaste RUC + Clave del SRI (login válido al portal).
- El sistema calcula la clave usando el código numérico de la propia Nota de Crédito (las facturas y NCs del mismo emisor suelen usar el mismo código).
- Si una factura sigue saliendo como "no encontrada", el log diagnóstico te muestra la clave calculada para que la verifiques manualmente en el portal.

#### Campos

| Campo | Tipo | Descripción |
|---|---|---|
| _RUC del emisor de las Notas de Crédito_ | Texto | RUC con el que se emitieron las NC. |
| _Clave del SRI_ | Texto (oculto) | Contraseña SRI del mismo RUC. Solo se usa para la búsqueda remota cuando una factura no está localmente. |
| _Carpeta con Notas de Crédito ya descargadas_ | Texto + Botón | Carpeta del mes (o nivel superior) que contiene los XML/PDF de las NC. El sistema explora recursivamente. |

⚠️ Las credenciales deben corresponder al **mismo RUC** que emitió las Notas de Crédito. El sistema no reutiliza las del Módulo 1 ("Descarga de Comprobantes") para evitar incongruencias.

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Seleccionar carpeta de Notas de Crédito** | Abre selector de carpetas. |
| **📊 Generar reporte Valor Neto Facturacion Electronica** | Procesa cada NC y genera el Excel. Si hay facturas no encontradas localmente, hace login al portal del SRI para buscarlas. |
| **⬇️ Descargar reporte Valor Neto (Excel)** | (Aparece al finalizar) descarga el Excel generado. |

#### Estructura del Excel resultante

El Excel `NC_vs_Facturas_AAAAMMDD_HHMMSS.xlsx` tiene **una fila por Nota de Crédito procesada** con estas columnas:

| Columna | Contenido |
|---|---|
| A | RUC |
| B | Fecha nota de crédito |
| C | Serie nota de crédito |
| D | Clave acceso nota de crédito (49 dígitos) |
| E | Valor total nota de crédito |
| F | Factura modificada (serie EEE-PPP-SSSSSSSSS) |
| G | Fecha factura modificada |
| H | Clave acceso factura (49 dígitos, vacío si no se encontró) |
| I | Valor total factura |
| J | **Valor neto** = I − E |
| K | Estado: `OK (local)`, `OK (remoto)`, `Factura no encontrada`, `Error`, `Omitido` |
| L | Observación: detalle del estado (e.g. "Factura encontrada localmente en X.xml") |

#### Flujo

1. Ingrese el RUC y la Clave del SRI del emisor (vea la advertencia sobre el mismo RUC arriba).
2. Indique o seleccione la carpeta con las Notas de Crédito descargadas.
3. Presione **Generar reporte Valor Neto Facturacion Electronica**.
4. El sistema muestra una sección de progreso ("Procesando Notas de Crédito vs Facturas…") con mensajes en tiempo real:
   - Cantidad de NC detectadas.
   - Búsqueda local (instantánea).
   - Login al portal SRI (solo si hay pendientes).
   - Búsqueda factura por factura: `[N/T] OK 002-002-XXXXXX (DD/MM/AAAA): importe=XXX.YY`.
5. Al finalizar, se muestra el resumen: `N NC procesadas — local: X, remoto SRI: Y, no encontradas: Z, errores: W`.
6. Presione **⬇️ Descargar reporte Valor Neto (Excel)** para guardar el archivo.

#### Estados posibles por fila

| Estado | Significado | Acción recomendada |
|---|---|---|
| `OK (local)` | La factura se encontró en la misma carpeta que las NC. Valor neto calculado con datos del XML/PDF local. | Nada — todo bien. |
| `OK (remoto)` | La factura se encontró consultando el portal del SRI. | Nada — todo bien. |
| `Factura no encontrada` | Ni localmente ni en el portal. | Verificar manualmente: ¿es del mismo RUC?, ¿fue anulada?, ¿el código numérico coincide con el de la NC? |
| `Error` | No se pudo leer el XML de la NC, o falta un campo obligatorio (`numDocModificado`, `fechaEmisionDocSustento`). | Revisar el archivo XML de la NC. |
| `Omitido` | La NC no modifica una Factura (e.g., modifica una Nota de Débito). | Es normal — solo se procesan NCs que modifican Facturas (`codDocModificado == 01`). |

#### Casos típicos

**Caso A — todas locales**: si ya descargaste tanto las NC como las Facturas del mismo mes con el Módulo 1, el reporte se genera en segundos sin tocar el portal del SRI.

**Caso B — facturas viejas de meses anteriores**: las NC suelen modificar facturas de meses pasados. Esas facturas seguramente no están localmente. El sistema las busca remoto y el reporte tarda ~1-3 minutos (depende de cuántas haya).

**Caso C — código numérico distinto**: si el facturador electrónico del emisor cambió en algún momento, el código numérico puede variar entre comprobantes viejos y nuevos. En ese caso algunas facturas remotas pueden salir como "no encontradas". El log diagnóstico te indica las claves calculadas para verificar manualmente.

```
[CAPTURA 16 — Card NC vs Facturas con campos llenados]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 2 ("Reporte Notas de Crédito vs Facturas").
- Flecha 2: Campos RUC y Clave del SRI.
- Flecha 3: Campo Carpeta con Notas de Crédito.
- Flecha 4: Botón Generar reporte Valor Neto.
- Flecha 5: Sección de progreso con mensajes en tiempo real.
- Flecha 6: Botón Descargar reporte Valor Neto (Excel) tras el éxito.
```

---

### 7.2.5 Card "Reporte Retenciones vs Facturas"

**Propósito**: para cada Comprobante de Retención, ubicar la Factura que le sirve de sustento y generar un Excel que pone los datos de ambos lado a lado, cerrando con los **días transcurridos** entre la emisión de la factura y la de la retención.

#### Los dos sentidos

Lo primero que se elige es qué retenciones tenés, porque de eso depende dónde están las facturas:

| Sentido | Qué pasó | Quién emitió la factura | Dónde se busca |
|---|---|---|---|
| **Emitidas** | Vos le retuviste a tu proveedor | El proveedor | **Recibidos** |
| **Recibidas** | Te retuvieron sobre una venta | Vos | **Emitidos** |

En los dos casos, quien emitió la factura es el **sujeto retenido** del comprobante. Esa es la clave del cruce.

#### Cómo encuentra cada factura

El comprobante de retención ya trae, para cada documento de sustento, el RUC de quien emitió la factura y su número. Con esos dos datos se busca en el listado del portal — no hace falta reconstruir claves de acceso.

Además, la fecha de emisión que declara el propio comprobante indica **de qué mes es cada factura**, así que solo se consultan los días que hacen falta y no el período completo.

⚠️ **El módulo consulta el portal siempre**: por eso el RUC y la clave del SRI son obligatorios. Debe ser el RUC del contribuyente auditado — el mismo en los dos sentidos.

#### Campos

| Campo | Tipo | Descripción |
|---|---|---|
| _¿Qué retenciones tienes?_ | Radio | "Emitidas" o "Recibidas" (ver tabla arriba). |
| _Carpeta con Comprobantes de Retención ya descargados_ | Texto + Botón | Acepta PDF o XML. Puede ser el mes específico o cualquier nivel superior; se explora recursivamente. |
| _RUC del contribuyente auditado_ | Texto | RUC dueño de la carpeta de descargas. |
| _Clave del SRI_ | Texto (oculto) | Contraseña del mismo RUC. |

#### Botones y acciones

| Botón | Acción |
|---|---|
| **Seleccionar carpeta de Retenciones** | Abre selector de carpetas. |
| **Generar reporte Retenciones vs Facturas** | Lee las retenciones, consulta el portal y genera el Excel. |
| **⬇️ Descargar reporte Retenciones vs Facturas (Excel)** | (Aparece al finalizar) descarga el Excel generado. |

#### Resultado en pantalla

Al terminar se muestran cuatro tarjetas con los números de la corrida:

| Tarjeta | Significado |
|---|---|
| **Retenciones leídas** | Comprobantes procesados de la carpeta. |
| **Con factura** (verde) | Cruzaron contra una factura del portal. |
| **Sin factura electrónica** (ámbar) | El mes se revisó y la factura no figura. Normalmente son facturas preimpresas. |
| **Otro sustento** | El documento de sustento no es una factura (IFIS, notas de venta). No cruzan por definición. |

La suma de las tres últimas es igual a la primera: **una fila por comprobante**.

Debajo, el desplegable **"Ver detalle del proceso"** guarda el paso a paso completo. Está cerrado porque es información de diagnóstico, pero conviene abrirlo si algún número no cuadra.

#### Estructura del Excel resultante

El archivo `Retenciones_vs_Facturas_AAAAMMDD_HHMMSS.xlsx` tiene **dos hojas**.

**Hoja 1 — "Retenciones vs Facturas"**: una fila por factura de sustento.

| Bloque | Columnas |
|---|---|
| Retención | RUC y razón social del agente de retención, número, fecha, clave de acceso |
| Sujeto retenido | RUC y razón social (es quien emitió la factura) |
| Sustento | Tipo, número de factura, fecha declarada, ejercicio fiscal |
| Impuestos | Base, porcentaje y valor retenido de IVA y de Renta |
| Factura | Clave de acceso, nro. de autorización, fecha, subtotal, IVA, importe total, origen del dato |
| Cierre | Estado, observación y **Días entre factura y retención** |

**Hoja 2 — "Sustento no factura"**: los documentos que no son facturas, con la misma estructura. Van aparte para no ensuciar el cruce.

#### Estados posibles por fila

| Estado | Significado | Acción recomendada |
|---|---|---|
| `Factura encontrada` | Se ubicó la factura y se completaron sus importes. | Nada — todo bien. |
| `Sin factura electronica` | El mes se consultó y la factura no figura. Casi siempre es una factura preimpresa: sustenta la retención igual, pero no es electrónica. | Verificar contra el papel. No hay nada que descargar. |
| `Factura no encontrada` | No se indexó ninguna factura de ese mes. | Falta descargar ese mes — la columna "Fecha factura (segun retencion)" indica cuál. |
| `Sustento no es factura` | El sustento es un documento IFIS, nota de venta u otro. | Es normal. Va en la hoja 2. |
| `No se pudo leer el sustento` | No se reconoció el tipo en el PDF. | Formato de emisor no soportado — reportarlo con el archivo. |

#### Sobre la columna "Origen datos factura"

Dice de dónde salieron el subtotal, el IVA y el total:

- `listado` — los confirmó el portal del SRI.
- `retencion` — los declara el propio comprobante de retención. Ocurre cuando la factura no aparece en el portal; el esquema del SRI obliga a que la retención incluya esos importes, así que la fila se completa igual.

#### Casos típicos

**Caso A — retenciones emitidas**: las facturas están en Recibidos, que se consulta por mes. Es el sentido rápido: un par de consultas y listo.

**Caso B — retenciones recibidas**: las facturas están en Emitidos, y ese módulo del portal filtra por día. Tarda más, en proporción a cuántas fechas distintas tengan las facturas.

**Caso C — muchos sustentos que no son facturas**: es normal, sobre todo en retenciones recibidas de bancos y aseguradoras. Si de 97 comprobantes 79 van a la hoja 2, no es un error: son documentos IFIS que no figuran en el listado de facturas.

```
[CAPTURA 16b — Card Retenciones vs Facturas con resultado]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 3 ("Reporte Retenciones vs Facturas").
- Flecha 2: Radio "¿Qué retenciones tienes?" con los dos sentidos.
- Flecha 3: Campo Carpeta con Comprobantes de Retención.
- Flecha 4: Campos RUC y Clave del SRI.
- Flecha 5: Las cuatro tarjetas de resultado (verde en "Con factura", ámbar en "Sin factura electrónica").
- Flecha 6: Desplegable "Ver detalle del proceso" cerrado.
- Flecha 7: Botón Descargar reporte (Excel).
```

---

## 7.3 MÓDULO 3 — CONSOLIDACIÓN DE DOCUMENTOS (Tab 3)

### 7.3.1 Propósito del módulo

Consolidar reportes mensuales previamente generados en un reporte único anual o por periodo, y copiar los documentos correspondientes a una carpeta de "Consolidados".

### 7.3.2 Card 1 — Carpeta origen a consolidar

**Propósito**: indicar dónde están los documentos previamente descargados.

#### Campos y botones

| Elemento | Tipo | Descripción |
|---|---|---|
| _Carpeta origen para consolidar_ | Texto (deshabilitado) | Ruta activa. |
| **Seleccionar carpeta para consolidar** | Botón | Abre el selector de carpetas. |
| **Usar carpeta de descargas activa** | Botón | Asigna como carpeta de consolidación la misma que el módulo de descarga. |
| _Ruta de carpeta para consolidar (manual)_ | Texto | (Fallback) si el selector nativo no está disponible. |
| **Guardar carpeta de consolidacion** | Botón | Confirma la ruta manual. |

```
[CAPTURA 16 — Card Carpeta origen]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 1 ("Carpeta origen a consolidar").
- Flecha 2: Campo de ruta.
- Flecha 3: Botón Seleccionar carpeta para consolidar.
- Flecha 4: Botón Usar carpeta de descargas activa.
- Flecha 5: Mensaje "Carpeta de búsqueda activa para consolidación".
```

### 7.3.3 Card 2 — Filtros

**Propósito**: definir qué consolidar.

#### Campos

| Campo | Tipo | Descripción |
|---|---|---|
| _RUC a buscar (opcional)_ | Texto | Si se deja vacío, consolida toda la carpeta seleccionada. |
| _Origen a consolidar_ | Lista | "Recibidos" o "Emitidos". |
| _Tipo de comprobante_ | Lista | Igual que en módulo 1. |
| _Año a consolidar_ | Numérico | Año del periodo. |
| _Estado autorización_ | Lista | Solo Emitidos. |
| _Modo de fecha_ | Radio | "Mes y día", "Rango de meses", "Año completo". |
| _Mes / Mes inicio / Mes fin / Día_ | Variables | Según el modo seleccionado. |

```
[CAPTURA 17 — Card Filtros]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 2 ("Filtros / RUC, origen, tipo, año y periodo").
- Flecha 2: Campo RUC a buscar.
- Flecha 3: Lista Origen.
- Flecha 4: Lista Tipo.
- Flecha 5: Campo Año.
- Flecha 6: Radio Modo de fecha.
- Flecha 7: Selectores de mes y día.
```

### 7.3.4 Card 3 — Salida

**Propósito**: definir qué formatos consolidar y ejecutar.

#### Campos y botones

| Elemento | Tipo | Descripción |
|---|---|---|
| _Consolidar XML_ | Checkbox | Activa la consolidación de archivos XML. |
| _Consolidar PDF_ | Checkbox | Activa la consolidación de archivos PDF. |
| **Consolidar desde carpeta** | Botón verde | Ejecuta la consolidación. |

#### Resultado

Al ejecutarse, el sistema:

1. Busca los reportes existentes y/o los documentos en la carpeta origen.
2. Genera reportes Excel consolidados (`recibidos_reporte_…xlsx` / `emitidos_reporte_…xlsx`).
3. Crea una carpeta `Consolidados/[Origen]/[Estado]/[Periodo]` con subcarpetas `XML/` y `PDF/`.
4. Copia los documentos únicos encontrados.
5. Muestra mensajes de éxito con la ruta del Excel y la cantidad de archivos copiados.

```
[CAPTURA 18 — Card Salida y resultados]
Descripción de la captura:
- Flecha 1: Insignia verde con el número 3 ("Salida / Formatos y ejecución").
- Flecha 2: Checkbox Consolidar XML.
- Flecha 3: Checkbox Consolidar PDF.
- Flecha 4: Botón Consolidar desde carpeta.
- Flecha 5: Mensajes de éxito con la ruta del reporte consolidado.
```

---

## 7.4 MÓDULO 4 — AYUDA (Tab 4)

### 7.4.1 Quickstart (pasos 01-04)

Cuatro tarjetas resumen los pasos básicos para empezar:

| Tarjeta | Título | Resumen |
|---|---|---|
| 01 | Inicia sesión | Verifica correo, contraseña y licencia activa. |
| 02 | Descarga comprobantes | Ingresa RUC, clave del SRI y filtros. |
| 03 | Genera o consolida | Elige formato (XML/PDF) según lo que necesites. |
| 04 | Revisa el historial | Consulta el estado de cada ejecución. |

### 7.4.2 Acordeón de temas frecuentes

Contiene secciones desplegables con orientación rápida:

- **Inicio de sesión y licencias**: qué hacer si no puede entrar.
- **Descarga de Recibidos**: cómo usar el modo de fecha.
- **Descarga de Emitidos**: cómo definir estado, establecimiento y punto de emisión.
- **Consolidación desde carpeta**: cómo unir reportes.
- **Errores frecuentes y qué hacer**: timeouts, captcha, sin resultados, sin permisos.

### 7.4.3 Botones de tour

| Botón | Acción |
|---|---|
| **Activar tour de primer uso** | Inicia el recorrido guiado. |
| **Marcar tour como no visto** | Restablece el estado para que el tour vuelva a aparecer. |

### 7.4.4 Card "Acerca de la aplicación"

Muestra información sobre la versión y permite buscar actualizaciones.

#### Información mostrada

| Dato | Descripción |
|---|---|
| _Versión actual_ | Versión instalada (ej. `v2026.05.05.88`). |
| _Modo_ | "compilado (.exe)" o "desarrollo". |

#### Botón

| Botón | Acción |
|---|---|
| **🔄 Buscar actualizaciones** | Consulta al servidor si hay una versión más reciente. |

```
[CAPTURA 19 — Quickstart cards]
Descripción de la captura:
- Flecha 1: Tarjeta 01 Inicia sesión.
- Flecha 2: Tarjeta 02 Descarga comprobantes.
- Flecha 3: Tarjeta 03 Genera o consolida.
- Flecha 4: Tarjeta 04 Revisa el historial.
```

```
[CAPTURA 20 — Acordeón de ayuda]
Descripción de la captura:
- Flecha 1: Sección expandida "Inicio de sesión y licencias".
- Flecha 2: Sección colapsada "Descarga de Recibidos".
- Flecha 3: Sección colapsada "Errores frecuentes".
```

```
[CAPTURA 21 — Card Acerca de + actualizaciones]
Descripción de la captura:
- Flecha 1: Insignia "i" del card "Acerca de la aplicación".
- Flecha 2: Versión actual.
- Flecha 3: Modo (compilado .exe / desarrollo).
- Flecha 4: Botón 🔄 Buscar actualizaciones.
```

---

# 8. PROCEDIMIENTOS PASO A PASO

## 8.1 Cómo iniciar sesión por primera vez

1. Abra la aplicación (icono **SRI Robot Audit** del escritorio si tiene la versión .exe, o la URL proporcionada para la versión web).
2. En la pantalla de **Iniciar sesión**, ingrese su correo electrónico.
3. Ingrese su contraseña.
4. Presione **Iniciar sesión**.
5. Si es su primer ingreso desde este equipo, continúe con el procedimiento 8.2.

## 8.2 Cómo activar una licencia nueva

1. Tras el login, aparece la pantalla **Activación de licencia**.
2. Solicite a su administrador el **código de licencia** que le corresponde.
3. Pegue el código en el campo _Código de licencia_.
4. Verifique que el campo _Identificador del equipo_ esté rellenado automáticamente (no lo modifique).
5. Presione **Activar licencia**.
6. Si la activación es correcta, el sistema lo lleva al menú principal.

⚠️ Si ya activó la licencia en otro equipo, contacte a su administrador para que la libere antes de activarla aquí.

## 8.3 Cómo recuperar la contraseña

1. En la pantalla de inicio de sesión, presione **¿Olvidaste tu contraseña?**.
2. Ingrese el correo electrónico registrado.
3. Presione **Enviar enlace**.
4. Revise su correo (incluida la carpeta de spam).
5. Abra el enlace recibido.
6. Defina la nueva contraseña (mínimo 8 caracteres) y confírmela.
7. Presione **Guardar contraseña**.
8. Inicie sesión con la nueva contraseña.

## 8.4 Cómo descargar comprobantes Recibidos (XML + PDF) de un mes

1. Vaya a la pestaña **Descarga de Comprobantes**.
2. **Paso 1 — Credenciales**: ingrese su _RUC_ y _Clave del SRI_.
3. **Paso 2 — Filtros**:
   - _Origen_: seleccione "Recibidos".
   - _Tipo de comprobante_: seleccione, por ejemplo, "Facturas".
   - _Modo de fecha_: "Mes y día".
   - _Año_: ingrese el año.
   - _Mes_: elija el mes.
   - _Día (0 = Todos)_: deje en `0` para descargar el mes completo.
   - _Formatos a descargar_: seleccione "XML" y "PDF".
4. **Paso 3 — Carpeta base**: confirme la carpeta destino (o cambie con **Seleccionar carpeta de descarga**).
5. **Paso 4 — Ejecutar**: presione **▶️ Iniciar proceso**.
6. Espere a que termine. Aparecerá el modal **Proceso terminado** con el resumen.

## 8.5 Cómo descargar comprobantes Emitidos Autorizados (con límite de 30 días)

1. En la pestaña **Descarga de Comprobantes**, ingrese sus credenciales.
2. En **Filtros**:
   - _Origen_: "Emitidos".
   - _Tipo de comprobante_: el que necesite.
   - _Estado de autorización_: "Autorizados".
   - _Establecimiento_ y _Punto de emisión_: (opcional) acote por sucursal.
   - _Modo de fecha_: cualquiera. Si elige un rango mayor a 30 días, el sistema lo dividirá automáticamente respetando el límite del portal.
3. Defina la _Carpeta base_.
4. Presione **▶️ Iniciar proceso**.

ℹ️ Para XML de Emitidos Autorizados, el sistema **respeta automáticamente el límite operativo de 30 días** definido por el portal del SRI.

## 8.6 Cómo reanudar una descarga interrumpida

Si una descarga fue interrumpida (por cierre de la app, pérdida de conexión, etc.), al abrir la pestaña **Descarga de Comprobantes** aparece un aviso con dos botones:

1. Presione **Reanudar descarga** para continuar desde el punto donde quedó.
2. Si prefiere empezar desde cero, presione **Descartar** y luego configure y ejecute una nueva descarga.

## 8.7 Cómo generar un reporte por fechas

1. Vaya a la pestaña **Reportes e Historial**.
2. En la card **Reporte por fechas**, indique la _Carpeta fuente_ (donde ya están los XML/PDF).
3. Seleccione _Origen_ y _Tipo de comprobante_.
4. Elija el _Modo de fecha_ y rellene los campos correspondientes.
5. Presione **Generar reporte por fechas**.

## 8.8 Cómo descargar el Excel generado

1. Tras generar el reporte (procedimiento 8.7), aparece un mensaje de éxito con el resumen.
2. Presione **Descargar reporte por fechas**.
3. Elija la ubicación de guardado en su equipo.

## 8.9 Cómo consultar el historial y filtrar resultados

1. Vaya a la pestaña **Reportes e Historial**.
2. Desplácese hasta la sección **Historial de ejecuciones recientes**.
3. Use los filtros (_Búsqueda_, _Origen_, _Estado_, _RUC_, _Año_, etc.) para acotar.
4. La tabla se actualiza automáticamente.
5. Revise el conteo total en el badge verde inferior.

## 8.10 Cómo consolidar documentos ya descargados

1. Vaya a la pestaña **Consolidación de documentos**.
2. En la card **Carpeta origen a consolidar**, seleccione la carpeta donde están los documentos.
3. En la card **Filtros**:
   - Ingrese el _RUC_ (opcional).
   - Elija _Origen_, _Tipo_, _Año_ y _Estado_ (este último solo Emitidos).
   - Defina el _Modo de fecha_.
4. En la card **Salida**:
   - Marque _Consolidar XML_, _Consolidar PDF_ o ambos.
   - Presione **Consolidar desde carpeta**.
5. Lea los mensajes de éxito y revise la carpeta `Consolidados/` generada.

## 8.11 Cómo cambiar el tema claro/oscuro

1. Localice el botón **☀️** o **🌙** en la parte superior derecha del topbar.
2. Haga clic. El tema se alterna instantáneamente.

## 8.12 Cómo activar el tour interactivo

**Opción A — Desde la pestaña Descarga de Comprobantes**:
1. Presione el botón **● Primera vez? Ver tour**.

**Opción B — Desde la pestaña Ayuda**:
1. Vaya a **Ayuda**.
2. Presione **Activar tour de primer uso**.

El tour le mostrará paso a paso las funciones principales. Use los botones **Anterior**, **Siguiente**, **Mas tarde**, **Finalizar** u **Omitir y no mostrar** para navegar.

## 8.13 Cómo buscar y aplicar una actualización

1. Vaya a la pestaña **Ayuda**.
2. Desplácese al card **Acerca de la aplicación**.
3. Presione **🔄 Buscar actualizaciones**.
4. El sistema consulta al servidor.
5. Si hay una versión nueva (solo en la versión .exe), siga las instrucciones en pantalla para descargarla e instalarla.

ℹ️ En la versión web, las actualizaciones se aplican automáticamente en el servidor.

## 8.14 Cómo cerrar sesión y cerrar la aplicación

1. Presione el botón **👤 Perfil** en el topbar.
2. En el popover:
   - Para cerrar la sesión: presione **🚪 Cerrar sesión**.
   - Para cerrar la aplicación (.exe): presione **⏻ Cerrar app** y confirme en el cuadro de diálogo.

---

## 8.15 Cómo obtener solo el listado de un período, sin descargar comprobantes

1. Vaya a la pestaña **Descarga de Comprobantes**.
2. Complete el Paso 1 (Credenciales) con su RUC y clave del SRI.
3. En el Paso 2 (Filtros) elija el origen, el tipo de comprobante y el período.
4. Marque **"Modo rápido: solo reporte (sin descargar PDF ni XML)"**.
   - La selección de formatos queda deshabilitada: es correcto.
5. Presione **Iniciar descarga**.
6. Al terminar, use el botón de descarga del Excel.

El archivo queda además en `[Carpeta base]/[RUC]/[Origen]/[Tipo]/[Año]/[Mes]/TXT/`.

⚠️ Este modo **no descarga comprobantes**. Si después necesita los XML o PDF, vuelva a ejecutar sin el modo rápido.

---

## 8.16 Cómo generar el reporte de Retenciones vs Facturas

**Antes de empezar** necesita las retenciones descargadas. Si no las tiene:

1. Pestaña **Descarga de Comprobantes** → Tipo **Comprobante de Retención**.
2. Origen **Emitidos** o **Recibidos**, según cuáles necesite.
3. Formato **PDF** (para meses antiguos es lo único disponible) o **XML**.

**El reporte**:

1. Vaya a **Reportes e Historial** → card **Reporte Retenciones vs Facturas**.
2. Elija el sentido:
   - *Emitidas* si usted le retuvo a su proveedor.
   - *Recibidas* si le retuvieron sobre una venta.
3. Indique la carpeta con los comprobantes de retención.
4. Ingrese el **RUC del contribuyente auditado** y su clave del SRI.
5. Presione **Generar reporte Retenciones vs Facturas**.
6. Revise las cuatro tarjetas de resultado. La suma de las tres últimas debe dar la primera.
7. Descargue el Excel.

💡 Si la tarjeta ámbar ("Sin factura electrónica") tiene valores, abra el desplegable **"Ver detalle del proceso"**: ahí figura si el mes se revisó completo.

---

## 8.17 Qué hacer si una factura no aparece en el reporte de Retenciones

Mire la columna **Estado** de esa fila:

| Estado | Qué significa | Qué hacer |
|---|---|---|
| `Sin factura electronica` | El mes se consultó y la factura no está en el portal. | Casi siempre es una factura preimpresa. Verifíquela contra el papel; no hay nada que descargar. |
| `Factura no encontrada` | No se indexó ninguna factura de ese mes. | Vuelva a generar el reporte. Si persiste, revise el detalle del proceso. |
| `Sustento no es factura` | El sustento es un IFIS o una nota de venta. | Es normal. Esa fila está en la hoja 2 del Excel. |
| `No se pudo leer el sustento` | No se reconoció el formato del PDF. | Reporte el archivo a soporte. |

---

# 9. TABLA DE MENSAJES DE ERROR Y SOLUCIÓN

## 9.1 Errores de autenticación

| Mensaje | Causa | Solución |
|---|---|---|
| "Credenciales incorrectas." | Correo o contraseña incorrectos. | Verifique el correo y la contraseña. Use **¿Olvidaste tu contraseña?** si es necesario. |
| "Usuario inactivo." | Su cuenta fue desactivada por el administrador. | Contacte al administrador del sistema. |
| "Tu sesión ya no es válida. Inicia sesión nuevamente." | La sesión expiró o el token fue invalidado. | Vuelva a iniciar sesión. |
| "Error al autenticar: …" | Error de red o servidor caído. | Verifique su conexión a internet. Reintente en unos minutos. |
| "Completa todos los campos." | Quedó un campo vacío. | Rellene correo y contraseña. |

## 9.2 Errores de licencia

| Mensaje | Causa | Solución |
|---|---|---|
| "Licencia no encontrada." | El código ingresado no corresponde a ninguna licencia del usuario. | Verifique con el administrador. |
| "Licencia desactivada." | El administrador desactivó la licencia. | Contacte al administrador. |
| "Esta licencia ya está activada en otro equipo." | La licencia está vinculada a otra computadora. | Solicite al administrador liberar la licencia anterior. |
| "Licencia expirada." | La licencia venció. | Renueve con el administrador. |
| "Licencia no válida o no encontrada." | El fingerprint no coincide con la licencia. | Active nuevamente la licencia desde este equipo. |
| "Debes ingresar tu código de licencia." | El campo quedó vacío. | Ingrese el código entregado por el administrador. |
| "No se pudo activar la licencia: …" | Error de red o servidor. | Verifique su conexión y reintente. |

## 9.3 Errores de recuperación

| Mensaje | Causa | Solución |
|---|---|---|
| "No se pudo enviar el correo de recuperación: …" | Error en el servicio de correo. | Reintente más tarde. Si persiste, contacte a soporte. |
| "Ingresa el correo registrado." | Campo vacío. | Ingrese su correo. |
| "Completa todos los campos." | Falta alguna contraseña. | Rellene los dos campos. |
| "Las contraseñas no coinciden." | _Nueva contraseña_ ≠ _Confirmar contraseña_. | Vuelva a ingresarlas iguales. |
| "No se pudo actualizar la contraseña: …" | Token expirado o ya usado. | Solicite un nuevo enlace de recuperación. |
| "Abre el enlace de recuperación desde tu correo para continuar." | El usuario llegó a la pantalla sin token válido. | Abra el enlace desde el correo. |

## 9.4 Errores del robot SRI

| Mensaje / Síntoma | Causa | Solución |
|---|---|---|
| Timeout o portal lento | El portal del SRI está saturado. | Espere y vuelva a intentar en unos segundos. |
| Captcha incorrecta | El robot no pudo resolver el captcha. | Espere 1–2 minutos y reintente. |
| Sin resultados | Filtros muy restrictivos o sin movimientos en el periodo. | Valide rango de fechas, tipo y estado. |
| No descarga archivos | Sin permisos sobre la carpeta o sin espacio. | Verifique permisos de escritura y espacio libre en disco. |

## 9.5 Errores del reporte por fechas

| Mensaje | Causa | Solución |
|---|---|---|
| "La carpeta fuente no existe. Selecciona una ruta válida." | La carpeta indicada no existe o se eliminó. | Use **Seleccionar carpeta fuente** y elija una ruta válida. |
| "Debes definir un rango de fechas válido." | Falta fecha de inicio o fin. | Complete las fechas. |
| "La fecha inicio no puede ser mayor que la fecha fin." | Rango invertido. | Corrija las fechas. |

## 9.6 Errores de consolidación

| Mensaje | Causa | Solución |
|---|---|---|
| "La carpeta seleccionada no existe: …" | La carpeta origen no existe. | Vuelva a seleccionarla. |
| "Selecciona al menos una opción: XML o PDF." | Ningún checkbox marcado. | Marque al menos uno. |
| "El mes fin debe ser mayor o igual al mes inicio." | Rango de meses invertido. | Corrija el rango. |
| "No se encontró una carpeta válida para buscar reportes: …" | La estructura interna no coincide con lo esperado. | Verifique que la carpeta tenga los reportes mensuales correctos. |
| "No se pudo generar el reporte PDF consolidado." | Reportes PDF inválidos o corruptos. | Revise los archivos fuente. |
| "No se encontraron insumos XML para consolidar." | No hay XML para el periodo. | Verifique los filtros o descargue primero los XML. |

---

# 10. PREGUNTAS FRECUENTES (FAQ)

## 10.1 ¿La aplicación guarda mi clave del SRI?

No. La clave del SRI se utiliza únicamente durante la sesión activa para que el robot pueda autenticarse en el portal. **No se almacena en disco.** Cada vez que cierre y vuelva a abrir la aplicación deberá ingresarla nuevamente.

## 10.2 ¿Puedo usar mi cuenta en varios equipos?

Cada licencia está vinculada a un único equipo (identificado por su fingerprint). Si necesita cambiar de equipo, contacte al administrador para que libere la licencia del equipo anterior y la active en el nuevo.

## 10.3 ¿Por qué los XML de Emitidos solo bajan 30 días?

Por una limitación operativa del portal del SRI. La aplicación respeta ese límite automáticamente cuando descarga XML de Emitidos Autorizados, dividiendo internamente los rangos largos en bloques permitidos.

## 10.4 ¿Qué pasa si pierdo la conexión durante una descarga?

El sistema guarda un checkpoint automático. Al volver a abrir la aplicación, aparecerá un aviso con la opción **Reanudar descarga** para continuar desde donde quedó.

## 10.5 ¿Dónde se guardan los archivos descargados?

En la carpeta indicada en el paso **3 — Carpeta base**. La aplicación crea una estructura jerárquica por RUC, origen, tipo, año y mes (ver sección 6.1.5).

## 10.6 ¿Cómo se diferencia un comprobante XML de un RIDE/PDF?

- El **XML** es el documento electrónico oficial firmado.
- El **RIDE/PDF** es la representación impresa (legible para humanos).
Ambos pueden descargarse en la misma ejecución activando los dos formatos en el filtro **Formatos a descargar**.

## 10.7 ¿Puedo descargar comprobantes de meses anteriores?

Sí, el portal del SRI permite consultar varios años atrás. Para periodos anteriores a 2015 podrían existir limitaciones del propio portal.

## 10.8 ¿Cómo sé cuál es la última versión disponible?

Vaya a **Ayuda → Acerca de la aplicación** y presione **🔄 Buscar actualizaciones**. La versión actual instalada se muestra siempre en esa misma tarjeta.

## 10.9 ¿La aplicación funciona sin internet?

No. Se requiere conexión para:
- Autenticarse en el servidor de licencias.
- Conectarse al portal del SRI para descargar comprobantes.
- Validar la licencia.

Para tareas locales (generación de reportes desde archivos ya descargados, consolidación) sí se puede trabajar offline, pero el login y la validación inicial requieren conexión.

## 10.10 ¿Cómo solicito una licencia nueva?

Contacte al administrador del sistema (ver sección **10. Soporte**). Le entregará un código de licencia que deberá ingresar en la pantalla de **Activación de licencia**.

---

## 10.11 ¿Cuál es la diferencia entre el modo rápido y una descarga normal?

El modo rápido genera **solo el Excel del listado**, con los datos que el portal ya muestra en pantalla: comprobante, serie, RUC del emisor, clave de acceso, fechas, subtotal, IVA e importe total.

La descarga normal baja **un archivo por comprobante** (XML, PDF o ambos) y además arma sus reportes.

Use el modo rápido cuando necesite ver o cuadrar el período. Use la descarga normal cuando necesite los comprobantes en sí: para respaldo, para el módulo de Consolidación, o para el reporte de Retenciones vs Facturas.

---

## 10.12 En Retenciones vs Facturas, ¿qué RUC y clave debo poner?

Siempre los del **contribuyente auditado** — el dueño de la carpeta de descargas. Es el mismo en los dos sentidos:

- **Retenciones emitidas**: el RUC que emitió las retenciones.
- **Retenciones recibidas**: el RUC al que le retuvieron, que es quien emitió las facturas de sustento.

Nunca el del proveedor ni el del cliente.

---

## 10.13 ¿Por qué muchos sustentos salen como "no es factura"?

Porque no todo comprobante de retención se sustenta en una factura. Los bancos y aseguradoras emiten **documentos IFIS**, y también existen notas de venta y liquidaciones de compra.

Esos documentos son válidos como sustento pero **no figuran en el listado de facturas del portal**, así que no pueden cruzarse. Van a la hoja 2 del Excel para no ensuciar el cruce.

Es habitual que sean mayoría en retenciones recibidas.

---

## 10.14 ¿Por qué el sentido "Recibidas" tarda más?

Porque las facturas se buscan en **Emitidos**, y ese módulo del portal filtra por un solo día. El sistema consulta únicamente las fechas que las propias retenciones indican, pero aun así son varias consultas.

En cambio "Emitidas" busca en **Recibidos**, donde una sola consulta resuelve el mes completo.

---

## 10.15 Una factura existe en el portal pero el reporte dice que no la encontró

Revise la columna **Estado**:

- Si dice `Sin factura electronica`, el mes **sí** se revisó y la factura no figura en el listado electrónico. Lo más probable es que sea una **factura preimpresa**: sustenta la retención igual, pero no es electrónica y por definición no aparece.
- Si dice `Factura no encontrada`, no se indexó ninguna factura de ese mes. Vuelva a generar el reporte y revise el detalle del proceso.

En ambos casos la fila conserva todos los datos de la retención, así que puede resolverla a mano.

---

# 11. SOPORTE

## 11.1 Datos de contacto del soporte

| Canal | Detalle |
|---|---|
| Correo electrónico | _A completar_ |
| WhatsApp | _A completar_ |
| Mesa de ayuda | _A completar_ |
| Horario de atención | _A completar_ (sugerido: lunes a viernes, 09:00–18:00) |

## 11.2 Información a incluir al reportar un problema

Para una atención más rápida, al contactar al soporte incluya:

1. **Versión del sistema** (visible en **Ayuda → Acerca de la aplicación**).
2. **Modo** (compilado .exe o desarrollo).
3. **Correo electrónico** con el que inició sesión.
4. **Mensaje de error** completo (captura de pantalla).
5. **Pasos para reproducir** el problema.
6. **Logs** del launcher (solo .exe): el archivo `desktop_launcher.log` se ubica junto al ejecutable.

## 11.3 Canales

_A completar con los canales oficiales que su empresa designe._

## 11.4 Horarios de atención

_A completar con los horarios de atención al cliente._

---

# 12. ANEXOS

## Anexo A — Glosario completo

| Término | Definición |
|---|---|
| **API** | Application Programming Interface. Interfaz por la cual la aplicación se comunica con el servidor de licencias. |
| **Captcha** | Prueba antibots que aparece en el portal del SRI. |
| **Checkpoint** | Punto de control automático que permite reanudar descargas interrumpidas. |
| **Comprobante** | Documento electrónico tributario (factura, retención, nota de crédito, etc.). |
| **Consolidación** | Proceso de unificar varios reportes mensuales o por periodo. |
| **Fingerprint** | Huella digital única del equipo donde corre la aplicación. |
| **JWT** | JSON Web Token. Token de sesión que devuelve el servidor de licencias. |
| **Licencia** | Permiso digital que habilita el uso del software en un equipo concreto. |
| **OCR** | Optical Character Recognition. Tecnología que extrae texto de imágenes (usada por el robot para leer PDFs). |
| **Pestaña / Tab** | Cada una de las cuatro secciones principales de la app. |
| **Popover** | Pequeño panel que aparece sobre un botón al hacer clic. |
| **RIDE** | Representación Impresa del Documento Electrónico. Versión PDF del XML. |
| **Robot** | Componente automatizado que navega y descarga del portal del SRI. |
| **SRI** | Servicio de Rentas Internas del Ecuador. |
| **Topbar** | Barra superior fija del sistema (logo + título + controles de la derecha). |
| **Tour** | Recorrido guiado paso a paso por las funciones de la app. |
| **XML** | eXtensible Markup Language. Formato oficial del comprobante electrónico. |

## Anexo B — Estructura de carpetas generada

Cuando el sistema descarga, organiza los archivos así (dentro de la **Carpeta base** definida en el paso 3):

```
[Carpeta base]
├── [RUC]
│   ├── Recibidos
│   │   └── [Tipo de comprobante]
│   │       └── [Año]
│   │           └── [Mes]
│   │               ├── XML/
│   │               └── PDF/
│   └── Emitidos
│       └── [Estado autorización]
│           └── [Tipo de comprobante]
│               └── [Año]
│                   └── [Mes]
│                       ├── XML/
│                       └── PDF/
├── Consolidados
│   └── [Origen]
│       └── [Estado] (solo Emitidos)
│           └── [Periodo]
│               ├── XML/
│               ├── PDF/
│               └── *.xlsx (reportes consolidados)
├── Reportes
│   └── *.xlsx
└── historiales/
    └── historial_<device_id>.json
```

## Anexo C — Convenciones de nombres de archivos generados

| Tipo de archivo | Patrón aproximado |
|---|---|
| Reporte mensual Recibidos | `recibidos_reporte_xml_<tipo>_<año><mes>.xlsx` |
| Reporte mensual Emitidos | `emitidos_reporte_xml_<tipo>_<año><mes>.xlsx` |
| Reporte consolidado | `recibidos_reporte_xml_<tipo>_<periodo>.xlsx` / `emitidos_reporte_pdf_<tipo>_<periodo>.xlsx` |
| Historial por equipo | `historial_<device_id>.json` |
| Log del launcher | `desktop_launcher.log` |

## Anexo D — Mapa de pantallas (diagrama de navegación)

```
┌────────────────────────────────────┐
│      Inicio de sesión              │
│      (login_form)                  │
└────────────────────────────────────┘
   │                       │
   │ ¿Olvidaste            │ Iniciar sesión
   │   contraseña?         │
   ▼                       ▼
┌──────────────────┐   ┌──────────────────────────┐
│ Recuperar        │   │ Activación de licencia   │
│ contraseña       │   │ (solo si no activada)    │
└──────────────────┘   └──────────────────────────┘
        │                          │
        │ (correo con enlace)      │ Activar licencia
        ▼                          ▼
┌──────────────────┐   ┌────────────────────────────────────────┐
│ Restablecer      │   │              MENÚ PRINCIPAL            │
│ contraseña       │   │  Topbar: logo · título · ☀️/🌙 · Perfil│
└──────────────────┘   ├────────────────────────────────────────┤
                       │ Tab 1: Descarga de Comprobantes        │
                       │   1·Credenciales 2·Filtros             │
                       │   3·Carpeta base 4·Ejecutar            │
                       ├────────────────────────────────────────┤
                       │ Tab 2: Reportes e Historial            │
                       │   Reporte por fechas                   │
                       │   Historial de ejecuciones recientes   │
                       ├────────────────────────────────────────┤
                       │ Tab 3: Consolidación de documentos     │
                       │   1·Carpeta origen 2·Filtros 3·Salida  │
                       ├────────────────────────────────────────┤
                       │ Tab 4: Ayuda                           │
                       │   Quickstart 01-04                     │
                       │   Acordeón de FAQ                      │
                       │   Tour                                 │
                       │   Acerca de + Buscar actualizaciones   │
                       └────────────────────────────────────────┘
                                     │
                                     │ Popover Perfil
                                     ▼
                       ┌──────────────────────────────┐
                       │ 🚪 Cerrar sesión             │
                       │ ⏻  Cerrar app (solo .exe)   │
                       └──────────────────────────────┘
```

---

**FIN DEL MANUAL**

_Documento generado a partir del código fuente verificado del proyecto SRI Robot Audit, versión 2026.05.05.88._
