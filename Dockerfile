# --------------------------------------------------------
# SRI ROBOT AUDIT — Dockerfile
# Compatible con Render.com + Chromium (Playwright)
# Endurecido: usuario no-root, healthcheck, caché de pip limpia.
# --------------------------------------------------------

FROM python:3.11-slim

# Evita prompts y logs truncados
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

WORKDIR /app

# --------------------------------------------------------
# Librerías del sistema necesarias para Chromium / Playwright.
# Se instalan en una sola capa y se limpian las listas de apt para
# reducir el tamaño final de la imagen.
# --------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget gnupg unzip curl fonts-liberation \
    libnss3 libnspr4 libdbus-1-3 libatk1.0-0 libatk-bridge2.0-0 libcups2 \
    libdrm2 libxkbcommon0 libxcomposite1 libxrandr2 libxdamage1 libxfixes3 \
    libgbm1 libgtk-3-0 libpango-1.0-0 libcairo2 libasound2 \
    libx11-6 libx11-xcb1 libxext6 libxrender1 libxi6 libxtst6 \
    libwayland-client0 libwayland-cursor0 libwayland-egl1 libegl1 libgl1 \
    libxshmfence1 libglib2.0-0 libxss1 xvfb \
    gcc python3-dev libxml2-dev libxslt1-dev && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# --------------------------------------------------------
# Playwright + Chromium. Se instalan como root porque algunas libs
# requieren permisos del sistema, pero el directorio se hace legible
# para el usuario no-root que ejecutará la app.
# --------------------------------------------------------
RUN pip install --upgrade pip==24.2 setuptools wheel && \
    pip install playwright==1.47.0 && \
    python -m playwright install chromium && \
    chmod -R a+rX /root /root/.cache

# --------------------------------------------------------
# Dependencias del proyecto (en su propia capa para aprovechar la
# caché de Docker entre builds).
# --------------------------------------------------------
COPY requirements.txt .
RUN pip install -r requirements.txt

# --------------------------------------------------------
# Copiar el resto del proyecto y preparar directorios.
# --------------------------------------------------------
COPY . .
RUN mkdir -p /app/descargas /app/historiales

# --------------------------------------------------------
# Variables Playwright
# --------------------------------------------------------
ENV PLAYWRIGHT_BROWSERS_PATH=/root/.cache/ms-playwright \
    PYPPETEER_HOME=/root/.cache/ms-playwright

# --------------------------------------------------------
# Crear usuario no-root y traspasarle la propiedad de /app.
# UID 1001 evita colisión con UIDs reservados del sistema.
# --------------------------------------------------------
RUN groupadd --system --gid 1001 appuser && \
    useradd --system --uid 1001 --gid appuser --shell /sbin/nologin appuser && \
    chown -R appuser:appuser /app

USER appuser

# --------------------------------------------------------
# Healthcheck — Streamlit expone /_stcore/health.
# Render ignora esto y usa healthCheckPath de render.yaml,
# pero sirve para builds locales y otros entornos (compose, k8s).
# --------------------------------------------------------
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD curl -fsS http://localhost:8501/_stcore/health || exit 1

EXPOSE 8501
CMD ["streamlit", "run", "aplicacion.py", "--server.port=8501", "--server.address=0.0.0.0"]
