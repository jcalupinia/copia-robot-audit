# --------------------------------------------------------
# 🤖 SRI ROBOT AUDIT — DOCKERFILE FINAL DEFINITIVO
# Compatible con Render.com + Chromium instalado correctamente
# --------------------------------------------------------

FROM python:3.11-slim

# Evita prompts y logs truncados
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# Crear carpeta de trabajo
WORKDIR /app

# Copiar dependencias
COPY requirements.txt .

# --------------------------------------------------------
# 🧠 Instalar librerías del sistema necesarias para Chromium y Playwright
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
# 🧩 Instalar Playwright y navegador Chromium dentro de la imagen
# --------------------------------------------------------
RUN pip install --upgrade pip==24.2 setuptools wheel && \
    pip install --no-cache-dir playwright==1.47.0 && \
    python -m playwright install chromium && \
    chmod -R 777 /root/.cache/ms-playwright

# --------------------------------------------------------
# 📦 Instalar dependencias del proyecto
# --------------------------------------------------------
RUN pip install --no-cache-dir -r requirements.txt

# --------------------------------------------------------
# 📂 Copiar el resto del proyecto
# --------------------------------------------------------
COPY . .

# Crear carpeta de descargas
RUN mkdir -p /app/descargas

# --------------------------------------------------------
# ⚙️ Variables de entorno Playwright
# --------------------------------------------------------
ENV PLAYWRIGHT_BROWSERS_PATH=/root/.cache/ms-playwright
ENV PYPPETEER_HOME=/root/.cache/ms-playwright

# --------------------------------------------------------
# 🚀 Comando de inicio Streamlit
# --------------------------------------------------------
EXPOSE 8501
CMD ["streamlit", "run", "aplicacion.py", "--server.port=8501", "--server.address=0.0.0.0"]
