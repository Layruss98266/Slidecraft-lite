FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=10000 \
    HOST=0.0.0.0 \
    WEB_CONCURRENCY=1 \
    GUNICORN_THREADS=2

# Minimal LibreOffice — only Impress + core. No Java, no help, no l10n.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-core \
        fonts-dejavu-core \
    && apt-get purge -y libreoffice-help-common libreoffice-help-en-us 2>/dev/null || true \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/* /var/cache/apt/archives/* /usr/share/doc/* /usr/share/man/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt gunicorn

COPY . .

RUN mkdir -p sessions

EXPOSE 10000

# Single worker to fit in 512 MB. Preload to share libs across threads.
CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --workers ${WEB_CONCURRENCY:-1} \
    --threads ${GUNICORN_THREADS:-2} \
    --timeout 300 \
    --preload \
    --max-requests 50 \
    --max-requests-jitter 10
