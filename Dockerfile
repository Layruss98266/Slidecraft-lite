FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=10000 \
    HOST=0.0.0.0

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-core \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt gunicorn

COPY . .

RUN mkdir -p uploads exports history static/slides static/originals

EXPOSE 10000

CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 300
