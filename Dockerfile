FROM python:3.13-slim

# Системные библиотеки для WeasyPrint (Pango/Cairo/GObject) — без них PDF-экспорт падает.
# ВАЖНО: python:3.13-slim сейчас на Debian 13 (trixie), где пакет gdk-pixbuf переименован:
# libgdk-pixbuf2.0-0 -> libgdk-pixbuf-2.0-0 (лишний дефис). Если когда-нибудь база образа
# сменится обратно на Debian 12 (bookworm) — вернуть старое имя без дефиса.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libpangoft2-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libgobject-2.0-0 \
    libglib2.0-0 \
    libffi-dev \
    shared-mime-info \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120"]
