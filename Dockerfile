FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8000

WORKDIR /app

# Sistem bağımlılıkları ve C derleyicisi
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Bağımlılıkları yükle
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Proje dosyalarını kopyala
COPY . .

EXPOSE 8000

# app.py veya main.py hangisi varsa otomatik algılayıp uvicorn başlatır
CMD ["sh", "-c", "if [ -f app.py ]; then uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1; else uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1; fi"]
