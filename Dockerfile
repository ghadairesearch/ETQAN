FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1
ENV UPLOAD_FOLDER=/var/data/uploads
ENV ORG_LOGO_FOLDER=/var/data/uploads/organization_logos

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential \
        fonts-dejavu-core \
        fonts-noto-core \
        fonts-noto-unhinted \
        fonts-noto-color-emoji \
        poppler-utils \
        tesseract-ocr \
        tesseract-ocr-ara \
        tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /var/data/uploads /var/data/uploads/organization_logos

CMD ["sh", "-c", "gunicorn course_report:app --bind 0.0.0.0:${PORT:-10000} --workers 2 --threads 4 --timeout 180"]
