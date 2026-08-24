FROM python:3.11-slim

# System deps — tesseract for OCR, libGL for pymupdf
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
    libgl1 \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Railway sets PORT env var; default 5050 for local
ENV PORT=5050

EXPOSE $PORT

# gunicorn: 2 workers, 5-min timeout for large file processing
CMD gunicorn --bind 0.0.0.0:$PORT --timeout 600 --workers 1 --max-requests 50 --graceful-timeout 60 app:app
