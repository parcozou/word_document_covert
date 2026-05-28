FROM python:3.11-slim

WORKDIR /app
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer python3-uno \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p generated_files

ENV PORT=8000
ENV FINALIZE_FIELDS=false
ENV DOCX_FINALIZER_PYTHON=/usr/bin/python3
ENV SOFFICE_PATH=/usr/bin/soffice
EXPOSE 8000
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
