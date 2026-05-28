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
ENV FINALIZE_FIELDS=true
ENV DOCX_CONVERSION_TIMEOUT=300
ENV FIELD_FINALIZATION_TIMEOUT=180
ENV DOCX_FINALIZER_PYTHON=/usr/bin/python3
ENV SOFFICE_PATH=/usr/bin/soffice
ENV MALLOC_ARENA_MAX=1
ENV PYTHONDONTWRITEBYTECODE=1
ENV OMP_NUM_THREADS=1
ENV OPENBLAS_NUM_THREADS=1
ENV SAL_USE_VCLPLUGIN=svp
EXPOSE 8000
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
