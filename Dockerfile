# Dockerfile
FROM python:3.11-slim

WORKDIR /app

# Install dependensi sistem (untuk pandas/openpyxl)
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements dulu (cache layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY main.py .
COPY setup_db.py .

# Folder data akan di-mount dari host (lihat docker-compose.yml)
RUN mkdir -p data

CMD ["python", "main.py"]