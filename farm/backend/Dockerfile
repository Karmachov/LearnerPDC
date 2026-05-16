FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# ─── Runtime binary deps only — no source compilation headers ─────────────────
# libssl3 and libffi8 satisfy cryptography/cffi at runtime without triggering
# gcc/make compilation that causes OOM on constrained ARM64 environments.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libssl3 \
    libffi8 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
# --prefer-binary forces pip to choose pre-compiled manylinux/musllinux wheels
# for packages like cryptography, bcrypt, and pydantic-core instead of
# building from source tarballs.
RUN pip install --upgrade pip --no-cache-dir && \
    pip install --no-cache-dir --prefer-binary -r requirements.txt

COPY . .

# Shared volume mount point
RUN mkdir -p /data/shared/uploads /data/shared/reports

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
