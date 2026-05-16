FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# ─── LAYER 1: Minimal compile-and-runtime suite ──────────────────────────────
# - libssl3, libffi8 : pre-built runtime binaries for cryptography/cffi (.so)
# - g++              : strictly required to build pykcs11 (endesive's PKCS#11
#                      C++ wrapper). Installed in isolation — NOT via the
#                      build-essential meta-package, which would pull in
#                      make/binutils/libc-dev and over-allocate temporary
#                      installation buffers, causing OOM on ARM64.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libssl3 \
    libffi8 \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# ─── LAYER 2: LibreOffice + Java (heaviest layer — keep near top for cache) ──
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

# ─── LAYER 3: MS Core Fonts (slow download — isolated layer) ─────────────────
RUN sed -i -e's/ main/ main contrib non-free/g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || \
    sed -i -e's/ main/ main contrib non-free/g' /etc/apt/sources.list && \
    apt-get update && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y ttf-mscorefonts-installer fontconfig && \
    fc-cache -f -v && \
    rm -rf /var/lib/apt/lists/*

# ─── LAYER 4: Python dependencies ────────────────────────────────────────────
# --prefer-binary forces pip to use pre-compiled manylinux/musllinux wheels for
# packages like cryptography, pandas, and PyMuPDF instead of source tarballs,
# preventing C extension compilation and associated memory spikes.
COPY requirements.txt .
RUN pip install --upgrade pip --no-cache-dir && \
    pip install --no-cache-dir --prefer-binary -r requirements.txt

# ─── LAYER 5: Application code (changes most often — last layer) ─────────────
COPY . .

RUN mkdir -p /data/shared/uploads /data/shared/reports

CMD ["celery", "-A", "tasks", "worker", "--loglevel=info", "--concurrency=2"]
