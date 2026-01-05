FROM python:3.9-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# STEP 1: Heavy Install (LibreOffice)
# We keep this EXACTLY like your first successful build to try and hit the cache.
RUN apt-get update && apt-get install -y \
    build-essential \
    libffi-dev \
    libssl-dev \
    libreoffice-writer \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

# STEP 2: Font Install (Separate Layer)
# This way, if this fails or is slow, we don't lose the LibreOffice download.
RUN sed -i -e's/ main/ main contrib non-free/g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || \
    sed -i -e's/ main/ main contrib non-free/g' /etc/apt/sources.list && \
    apt-get update && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y \
    ttf-mscorefonts-installer \
    fontconfig \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads Learner_Monitor_Reports

EXPOSE 5001

CMD ["python", "app.py"]