# LearnerPDC — Student Learner Report Generator

A production-grade **FARM stack** (FastAPI · React · MongoDB · Celery/Redis) application that generates faculty-signed PDF reports categorising students as slow or advanced learners from mid-term Excel data.

> The original Flask monolith lives in `WebApp/` and remains untouched as a reference. All active development is in `farm/`.

---

## Architecture

```
farm/
├── backend/        FastAPI — REST API, JWT auth, Fernet encryption vault
├── worker/         Celery worker — LibreOffice PDF conversion, digital signing
├── frontend/       React + Vite + Tailwind, served by Nginx in production
└── docker-compose.yml
```

| Service | Port | Technology |
|---|---|---|
| Frontend | `80` | React → Nginx |
| Backend API | `8000` | FastAPI + Motor (async MongoDB) |
| MongoDB | `27017` | Persistent data + encrypted secret vault |
| Redis | `6379` | Celery broker + result backend |
| Worker | — | Celery + LibreOffice + PyMuPDF |

---

## Quick Start (Docker — recommended)

### Prerequisites
- Docker Desktop ≥ 4.x with Compose v2
- Apple Silicon (M1/M2/M3) or x86-64

### 1 — Clone

```bash
git clone <repo-url>
cd LearnerPDC/farm
```

### 2 — Configure secrets

```bash
cp .env.example .env
```

Open `.env` and fill in **two required secrets**:

```bash
# Generate MASTER_KEY (Fernet symmetric key):
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"

# Generate JWT_SECRET (random 256-bit hex):
python -c "import secrets; print(secrets.token_hex(32))"
```

Paste the outputs into `.env`. All other variables have working defaults.

### 3 — Build and start

```bash
docker compose up --build -d
```

First build takes ~10 minutes (LibreOffice image is large). Subsequent builds are fast due to layer caching.

### 4 — Open the app

Navigate to **http://localhost** and register your faculty account.

---

## Developer Workflow

```bash
# Tail all service logs
docker compose logs -f

# Tail a single service
docker compose logs -f worker

# Rebuild a single service after code changes
docker compose up --build backend -d

# Scale workers for parallel report generation
docker compose up --scale worker=3 -d

# Stop all services (data is preserved in ./mongo_data and ./shared_data)
docker compose down

# Full reset including data volumes
docker compose down && rm -rf mongo_data/ shared_data/
```

---

## Environment Variables

All configuration is in `farm/.env` (git-ignored). See [`farm/.env.example`](farm/.env.example) for full documentation of every variable.

| Variable | Required | Description |
|---|---|---|
| `MASTER_KEY` | **Yes** | Fernet key for encrypting signing credentials at rest |
| `JWT_SECRET` | **Yes** | Secret for signing JWT access tokens |
| `MONGO_DB` | No | Database name (default: `learner_pdc`) |
| `ACCESS_TOKEN_EXPIRE_MINUTES` | No | Token TTL in minutes (default: `480`) |
| `ALLOWED_ORIGINS` | No | CORS allowed origins (default: localhost) |

---

## Security Model

- **Encryption at rest**: Private keys, certificates, passphrases, and signature images are Fernet-encrypted before being written to MongoDB. The `MASTER_KEY` is the only secret that unlocks them.
- **In-memory signing**: The worker decrypts credentials in RAM at signing time. No secret material is ever written to the worker filesystem.
- **JWT auth**: All API routes (except `/auth/token` and `/auth/register`) require a valid Bearer token.
- **bcrypt passwords**: Faculty passwords are hashed with bcrypt directly (bypassing the passlib/bcrypt 4.1 compatibility bug).

---

## Directory Layout

```
LearnerPDC/
├── farm/                   ← Active FARM stack (all new development here)
│   ├── backend/
│   │   ├── main.py         FastAPI routes
│   │   ├── auth.py         JWT + bcrypt
│   │   ├── database.py     Motor async client + Fernet vault
│   │   ├── models.py       Pydantic v2 schemas
│   │   └── celery_app.py   Shared Celery factory
│   ├── worker/
│   │   ├── tasks.py        Celery task: decrypt → generate → sign
│   │   ├── logic.py        ReportController (LibreOffice + PyMuPDF)
│   │   └── database.py     Sync PyMongo + Fernet (for Celery tasks)
│   ├── frontend/
│   │   └── src/
│   │       ├── pages/      Login.jsx · Dashboard.jsx · Profile.jsx
│   │       ├── components/ Navbar.jsx · TaskStatusCard.jsx
│   │       ├── context/    AuthContext.jsx (JWT state)
│   │       └── api/        client.js (Axios + interceptors)
│   ├── docker-compose.yml
│   └── .env.example
├── WebApp/                 ← Original Flask app (reference only, untouched)
└── README.md
```

---

## Contributing

1. Never commit `.env`, `*.pem`, `*.key`, or any file containing real credentials.
2. All new backend routes go in `farm/backend/main.py`.
3. Worker logic changes go in `farm/worker/logic.py` (the `ReportController` class).
4. Run `npm run build` inside `farm/frontend/` before pushing UI changes to verify the build is clean.
