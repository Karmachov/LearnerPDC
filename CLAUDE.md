# LearnerPDC

Generates faculty-signed PDF reports that categorise students as slow or advanced learners from mid-term Excel data. FARM stack: FastAPI + React + MongoDB, with Celery/Redis for async report generation.

## Two codebases

`farm/` is the active stack and the only place to make changes. `WebApp/` is the original Flask monolith, kept as a read-only reference — never edit it. Behaviour that exists only in `WebApp/` is a spec to port, not code to modify.

## Commands

Run from `farm/`:

```bash
docker compose up --build -d          # start everything
docker compose logs -f worker         # tail one service
docker compose up --build backend -d  # rebuild one service after changes
docker compose up --scale worker=3 -d # parallel report generation
docker compose down                   # stop; ./mongo_data and ./shared_data survive
```

Frontend: `npm run build` inside `farm/frontend/` — run this and confirm it's clean before considering any UI change finished.

First build takes ~10 minutes because of the LibreOffice layer, so rebuild a single service rather than the whole stack when possible. Never run `rm -rf mongo_data/ shared_data/` without asking first — it destroys all data.

Services: frontend `80` (Nginx), backend `8000`, MongoDB `27017`, Redis `6379`, worker has no exposed port.

`docker-compose.prod.yml` is a separate production compose file that pulls prebuilt `reebhuc/report-generator:v7-*` images instead of building locally (`docker compose -f docker-compose.prod.yml up -d`). Dev work uses `docker-compose.yml`; don't conflate the two.

## Invariants

Get these right; they are the usual source of broken changes here.

- Async/sync split. `backend/database.py` uses Motor and is async. `worker/database.py` uses sync PyMongo because Celery tasks are sync. Never `await` in worker code; never make blocking PyMongo calls inside a FastAPI route handler.
- Pydantic v2. Use `model_validate`, `model_dump`, `field_validator`, `ConfigDict`. Not `parse_obj`, `.dict()`, `validator`, or inner `class Config`.
- bcrypt is called directly, not via passlib. This is deliberate — it works around the passlib/bcrypt 4.1 compatibility bug. `auth.py` still imports `CryptContext` and keeps an unused `pwd_context` around as a dormant fallback for completeness — it is not wired into `hash_password`/`verify_password`. Don't wire it back in.
- PDF pipeline: report bodies are built with python-docx, converted to PDF via LibreOffice, then post-processed/overlaid with PyMuPDF; digital signing goes through `endesive`. Don't introduce ReportLab or WeasyPrint.
- Auth boundary. Every route requires a valid Bearer token except `/auth/token` and `/auth/register`. When adding a route, state which side of that line it falls on. TODO: `GET /faculty/{faculty_id}/photo` in `main.py` is currently unauthenticated — an undocumented exception to this rule, not yet reconciled.
- Target platform is Docker Desktop ≥ 4.x with Compose v2, on Apple Silicon or x86-64. Flag anything architecture-sensitive.

## Where changes go

```
farm/backend/main.py        All REST routes — new routes go here
farm/backend/auth.py        JWT + bcrypt
farm/backend/database.py    Motor (async) + Fernet vault
farm/backend/models.py      Pydantic v2 schemas
farm/backend/celery_app.py  Shared Celery factory
farm/worker/tasks.py        Task pipeline: decrypt → generate → sign
farm/worker/logic.py        ReportController — report generation and signing logic
farm/worker/database.py     Sync PyMongo + Fernet
farm/shared/vault.py         Fernet encryption implementation
farm/shared/mongo.py         Unified async/sync MongoDB accessors
farm/shared/audit.py         Audit event logging
farm/shared/pem_validation.py PEM key/certificate validation
farm/shared/task_errors.py   ReportTaskError and safe error shaping
farm/frontend/src/context/AuthContext.jsx   JWT state
farm/frontend/src/api/client.js             Axios instance, owns the token interceptor
farm/frontend/src/pages/                    Login, Dashboard, Profile
farm/frontend/src/components/               Navbar, TaskStatusCard
```

Frontend API calls go through `api/client.js` — don't attach auth headers manually in components.

A schema change often spans `backend/models.py`, `worker/logic.py`, `shared/`, and a frontend component at once. List every touch point before editing any of them.

Celery tasks should be idempotent where practical; report generation may be retried.

## Security

- `MASTER_KEY` (Fernet) and `JWT_SECRET` are required and have no defaults. Never write a real or realistic-looking value for either — use `<MASTER_KEY>` style placeholders. Optional: `MONGO_DB` (default `learner_pdc`), `ACCESS_TOKEN_EXPIRE_MINUTES` (default `60`), `ALLOWED_ORIGINS` (default localhost).
- Signing credentials — private keys, certificates, passphrases, signature images — are Fernet-encrypted at rest in MongoDB and decrypted in RAM at signing time only. Never write decrypted secret material to the worker filesystem, log it, or return it from an API response.
- Never stage or commit `.env`, `*.pem`, `*.key`, or anything containing real credentials. If a live-looking secret turns up in the working tree, stop and say so rather than committing around it.

## Docs

`README.md` and `SYSTEM_ARCHITECTURE.md` should stay consistent with each other. Match the existing README voice: tables for services and env vars, fenced blocks for commands, short declarative sentences. If a code change makes either doc wrong, say so.
