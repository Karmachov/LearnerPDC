# LearnerPDC System Architecture

This document details the technical architecture, cryptographic implementation, and asynchronous data pipelines of the LearnerPDC system.

---

## 1. COMPONENT & REPOSITORY LANDSCAPE

### Directory Structure & Responsibilities

The project is organized into a modular **FARM stack** architecture (FastAPI, React, MongoDB, Celery), strictly separating synchronous API handling from heavy compute tasks.

```text
farm/
├── backend/            # FastAPI REST API (Synchronous Boundary)
│   ├── main.py         # Entry point, route definitions, and multipart handling
│   ├── auth.py         # JWT issuance/validation and bcrypt password hashing
│   ├── database.py     # Async Motor client and Fernet vault re-exports
│   ├── models.py       # Pydantic v2 schemas for request validation and response shapes
│   └── celery_app.py   # Celery client factory for task dispatching
├── worker/             # Celery Worker (Asynchronous Compute)
│   ├── tasks.py        # Task definitions, state updates, and secret decryption
│   ├── logic.py        # ReportController: LibreOffice conversion & digital signing
│   └── database.py     # Sync PyMongo client for blocking database operations
├── shared/             # Common Logic (Cross-Service Library)
│   ├── vault.py        # Fernet symmetric encryption implementation
│   ├── mongo.py        # Unified MongoDB accessors (Async/Sync clients)
│   ├── audit.py        # Centralized event logging for security/monitoring
│   └── pem_validation.py # Cryptographic validation of PEM keys and certificates
└── frontend/           # React + Vite (Client Layer)
    └── src/api/client.js # Axios instance with JWT interceptors
```

### Container Orchestration Mesh

The system utilizes `farm/docker-compose.yml` to define a four-service mesh isolated within a shared bridge network.

| Service | Environment | Responsibility | Isolation |
| :--- | :--- | :--- | :--- |
| `learner_frontend` | Alpine/Nginx | Serves static React assets | Public-facing port 80 |
| `learner_backend` | Python/FastAPI | Request validation, auth, and DB orchestration | Internal compute, port 8000 |
| `learner_worker` | Python/Celery | Heavy PDF conversion and signing | No public ports, scaleable |
| `learner_mongodb` | Mongo 7 | Encrypted document storage | Persistent storage |
| `learner_redis` | Redis 7 | Message broker for Celery | In-memory state |

---

## 2. CRYPTOGRAPHIC DATAFLOW & PERSISTENCE MATRIX

### Lifecycle: `PUT /profile/keys` Asset Compilation

This route handles the most sensitive data in the system: faculty private keys and certificates.

1.  **Ingestion**: `backend/main.py` receives a multipart form containing raw PEM bytes and a passphrase string.
2.  **Validation**: Before persistence, `shared/pem_validation.py` attempts to load the private key in-memory using the provided passphrase. It validates that the public key matches the certificate.
3.  **Encryption (Fernet)**:
    -   The `MASTER_KEY` (AES-128 in CBC mode) is pulled from environment variables.
    -   Raw PEM bytes are encrypted via `shared/vault.py`.
    -   Passphrase strings are encoded to UTF-8 before encryption.
4.  **Persistence**: The resulting encrypted `bytes` are stored in the `faculty` collection. The `_id` field (BSON ObjectId) ensures precise document mapping.

### Data Transformation Matrix

| State | Format | Location | Layer |
| :--- | :--- | :--- | :--- |
| **In-Transit** | Raw PEM (Binary) | Multipart Form | Frontend → Backend |
| **Volatile** | Plaintext Bytes | Memory (RAM) | Backend Validation |
| **At-Rest** | Fernet Encrypted (Base64) | BSON `binData` | MongoDB |
| **Decrypted** | Plaintext Bytes | Memory (RAM) | Celery Worker (Signing) |

---

## 3. EVENT-DRIVEN ASYNCHRONOUS PIPELINE (CELERY + REDIS)

### Report Generation Trace

1.  **Trigger**: User submits form; `backend/main.py` generates a UUID `task_id`.
2.  **Staging**: Uploaded Excel files are written to the host-relative `./shared_data/uploads/{task_id}/` directory.
3.  **Dispatch**: The backend sends a JSON payload to Redis. The request returns `202 Accepted` immediately.
4.  **Consumption**: `worker/tasks.py` picks up the job, fetches encrypted secrets from MongoDB, and initializes the `ReportController`.

### Celery State Machine

The task transitions through the following states, visible via `GET /task-status/{task_id}`:

- `PENDING`: Task is in Redis queue, waiting for a worker.
- `STARTED`: Worker has acknowledged the task; metadata shows progress (e.g., "Generating report...").
- `SUCCESS`: PDF is written to `./shared_data/reports/{task_id}/`.
- `FAILURE`: Error caught; `shared/task_errors.py` sanitizes the message for the user.

### Shared Data File-Swapping

The system avoids passing large binary files through Redis. Instead, it utilizes root-relative bind mounts:

```yaml
# docker-compose.yml excerpt
volumes:
  - ./shared_data:/data/shared
```

- **Write (Backend)**: Input datasets are written to `/data/shared/uploads/`.
- **Read/Write (Worker)**: Worker reads inputs, runs LibreOffice, and writes output to `/data/shared/reports/`.
- **Read (Backend)**: `GET /download/{task_id}` serves the file directly from the reports directory.

---

## 4. LOCAL DATA ISOLATION MODEL

### Bind Mounts vs. Named Volumes

LearnerPDC explicitly uses **Bind Mounts** (`./mongo_data` and `./shared_data`) rather than Docker-internal named volumes.

| Feature | Bind Mounts (LearnerPDC) | Named Volumes (Docker Internal) |
| :--- | :--- | :--- |
| **Host Visibility** | Directly accessible in project root | Obscured in `/var/lib/docker/` |
| **Backup/Restore** | Standard `cp` or `tar` on the host | Requires `docker run --volumes-from` |
| **Persistence** | Data persists even if containers/volumes are purged | Data lost if `docker volume prune` is run |
| **Performance** | Native host filesystem speed | Slight overhead on non-Linux hosts |

This model ensures that generated reports and database files are easily auditable and manageable by the developer without specialized Docker tooling.
