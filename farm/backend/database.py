"""
database.py — Motor async client and Fernet encryption/decryption vault.

All sensitive blobs (private key PEM, certificate PEM, signature image,
key passphrase) are encrypted with Fernet before being stored in MongoDB
and decrypted in-memory only when needed by the worker.

MASTER_KEY must be a URL-safe base64-encoded 32-byte key.
Generate one with: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
"""

import os
import base64

from motor.motor_asyncio import AsyncIOMotorClient
from cryptography.fernet import Fernet, InvalidToken
from pymongo import MongoClient  # sync client for Celery worker context

# ---------------------------------------------------------------------------
# Environment
# ---------------------------------------------------------------------------
MONGO_URL: str = os.environ.get("MONGO_URL", "mongodb://localhost:27017")
DB_NAME: str = os.environ.get("MONGO_DB", "learner_pdc")

_raw_key = os.environ.get("MASTER_KEY", "")
if not _raw_key:
    raise RuntimeError(
        "MASTER_KEY environment variable is not set. "
        "Generate one with: python -c \"from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())\""
    )

try:
    _fernet = Fernet(_raw_key.encode())
except Exception as exc:
    raise RuntimeError(f"MASTER_KEY is invalid: {exc}") from exc

# ---------------------------------------------------------------------------
# Encryption helpers (used by both API and worker)
# ---------------------------------------------------------------------------

def encrypt_bytes(data: bytes) -> bytes:
    """Encrypt arbitrary bytes with Fernet. Safe to store in MongoDB."""
    return _fernet.encrypt(data)


def decrypt_bytes(token: bytes) -> bytes:
    """Decrypt a Fernet token back to its original bytes."""
    try:
        return _fernet.decrypt(token)
    except InvalidToken as exc:
        raise ValueError("Decryption failed — token is invalid or MASTER_KEY has changed.") from exc


def encrypt_str(plaintext: str) -> bytes:
    """Convenience wrapper: encrypt a UTF-8 string."""
    return encrypt_bytes(plaintext.encode("utf-8"))


def decrypt_str(token: bytes) -> str:
    """Convenience wrapper: decrypt a Fernet token to a UTF-8 string."""
    return decrypt_bytes(token).decode("utf-8")


# ---------------------------------------------------------------------------
# Async Motor client (used by FastAPI)
# ---------------------------------------------------------------------------
_async_client: AsyncIOMotorClient | None = None


def get_async_client() -> AsyncIOMotorClient:
    global _async_client
    if _async_client is None:
        _async_client = AsyncIOMotorClient(MONGO_URL)
    return _async_client


def get_database():
    """Return the Motor async database instance."""
    return get_async_client()[DB_NAME]


def get_faculty_collection():
    return get_database()["faculty"]


def get_task_meta_collection():
    """Optional: store task metadata for richer status queries."""
    return get_database()["tasks"]


async def close_async_client():
    global _async_client
    if _async_client is not None:
        _async_client.close()
        _async_client = None


# ---------------------------------------------------------------------------
# Sync PyMongo client (used by Celery worker — no event loop available)
# ---------------------------------------------------------------------------
_sync_client: MongoClient | None = None


def get_sync_db():
    global _sync_client
    if _sync_client is None:
        _sync_client = MongoClient(MONGO_URL)
    return _sync_client[DB_NAME]


def get_sync_faculty_collection():
    return get_sync_db()["faculty"]
