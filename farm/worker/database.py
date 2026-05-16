"""
database.py — Sync-only copy for the worker service.
The worker uses PyMongo (sync) because Celery tasks don't have an asyncio event loop.
This module shares the same Fernet vault logic as the backend's database.py.
"""

import os

from cryptography.fernet import Fernet, InvalidToken
from pymongo import MongoClient

MONGO_URL: str = os.environ.get("MONGO_URL", "mongodb://localhost:27017")
DB_NAME: str = os.environ.get("MONGO_DB", "learner_pdc")

_raw_key = os.environ.get("MASTER_KEY", "")
if not _raw_key:
    raise RuntimeError("MASTER_KEY environment variable is not set.")

_fernet = Fernet(_raw_key.encode())


def encrypt_bytes(data: bytes) -> bytes:
    return _fernet.encrypt(data)


def decrypt_bytes(token: bytes) -> bytes:
    try:
        return _fernet.decrypt(token)
    except InvalidToken as exc:
        raise ValueError("Decryption failed — MASTER_KEY may have changed.") from exc


def decrypt_str(token: bytes) -> str:
    return decrypt_bytes(token).decode("utf-8")


def encrypt_str(plaintext: str) -> bytes:
    return encrypt_bytes(plaintext.encode("utf-8"))


_sync_client: MongoClient | None = None


def get_sync_db():
    global _sync_client
    if _sync_client is None:
        _sync_client = MongoClient(MONGO_URL)
    return _sync_client[DB_NAME]


def get_sync_faculty_collection():
    return get_sync_db()["faculty"]
