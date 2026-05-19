"""Backward-compatible re-exports from the shared package."""

from shared.mongo import get_sync_db, get_sync_faculty_collection  # noqa: F401
from shared.vault import decrypt_bytes, decrypt_str, encrypt_bytes, encrypt_str  # noqa: F401
