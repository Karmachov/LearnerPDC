"""Backward-compatible re-exports from the shared package."""

from shared.mongo import (  # noqa: F401
    close_async_client,
    get_faculty_collection,
    get_sync_db,
    get_sync_faculty_collection,
)
from shared.vault import decrypt_bytes, decrypt_str, encrypt_bytes, encrypt_str  # noqa: F401
