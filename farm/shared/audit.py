"""
Operational audit log — persisted to MongoDB ``audit_logs`` collection.

Writes are best-effort: failures are logged and never propagate to callers.
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any, Optional

from shared.mongo import get_audit_logs_collection, get_sync_audit_logs_collection

logger = logging.getLogger(__name__)


def _base_doc(
    event_type: str,
    *,
    faculty_id: Optional[str] = None,
    task_id: Optional[str] = None,
    duration_ms: Optional[int] = None,
    file_size_bytes: Optional[int] = None,
    error_summary: Optional[str] = None,
    details: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    doc: dict[str, Any] = {
        "event_type": event_type,
        "created_at": datetime.now(timezone.utc),
    }
    if faculty_id is not None:
        doc["faculty_id"] = faculty_id
    if task_id is not None:
        doc["task_id"] = task_id
    if duration_ms is not None:
        doc["duration_ms"] = duration_ms
    if file_size_bytes is not None:
        doc["file_size_bytes"] = file_size_bytes
    if error_summary is not None:
        doc["error_summary"] = error_summary[:500]
    if details:
        doc["details"] = details
    return doc


async def audit_log_async(
    event_type: str,
    *,
    faculty_id: Optional[str] = None,
    task_id: Optional[str] = None,
    duration_ms: Optional[int] = None,
    file_size_bytes: Optional[int] = None,
    error_summary: Optional[str] = None,
    details: Optional[dict[str, Any]] = None,
) -> bool:
    try:
        col = get_audit_logs_collection()
        await col.insert_one(
            _base_doc(
                event_type,
                faculty_id=faculty_id,
                task_id=task_id,
                duration_ms=duration_ms,
                file_size_bytes=file_size_bytes,
                error_summary=error_summary,
                details=details,
            )
        )
        return True
    except Exception as exc:
        logger.warning("Audit log failed [%s]: %s", event_type, exc)
        return False


def audit_log_sync(
    event_type: str,
    *,
    faculty_id: Optional[str] = None,
    task_id: Optional[str] = None,
    duration_ms: Optional[int] = None,
    file_size_bytes: Optional[int] = None,
    error_summary: Optional[str] = None,
    details: Optional[dict[str, Any]] = None,
) -> bool:
    try:
        col = get_sync_audit_logs_collection()
        col.insert_one(
            _base_doc(
                event_type,
                faculty_id=faculty_id,
                task_id=task_id,
                duration_ms=duration_ms,
                file_size_bytes=file_size_bytes,
                error_summary=error_summary,
                details=details,
            )
        )
        return True
    except Exception as exc:
        logger.warning("Audit log failed [%s]: %s", event_type, exc)
        return False
