"""
tasks.py — Celery task definitions for the worker service.

The worker fetches encrypted secrets from MongoDB using a synchronous PyMongo
client (Motor async is not usable inside a Celery task without an event loop),
decrypts them in-memory, and passes raw bytes to the ReportController.

No secret material ever touches the filesystem.
"""

import os
import sys
import shutil
import time
import traceback
from pathlib import Path

# ── Guarantee /app is on sys.path before any import ─────────────────────────
_WORKER_DIR = "/app"
if _WORKER_DIR not in sys.path:
    sys.path.insert(0, _WORKER_DIR)

from database import get_sync_faculty_collection, decrypt_bytes, decrypt_str  # noqa: E402
from logic import ReportController  # noqa: E402
from shared.audit import audit_log_sync  # noqa: E402
from shared.pem_validation import normalize_passphrase  # noqa: E402
from shared.task_errors import ReportTaskError, user_facing_task_error  # noqa: E402

from celery import Celery, states
from celery.utils.log import get_task_logger

# ---------------------------------------------------------------------------
# Celery app (mirrors backend/celery_app.py)
# ---------------------------------------------------------------------------
REDIS_URL = os.environ.get("REDIS_URL", "redis://localhost:6379/0")

celery_app = Celery(
    "learner_pdc",
    broker=REDIS_URL,
    backend=REDIS_URL,
)

celery_app.conf.update(
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],
    timezone="Asia/Kolkata",
    enable_utc=True,
    result_expires=86400,
    worker_prefetch_multiplier=1,
    task_acks_late=True,
)

logger = get_task_logger(__name__)


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _get_faculty_secrets(faculty_id: str) -> dict:
    """
    Fetch the faculty document from MongoDB synchronously and decrypt secrets.
    Returns a dict with plaintext bytes for key, cert, image, and password.
    """
    from bson import ObjectId

    col = get_sync_faculty_collection()
    doc = col.find_one({"_id": ObjectId(faculty_id)})
    if not doc:
        raise ValueError(f"Faculty not found: {faculty_id}")

    secrets = {
        "key_bytes": decrypt_bytes(doc["private_key_enc"]) if doc.get("private_key_enc") else None,
        "cert_bytes": decrypt_bytes(doc["certificate_enc"]) if doc.get("certificate_enc") else None,
        "image_bytes": decrypt_bytes(doc["signature_image_enc"]) if doc.get("signature_image_enc") else None,
        "password": normalize_passphrase(
            decrypt_str(doc["key_password_enc"]) if doc.get("key_password_enc") else ""
        ),
    }
    return secrets


def _excel_file_size(payload: dict) -> int | None:
    try:
        return Path(payload["excel_path"]).stat().st_size
    except OSError:
        return None


# ---------------------------------------------------------------------------
# Main Celery task
# ---------------------------------------------------------------------------

@celery_app.task(
    bind=True,
    name="generate_report_task",
    max_retries=2,
    soft_time_limit=300,
    time_limit=360,
)
def generate_report_task(self, payload: dict):
    """
    Background task that:
      1. Fetches and decrypts faculty secrets from MongoDB
      2. Runs the ReportController to generate DOCX/PDF
      3. Returns the output file path(s) on the shared volume
    """
    task_id: str = payload["task_id"]
    faculty_id: str = payload["faculty_id"]
    started_at = time.monotonic()
    file_size = _excel_file_size(payload)

    try:
        self.update_state(
            state=states.STARTED,
            meta={"message": "Fetching credentials from vault…", "progress": 5},
        )

        sign_info: dict = {"should_sign": False}

        if payload.get("enable_signing"):
            logger.info(f"[{task_id}] Decrypting signing credentials for faculty {faculty_id}")
            secrets = _get_faculty_secrets(faculty_id)
            sign_info = {
                "should_sign": True,
                "key_bytes": secrets["key_bytes"],
                "cert_bytes": secrets["cert_bytes"],
                "image_bytes": secrets["image_bytes"],
                "password": secrets["password"],
            }
        else:
            try:
                secrets = _get_faculty_secrets(faculty_id)
                sign_info["image_bytes"] = secrets.get("image_bytes")
            except Exception:
                sign_info["image_bytes"] = None

        self.update_state(
            state=states.STARTED,
            meta={"message": "Generating report document…", "progress": 25},
        )

        output_dir = payload["output_dir"]
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        controller = ReportController(
            excel_path=payload["excel_path"],
            cgpa_path=payload.get("cgpa_path"),
            grade_path=payload.get("grade_path"),
            format_choice=payload["format_choice"],
            learner_type=payload["learner_type"],
            slow_thresh=payload["slow_threshold"],
            advanced_thresh=payload["advanced_threshold"],
            output_type=payload["output_type"],
            semester=payload["semester"],
            sign_info=sign_info,
            common_comment=payload.get("common_comment", ""),
            faculty_name=payload.get("faculty_name", ""),
            output_dir=output_dir,
        )

        self.update_state(
            state=states.STARTED,
            meta={"message": "Processing Excel data and building document…", "progress": 50},
        )

        output_path = controller.run()

        if not output_path:
            raise RuntimeError("ReportController returned no output path. Check the Excel file and filters.")

        self.update_state(
            state=states.STARTED,
            meta={"message": "Finalising and cleaning up…", "progress": 90},
        )

        upload_task_dir = Path(payload["excel_path"]).parent
        try:
            shutil.rmtree(str(upload_task_dir), ignore_errors=True)
        except Exception as e:
            logger.warning(f"[{task_id}] Could not clean upload dir: {e}")

        duration_ms = int((time.monotonic() - started_at) * 1000)
        logger.info(f"[{task_id}] Report generation complete: {output_path}")

        audit_log_sync(
            "report.completed",
            faculty_id=faculty_id,
            task_id=task_id,
            duration_ms=duration_ms,
            file_size_bytes=file_size,
            details={"output_type": payload.get("output_type")},
        )

        return {
            "status": "SUCCESS",
            "download_token": task_id,
            "output_path": output_path if isinstance(output_path, str) else str(output_path),
        }

    except Exception as exc:
        duration_ms = int((time.monotonic() - started_at) * 1000)
        safe_message = user_facing_task_error(exc)
        logger.error(f"[{task_id}] Task failed: {exc}")
        traceback.print_exc()

        audit_log_sync(
            "report.failed",
            faculty_id=faculty_id,
            task_id=task_id,
            duration_ms=duration_ms,
            file_size_bytes=file_size,
            error_summary=safe_message,
            details={"exc_type": type(exc).__name__},
        )

        # Raise only ReportTaskError so Celery's FAILURE metadata stays user-safe.
        # Do not call update_state(FAILURE) — Celery overwrites it with the full traceback.
        raise ReportTaskError(safe_message) from exc
