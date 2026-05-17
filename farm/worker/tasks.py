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
import traceback
from pathlib import Path

# ── Guarantee /app is on sys.path before any import ─────────────────────────
# WORKDIR in the Dockerfile is /app; all worker modules (database.py, logic.py)
# live there. We insert it explicitly because:
#   - Celery prefork children may run with a different cwd than the parent
#   - 'python -c' and subprocesses do not inherit the parent's cwd reliably
#   - Path(__file__).parent works at module level but may not resolve correctly
#     inside forked child processes on some container runtimes
# Using the hardcoded value is safe — it is the canonical WORKDIR and never changes.
_WORKER_DIR = "/app"
if _WORKER_DIR not in sys.path:
    sys.path.insert(0, _WORKER_DIR)

# ── Module-level imports (resolved once in main process, inherited by forks) ─
# Importing here rather than inside the task function body ensures:
#   1. Import errors surface immediately at worker startup (not mid-task)
#   2. Forked child processes inherit the already-resolved module objects
#   3. No risk of cwd-relative lookup failing in a child's changed working dir
from database import get_sync_faculty_collection, decrypt_bytes, decrypt_str  # noqa: E402
from logic import ReportController  # noqa: E402

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
        "password": decrypt_str(doc["key_password_enc"]) if doc.get("key_password_enc") else "",
    }
    return secrets


# ---------------------------------------------------------------------------
# Main Celery task
# ---------------------------------------------------------------------------

@celery_app.task(
    bind=True,
    name="generate_report_task",
    max_retries=2,
    soft_time_limit=300,  # 5 min soft limit → raises SoftTimeLimitExceeded
    time_limit=360,        # 6 min hard kill
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

    try:
        # --- Stage 1: Fetch secrets ---
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
            # Still try to fetch image for the visual signature line
            try:
                secrets = _get_faculty_secrets(faculty_id)
                sign_info["image_bytes"] = secrets.get("image_bytes")
            except Exception:
                sign_info["image_bytes"] = None

        # --- Stage 2: Generate report ---
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

        # --- Stage 3: Cleanup upload temp files ---
        self.update_state(
            state=states.STARTED,
            meta={"message": "Finalising and cleaning up…", "progress": 90},
        )

        upload_task_dir = Path(payload["excel_path"]).parent
        try:
            shutil.rmtree(str(upload_task_dir), ignore_errors=True)
        except Exception as e:
            logger.warning(f"[{task_id}] Could not clean upload dir: {e}")

        logger.info(f"[{task_id}] Report generation complete: {output_path}")

        return {
            "status": "SUCCESS",
            "download_token": task_id,  # the API uses task_id as the download key
            "output_path": output_path if isinstance(output_path, str) else str(output_path),
        }

    except Exception as exc:
        logger.error(f"[{task_id}] Task failed: {exc}")
        traceback.print_exc()
        # Do not retry for user-input errors (bad Excel, wrong filters, etc.)
        self.update_state(
            state=states.FAILURE,
            meta={"exc_type": type(exc).__name__, "exc_message": str(exc)},
        )
        raise exc
