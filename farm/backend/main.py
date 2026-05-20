"""
main.py — FastAPI application: all routes, lifespan, CORS.
"""

from __future__ import annotations

import os
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Annotated

from bson import ObjectId
from fastapi import (
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    UploadFile,
    status,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from fastapi.security import OAuth2PasswordRequestForm

from auth import (
    create_access_token,
    get_current_faculty,
    hash_password,
    verify_password,
)
from database import (
    close_async_client,
    decrypt_bytes,
    decrypt_str,
    encrypt_bytes,
    encrypt_str,
    get_faculty_collection,
)
from shared.audit import audit_log_async
from shared.pem_validation import normalize_passphrase, validate_pem_key_pair
from models import (
    FacultyCreate,
    FacultyProfile,
    GenerateReportResponse,
    ProfileUpdateResponse,
    ReportHistoryItem,
    ReportHistoryResponse,
    ReportRequest,
    TaskStatusResponse,
    Token,
)

# ---------------------------------------------------------------------------
# Shared volume — mounted at /data/shared in both API and worker containers
# ---------------------------------------------------------------------------
SHARED_DIR = Path(os.environ.get("SHARED_DIR", "/data/shared"))
UPLOADS_DIR = SHARED_DIR / "uploads"
REPORTS_DIR = SHARED_DIR / "reports"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXCEL_EXTENSIONS = {".xls", ".xlsx"}

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(
    title="LearnerPDC API",
    description="Student Learner Report Generator — FARM Stack Backend",
    version="2.0.0",
)

# CORS — React dev server on 3000, production Nginx on 80
_origins = os.environ.get(
    "ALLOWED_ORIGINS",
    "http://localhost:3000,http://localhost,http://127.0.0.1"
).split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in _origins],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup():
    """Ensure MongoDB indexes exist."""
    col = get_faculty_collection()
    await col.create_index("email", unique=True)
    from shared.mongo import get_audit_logs_collection, get_tasks_collection

    audit_col = get_audit_logs_collection()
    await audit_col.create_index("created_at")
    await audit_col.create_index([("faculty_id", 1), ("created_at", -1)])
    await audit_col.create_index("task_id", sparse=True)

    tasks_col = get_tasks_collection()
    await tasks_col.create_index("task_id", unique=True)
    await tasks_col.create_index([("faculty_id", 1), ("created_at", -1)])


@app.on_event("shutdown")
async def shutdown():
    await close_async_client()


# ===========================================================================
# Auth Routes
# ===========================================================================

@app.post("/auth/register", response_model=FacultyProfile, status_code=status.HTTP_201_CREATED)
async def register(payload: FacultyCreate):
    col = get_faculty_collection()
    if await col.find_one({"email": payload.email}):
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="An account with this email already exists.",
        )
    doc = {
        "email": payload.email,
        "hashed_password": hash_password(payload.password),
        "name": payload.name,
        "role": payload.role,
        "department": payload.department,
        "signature_image_enc": None,
        "private_key_enc": None,
        "certificate_enc": None,
        "key_password_enc": None,
        "created_at": datetime.now(timezone.utc),
    }
    result = await col.insert_one(doc)
    created = await col.find_one({"_id": result.inserted_id})
    await audit_log_async(
        "auth.register",
        faculty_id=str(result.inserted_id),
        details={"email": payload.email},
    )
    return _faculty_to_profile(created)


@app.post("/auth/token", response_model=Token)
async def login(form: Annotated[OAuth2PasswordRequestForm, Depends()]):
    col = get_faculty_collection()
    faculty = await col.find_one({"email": form.username})
    if not faculty or not verify_password(form.password, faculty["hashed_password"]):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = create_access_token({"sub": str(faculty["_id"])})
    await audit_log_async("auth.login", faculty_id=str(faculty["_id"]))
    return Token(access_token=token)


@app.get("/auth/me", response_model=FacultyProfile)
async def me(faculty: dict = Depends(get_current_faculty)):
    return _faculty_to_profile(faculty)


# ===========================================================================
# Profile / Settings Routes
# ===========================================================================

@app.put("/profile/signature", response_model=ProfileUpdateResponse)
async def upload_signature(
    image: UploadFile = File(...),
    faculty: dict = Depends(get_current_faculty),
):
    """Encrypt and persist the faculty's visual signature image to MongoDB."""
    if image.content_type not in ("image/png", "image/jpeg", "image/jpg"):
        raise HTTPException(status_code=400, detail="Signature must be a PNG or JPEG image.")
    raw = await image.read()
    if len(raw) > 2 * 1024 * 1024:  # 2 MB cap
        raise HTTPException(status_code=400, detail="Image must be under 2 MB.")
    encrypted = encrypt_bytes(raw)
    col = get_faculty_collection()
    await col.update_one({"_id": faculty["_id"]}, {"$set": {"signature_image_enc": encrypted}})
    return ProfileUpdateResponse(message="Signature image saved securely.")


@app.put("/profile/photo", response_model=ProfileUpdateResponse)
async def upload_profile_photo(
    image: UploadFile = File(...),
    faculty: dict = Depends(get_current_faculty),
):
    """Save the faculty's profile photo to MongoDB."""
    if image.content_type not in ("image/png", "image/jpeg", "image/jpg", "image/webp"):
        raise HTTPException(status_code=400, detail="Photo must be a PNG, JPEG, or WEBP image.")
    raw = await image.read()
    if len(raw) > 2 * 1024 * 1024:  # 2 MB cap
        raise HTTPException(status_code=400, detail="Image must be under 2 MB.")
    col = get_faculty_collection()
    await col.update_one({"_id": faculty["_id"]}, {"$set": {"profile_photo": raw, "profile_photo_mime": image.content_type}})
    return ProfileUpdateResponse(message="Profile photo saved successfully.")


@app.get("/faculty/{faculty_id}/photo")
async def get_profile_photo(faculty_id: str):
    """Retrieve the faculty's profile photo."""
    try:
        obj_id = ObjectId(faculty_id)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid ID format")
    
    col = get_faculty_collection()
    doc = await col.find_one({"_id": obj_id}, {"profile_photo": 1, "profile_photo_mime": 1})
    if not doc or not doc.get("profile_photo"):
        raise HTTPException(status_code=404, detail="Photo not found")
        
    return Response(content=doc["profile_photo"], media_type=doc.get("profile_photo_mime", "image/jpeg"))


@app.put("/profile/keys", response_model=ProfileUpdateResponse)
async def upload_keys(
    private_key: UploadFile = File(...),
    certificate: UploadFile = File(...),
    key_password: str = Form(...),
    faculty: dict = Depends(get_current_faculty),
):
    """Encrypt and persist private key, certificate, and passphrase to MongoDB."""
    key_bytes = await private_key.read()
    cert_bytes = await certificate.read()

    passphrase = normalize_passphrase(key_password)
    try:
        validate_pem_key_pair(key_bytes, cert_bytes, passphrase)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    col = get_faculty_collection()
    await col.update_one(
        {"_id": faculty["_id"]},
        {
            "$set": {
                "private_key_enc": encrypt_bytes(key_bytes),
                "certificate_enc": encrypt_bytes(cert_bytes),
                "key_password_enc": encrypt_str(passphrase),
            }
        },
    )
    await audit_log_async(
        "profile.keys_uploaded",
        faculty_id=str(faculty["_id"]),
        file_size_bytes=len(key_bytes) + len(cert_bytes),
    )
    return ProfileUpdateResponse(message="Private key and certificate saved securely.")


# ===========================================================================
# Report Generation Routes
# ===========================================================================

@app.post("/generate-report", response_model=GenerateReportResponse, status_code=status.HTTP_202_ACCEPTED)
async def generate_report(
    excel_file: UploadFile = File(...),
    cgpa_file: UploadFile = File(None),
    grade_file: UploadFile = File(None),
    # Report parameters sent as Form fields alongside the file
    semester: str = Form(...),
    learner_type: str = Form(...),
    format_choice: str = Form(...),
    output_type: str = Form(...),
    slow_threshold: float = Form(40.0),
    advanced_threshold: float = Form(90.0),
    common_comment: str = Form(""),
    faculty_name: str = Form(""),
    enable_signing: bool = Form(False),
    proof_types: list[str] = Form([]),
    proof_files: list[UploadFile] = File(None),
    faculty: dict = Depends(get_current_faculty),
):
    # --- Validate inputs via Pydantic model ---
    try:
        req = ReportRequest(
            semester=semester,
            learner_type=learner_type,
            format_choice=format_choice,
            output_type=output_type,
            slow_threshold=slow_threshold,
            advanced_threshold=advanced_threshold,
            common_comment=common_comment,
            faculty_name=faculty_name,
            enable_signing=enable_signing,
            proof_types=proof_types,
        )
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    # --- Validate Excel file ---
    ext = Path(excel_file.filename).suffix.lower()
    if ext not in ALLOWED_EXCEL_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .xls and .xlsx files are accepted.")

    # --- Save uploaded files to shared volume ---
    task_id = str(uuid.uuid4())
    task_upload_dir = UPLOADS_DIR / task_id
    task_upload_dir.mkdir(parents=True)

    excel_path = str(task_upload_dir / f"main{ext}")
    excel_bytes = await excel_file.read()
    excel_size = len(excel_bytes)
    with open(excel_path, "wb") as f:
        f.write(excel_bytes)

    cgpa_path = None
    if cgpa_file and cgpa_file.filename:
        cgpa_ext = Path(cgpa_file.filename).suffix.lower()
        cgpa_path = str(task_upload_dir / f"cgpa{cgpa_ext}")
        with open(cgpa_path, "wb") as f:
            f.write(await cgpa_file.read())

    grade_path = None
    if grade_file and grade_file.filename:
        grade_ext = Path(grade_file.filename).suffix.lower()
        grade_path = str(task_upload_dir / f"grade{grade_ext}")
        with open(grade_path, "wb") as f:
            f.write(await grade_file.read())

    # --- Handle Proof Files ---
    task_report_dir = REPORTS_DIR / task_id
    if proof_files:
        valid_proofs = [pf for pf in proof_files if pf.filename]
        if len(valid_proofs) > 5:
            raise HTTPException(status_code=400, detail="Maximum of 5 proof files can be attached.")
        if len(valid_proofs) != len(req.proof_types):
            raise HTTPException(status_code=400, detail="Mismatched proof files and proof types.")
            
        if valid_proofs:
            task_report_dir.mkdir(parents=True, exist_ok=True)
            for idx, (pf, ptype) in enumerate(zip(valid_proofs, req.proof_types)):
                file_ext = Path(pf.filename).suffix.lower()
                if file_ext not in {".pdf", ".png", ".jpg", ".jpeg"}:
                    raise HTTPException(status_code=400, detail="Proof files must be PDF, PNG, or JPEG.")
                
                safe_type = ptype.replace(" ", "_").replace("/", "")
                proof_path = task_report_dir / f"proof_{safe_type}_{idx}{file_ext}"
                with open(proof_path, "wb") as f:
                    f.write(await pf.read())

    # --- Append Proof Type to Comment ---
    final_comment = req.common_comment
    proofs_text = ""
    if req.proof_types:
        unique_types = sorted(list(set(req.proof_types)))
        types_str = ", ".join(unique_types)
        proofs_text = f"Proofs: {types_str}"

    # --- Signing: verify keys exist and passphrase unlocks the key before queuing ---
    if req.enable_signing:
        if not faculty.get("private_key_enc") or not faculty.get("certificate_enc"):
            raise HTTPException(
                status_code=400,
                detail="Digital signing is enabled but no private key/certificate is stored. "
                       "Please upload them via Profile Settings first.",
            )
        if not faculty.get("key_password_enc"):
            raise HTTPException(
                status_code=400,
                detail="No key passphrase is stored. Re-upload your private key and certificate in Profile Settings.",
            )
        try:
            stored_key = decrypt_bytes(faculty["private_key_enc"])
            stored_cert = decrypt_bytes(faculty["certificate_enc"])
            stored_pass = normalize_passphrase(decrypt_str(faculty["key_password_enc"]))
            validate_pem_key_pair(stored_key, stored_cert, stored_pass)
        except ValueError as exc:
            raise HTTPException(
                status_code=400,
                detail=str(exc),
            ) from exc

    # --- Record task ownership (authorizes status + download) ---
    from shared.mongo import get_tasks_collection

    faculty_id = str(faculty["_id"])
    tasks_col = get_tasks_collection()
    await tasks_col.insert_one(
        {
            "task_id": task_id,
            "faculty_id": faculty_id,
            "status": "queued",
            "created_at": datetime.now(timezone.utc),
            "semester": req.semester,
            "learner_type": req.learner_type,
            "format_choice": req.format_choice,
            "output_type": req.output_type,
        }
    )

    # --- Dispatch Celery task ---
    from celery_app import celery_app  # import here to avoid startup side-effects

    task_payload = {
        "task_id": task_id,
        "faculty_id": faculty_id,
        "excel_path": excel_path,
        "cgpa_path": cgpa_path,
        "grade_path": grade_path,
        "output_dir": str(REPORTS_DIR / task_id),
        "semester": req.semester,
        "learner_type": req.learner_type,
        "format_choice": req.format_choice,
        "output_type": req.output_type,
        "slow_threshold": req.slow_threshold,
        "advanced_threshold": req.advanced_threshold,
        "common_comment": final_comment,
        "faculty_name": req.faculty_name,
        "proofs_text": proofs_text,
        "enable_signing": req.enable_signing,
    }

    try:
        celery_app.send_task("generate_report_task", args=[task_payload], task_id=task_id)
    except Exception as exc:
        await tasks_col.delete_one({"task_id": task_id})
        shutil.rmtree(task_upload_dir, ignore_errors=True)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Could not queue the report. Please try again in a moment.",
        ) from exc

    await audit_log_async(
        "report.queued",
        faculty_id=faculty_id,
        task_id=task_id,
        file_size_bytes=excel_size,
        details={
            "semester": req.semester,
            "learner_type": req.learner_type,
            "format_choice": req.format_choice,
            "output_type": req.output_type,
            "enable_signing": req.enable_signing,
        },
    )

    return GenerateReportResponse(task_id=task_id)


@app.get("/reports", response_model=ReportHistoryResponse)
async def get_reports(faculty: dict = Depends(get_current_faculty)):
    """Fetch the last 50 reports generated by this faculty."""
    from shared.mongo import get_tasks_collection
    from celery.result import AsyncResult
    from celery_app import celery_app

    from datetime import timedelta

    tasks_col = get_tasks_collection()
    cursor = tasks_col.find({"faculty_id": str(faculty["_id"])}).sort("created_at", -1).limit(50)
    docs = await cursor.to_list(length=50)

    reports = []
    for d in docs:
        # Check current status from Celery
        res = AsyncResult(d["task_id"], app=celery_app)
        status_str = res.state if res.state else d.get("status", "queued")

        has_proofs = False
        task_report_dir = REPORTS_DIR / d["task_id"]

        # Fix for old reports showing as PENDING due to Celery expiring results
        if status_str == "PENDING":
            if task_report_dir.exists() and any(task_report_dir.iterdir()):
                status_str = "SUCCESS"
            elif datetime.now(timezone.utc) - d["created_at"].replace(tzinfo=timezone.utc) > timedelta(hours=1):
                status_str = "FAILURE"
        
        if status_str == "SUCCESS" and task_report_dir.exists():
            has_proofs = any(f.name.startswith("proof_") for f in task_report_dir.iterdir())

        reports.append(
            ReportHistoryItem(
                task_id=d["task_id"],
                status=status_str,
                semester=d.get("semester", "N/A"),
                learner_type=d.get("learner_type", "N/A"),
                format_choice=d.get("format_choice", "N/A"),
                output_type=d.get("output_type", "N/A"),
                has_proofs=has_proofs,
                created_at=d["created_at"].replace(tzinfo=timezone.utc),
            )
        )

    return ReportHistoryResponse(reports=reports)


@app.get("/task-status/{task_id}", response_model=TaskStatusResponse)
async def task_status(task_id: str, faculty: dict = Depends(get_current_faculty)):
    await _require_task_owner(task_id, faculty)

    from celery.result import AsyncResult
    from celery_app import celery_app

    result = AsyncResult(task_id, app=celery_app)
    state = result.state  # PENDING | STARTED | SUCCESS | FAILURE | REVOKED

    if state == "PENDING":
        return TaskStatusResponse(task_id=task_id, status="PENDING", message="Queued — waiting for a worker.", progress=0)

    if state == "STARTED":
        meta = result.info or {}
        return TaskStatusResponse(
            task_id=task_id,
            status="STARTED",
            message=meta.get("message", "Processing…"),
            progress=meta.get("progress", 10),
        )

    if state == "SUCCESS":
        info = result.result or {}
        return TaskStatusResponse(
            task_id=task_id,
            status="SUCCESS",
            message="Report generated successfully.",
            progress=100,
            download_token=info.get("download_token"),
        )

    if state == "FAILURE":
        return TaskStatusResponse(
            task_id=task_id,
            status="FAILURE",
            message="Report generation failed.",
            progress=0,
            error=_safe_task_error(result.info),
        )

    if state == "REVOKED":
        return TaskStatusResponse(
            task_id=task_id,
            status="REVOKED",
            message="Report generation was cancelled.",
            progress=0,
        )

    return TaskStatusResponse(
        task_id=task_id,
        status="FAILURE",
        message=f"Unexpected task state ({state}).",
        progress=0,
        error="Report generation failed. Please try again.",
    )


@app.get("/download/{task_id}")
async def download_report(task_id: str, faculty: dict = Depends(get_current_faculty)):
    """Serve the generated report file from the shared volume."""
    await _require_task_owner(task_id, faculty)
    task_report_dir = REPORTS_DIR / task_id
    if not task_report_dir.exists():
        raise HTTPException(status_code=404, detail="Report not found or not yet generated.")

    files = [f for f in task_report_dir.iterdir() if not f.name.startswith("proof_")]
    if not files:
        raise HTTPException(status_code=404, detail="Report document not found.")

    if len(files) == 1:
        return FileResponse(
            path=str(files[0]),
            filename=files[0].name,
            media_type="application/octet-stream",
        )

    # Multiple files (format 5) → return as zip
    import zipfile
    import io
    from fastapi.responses import StreamingResponse

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in files:
            zf.write(fp, arcname=fp.name)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=report_{task_id[:8]}.zip"},
    )


@app.get("/download-proofs/{task_id}")
async def download_proofs(task_id: str, faculty: dict = Depends(get_current_faculty)):
    """Serve only the proof attachments in a zip archive."""
    await _require_task_owner(task_id, faculty)
    task_report_dir = REPORTS_DIR / task_id
    if not task_report_dir.exists():
        raise HTTPException(status_code=404, detail="Proofs not found.")

    proof_files = [f for f in task_report_dir.iterdir() if f.name.startswith("proof_")]
    if not proof_files:
        raise HTTPException(status_code=404, detail="No proofs attached to this report.")

    import zipfile
    import io
    from fastapi.responses import StreamingResponse

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in proof_files:
            # Optionally remove 'proof_' prefix for a cleaner zip, but keeping it is fine
            zf.write(fp, arcname=fp.name)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=proofs_{task_id[:8]}.zip"},
    )

# ===========================================================================
# Helpers
# ===========================================================================

async def _require_task_owner(task_id: str, faculty: dict) -> None:
    """Ensure the authenticated faculty owns this task (prevents IDOR on status/download)."""
    from shared.mongo import get_tasks_collection

    doc = await get_tasks_collection().find_one({"task_id": task_id})
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Task not found.",
        )
    if doc.get("faculty_id") != str(faculty["_id"]):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="You do not have access to this task.",
        )


def _safe_task_error(info) -> str:
    """Extract a user-safe error string from Celery result metadata (no tracebacks)."""
    if info is None:
        return "Report generation failed. Please try again."

    if isinstance(info, dict):
        exc_type = info.get("exc_type") or info.get("type", "")
        if exc_type in ("ReportTaskError", "shared.task_errors.ReportTaskError"):
            msg = info.get("exc_message") or info.get("message")
            if isinstance(msg, (list, tuple)) and msg:
                return str(msg[0])[:300]
            if msg:
                return str(msg)[:300]
        return "Report generation failed. Please check your Excel file and try again."

    text = str(info)
    if "Traceback" in text or "File \"" in text:
        return "Report generation failed. Please check your Excel file and try again."
    return text[:300]


def _faculty_to_profile(doc: dict) -> FacultyProfile:
    return FacultyProfile(
        **{
            "_id": str(doc["_id"]),
            "email": doc["email"],
            "name": doc["name"],
            "role": doc["role"],
            "department": doc["department"],
            "has_photo": doc.get("profile_photo") is not None,
            "has_signature": doc.get("signature_image_enc") is not None,
            "has_private_key": doc.get("private_key_enc") is not None,
            "has_certificate": doc.get("certificate_enc") is not None,
            "created_at": doc["created_at"],
        }
    )
