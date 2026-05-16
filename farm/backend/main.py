"""
main.py — FastAPI application: all routes, lifespan, CORS.
"""

from __future__ import annotations

import os
import re
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
from fastapi.responses import FileResponse
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
from models import (
    FacultyCreate,
    FacultyProfile,
    GenerateReportResponse,
    ProfileUpdateResponse,
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

    if not key_bytes.strip().startswith(b"-----BEGIN"):
        raise HTTPException(status_code=400, detail="Private key must be a valid PEM file.")
    if not cert_bytes.strip().startswith(b"-----BEGIN"):
        raise HTTPException(status_code=400, detail="Certificate must be a valid PEM file.")

    col = get_faculty_collection()
    await col.update_one(
        {"_id": faculty["_id"]},
        {
            "$set": {
                "private_key_enc": encrypt_bytes(key_bytes),
                "certificate_enc": encrypt_bytes(cert_bytes),
                "key_password_enc": encrypt_str(key_password),
            }
        },
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
    with open(excel_path, "wb") as f:
        f.write(await excel_file.read())

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

    # --- Signing: verify keys exist in DB if needed ---
    if req.enable_signing:
        if not faculty.get("private_key_enc") or not faculty.get("certificate_enc"):
            raise HTTPException(
                status_code=400,
                detail="Digital signing is enabled but no private key/certificate is stored. "
                       "Please upload them via Profile Settings first.",
            )

    # --- Dispatch Celery task ---
    from celery_app import celery_app  # import here to avoid startup side-effects

    task_payload = {
        "task_id": task_id,
        "faculty_id": str(faculty["_id"]),
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
        "common_comment": req.common_comment,
        "faculty_name": req.faculty_name,
        "enable_signing": req.enable_signing,
    }

    celery_app.send_task("generate_report_task", args=[task_payload], task_id=task_id)

    return GenerateReportResponse(task_id=task_id)


@app.get("/task-status/{task_id}", response_model=TaskStatusResponse)
async def task_status(task_id: str, _: dict = Depends(get_current_faculty)):
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
            error=str(result.info),
        )

    return TaskStatusResponse(task_id=task_id, status=state, message=state)  # type: ignore[arg-type]


@app.get("/download/{task_id}")
async def download_report(task_id: str, _: dict = Depends(get_current_faculty)):
    """Serve the generated report file from the shared volume."""
    task_report_dir = REPORTS_DIR / task_id
    if not task_report_dir.exists():
        raise HTTPException(status_code=404, detail="Report not found or not yet generated.")

    files = list(task_report_dir.iterdir())
    if not files:
        raise HTTPException(status_code=404, detail="Report directory is empty.")

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


# ===========================================================================
# Helpers
# ===========================================================================

def _faculty_to_profile(doc: dict) -> FacultyProfile:
    return FacultyProfile(
        **{
            "_id": str(doc["_id"]),
            "email": doc["email"],
            "name": doc["name"],
            "role": doc["role"],
            "department": doc["department"],
            "has_signature": doc.get("signature_image_enc") is not None,
            "has_private_key": doc.get("private_key_enc") is not None,
            "has_certificate": doc.get("certificate_enc") is not None,
            "created_at": doc["created_at"],
        }
    )
