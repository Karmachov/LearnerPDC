"""
models.py — Pydantic v2 request/response models and MongoDB document schemas.
"""

from __future__ import annotations

from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, EmailStr, Field, model_validator


# ===========================================================================
# Auth / Faculty Profile
# ===========================================================================

class FacultyCreate(BaseModel):
    """
    Registration payload. Accepts either `name` or `fullName` from the frontend
    so the form field key never causes a 422 field-not-found validation error.
    """
    email: EmailStr = Field(..., description="Faculty email address")
    password: str = Field(..., min_length=8, description="Minimum 8 characters")
    # Primary field name. The model_validator below also accepts `fullName`.
    name: str = Field(..., min_length=2, description="Faculty full name")
    role: str = Field(default="Faculty", description="e.g. Faculty, HOD, Lab Instructor")
    department: str = Field(default="Computer Science and Engineering")

    model_config = {"populate_by_name": True}

    @model_validator(mode="before")
    @classmethod
    def accept_full_name_alias(cls, values: dict) -> dict:
        """
        If the client sends `fullName` instead of `name` (or in addition to it),
        map it to `name` so neither spelling triggers a 422 validation failure.
        """
        if isinstance(values, dict):
            if "fullName" in values and "name" not in values:
                values["name"] = values.pop("fullName")
            elif "fullName" in values:
                # Both present — prefer explicit `name`, drop the alias
                values.pop("fullName")
        return values


class FacultyLogin(BaseModel):
    email: EmailStr
    password: str


class FacultyProfile(BaseModel):
    """Public-safe view of a faculty document (no secrets)."""
    id: str = Field(alias="_id")
    email: EmailStr
    name: str
    role: str
    department: str
    has_photo: bool = False
    has_signature: bool = False
    has_private_key: bool = False
    has_certificate: bool = False
    created_at: datetime

    model_config = {"populate_by_name": True}


class Token(BaseModel):
    access_token: str
    token_type: str = "bearer"


class TokenData(BaseModel):
    faculty_id: Optional[str] = None


# ===========================================================================
# Report Generation
# ===========================================================================

class ReportRequest(BaseModel):
    semester: str = Field(..., description="Roman numeral, e.g. 'III', 'V'")
    learner_type: Literal["slow", "advanced"]
    format_choice: Literal["1", "2", "3", "4", "5"]
    output_type: Literal["word", "pdf"]
    slow_threshold: float = Field(default=40.0, ge=0.0, le=100.0)
    advanced_threshold: float = Field(default=90.0, ge=0.0, le=100.0)
    common_comment: str = Field(default="")
    faculty_name: str = Field(default="")
    enable_signing: bool = False
    proof_types: list[str] = Field(default_factory=list)
    # Only required if enable_signing=True AND the private key is NOT yet stored in DB
    key_password: Optional[str] = None


class TaskStatusResponse(BaseModel):
    task_id: str
    status: Literal["PENDING", "STARTED", "SUCCESS", "FAILURE", "REVOKED"]
    progress: int = Field(default=0, ge=0, le=100)
    message: str = ""
    download_token: Optional[str] = None  # opaque token to fetch the file
    error: Optional[str] = None


class GenerateReportResponse(BaseModel):
    task_id: str
    message: str = "Report generation queued successfully."


class ReportHistoryItem(BaseModel):
    task_id: str
    status: str
    semester: str
    learner_type: str
    format_choice: str
    output_type: str
    has_proofs: bool = False
    created_at: datetime


class ReportHistoryResponse(BaseModel):
    reports: list[ReportHistoryItem]


# ===========================================================================
# Profile updates (multipart handled separately in routes)
# ===========================================================================

class ProfileUpdateResponse(BaseModel):
    message: str


# ===========================================================================
# MongoDB Document Shapes (dicts — for documentation / type hints only)
# ===========================================================================
# These are NOT Pydantic models; Motor returns plain dicts.
#
# FacultyDoc = {
#     "_id": ObjectId,
#     "email": str,
#     "hashed_password": str,
#     "name": str,
#     "role": str,
#     "department": str,
#     "signature_image_enc": bytes | None,   # Fernet-encrypted PNG/JPG bytes
#     "private_key_enc":     bytes | None,   # Fernet-encrypted PEM bytes
#     "certificate_enc":     bytes | None,   # Fernet-encrypted PEM bytes
#     "key_password_enc":    bytes | None,   # Fernet-encrypted passphrase
#     "created_at":          datetime,
# }
