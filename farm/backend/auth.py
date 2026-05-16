"""
auth.py — JWT creation/verification and the get_current_faculty dependency.
"""

import os
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt
from passlib.context import CryptContext

from database import get_faculty_collection
from models import TokenData

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
JWT_SECRET: str = os.environ.get("JWT_SECRET", "INSECURE_DEFAULT_CHANGE_ME")
ALGORITHM: str = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES: int = int(os.environ.get("ACCESS_TOKEN_EXPIRE_MINUTES", "480"))  # 8 hours

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/auth/token")

# ---------------------------------------------------------------------------
# BCrypt compatibility shim
# ---------------------------------------------------------------------------
# passlib 1.7.4 + bcrypt >= 4.1.0 incompatibility:
#   bcrypt 4.1 removed the `__about__` module that passlib uses for version
#   detection. When passlib can't detect the bcrypt version it falls back to
#   a broken stub backend that raises:
#       ValueError: password cannot be longer than 72 bytes
#   …even for passwords of 8 characters.
#
# Fix: call bcrypt directly and skip passlib's broken introspection layer.
# passlib is retained only as a fallback for environments running bcrypt < 4.1.
#
try:
    import bcrypt as _bcrypt

    def hash_password(plain: str) -> str:
        """Hash a plain-text password using bcrypt directly (bypasses passlib)."""
        return _bcrypt.hashpw(plain.encode("utf-8"), _bcrypt.gensalt()).decode("utf-8")

    def verify_password(plain: str, hashed: str) -> bool:
        """Verify a plain-text password against a bcrypt hash (bypasses passlib)."""
        try:
            return _bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))
        except Exception:
            return False

    # CryptContext retained for completeness; not used by the functions above.
    pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

except ImportError:
    # bcrypt package not installed — fall back to passlib (only works with bcrypt < 4.1)
    pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

    def hash_password(plain: str) -> str:  # type: ignore[misc]
        return pwd_context.hash(plain)

    def verify_password(plain: str, hashed: str) -> bool:  # type: ignore[misc]
        return pwd_context.verify(plain, hashed)


# ---------------------------------------------------------------------------
# JWT helpers
# ---------------------------------------------------------------------------

def create_access_token(data: dict, expires_delta: Optional[timedelta] = None) -> str:
    payload = data.copy()
    expire = datetime.now(timezone.utc) + (
        expires_delta if expires_delta else timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    )
    payload.update({"exp": expire})
    return jwt.encode(payload, JWT_SECRET, algorithm=ALGORITHM)


def decode_token(token: str) -> TokenData:
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials.",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[ALGORITHM])
        faculty_id: str = payload.get("sub")
        if faculty_id is None:
            raise credentials_exception
        return TokenData(faculty_id=faculty_id)
    except JWTError:
        raise credentials_exception


# ---------------------------------------------------------------------------
# FastAPI dependency: get_current_faculty
# ---------------------------------------------------------------------------

async def get_current_faculty(token: str = Depends(oauth2_scheme)) -> dict:
    """
    Dependency injected into any protected route.
    Decodes the JWT, fetches the full faculty document from MongoDB,
    and returns it as a plain dict.
    """
    from bson import ObjectId  # lazy import to avoid circular deps

    token_data = decode_token(token)
    collection = get_faculty_collection()

    try:
        oid = ObjectId(token_data.faculty_id)
    except Exception:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid token subject.")

    faculty = await collection.find_one({"_id": oid})
    if faculty is None:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Faculty account not found.")

    return faculty
