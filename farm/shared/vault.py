"""
Fernet encryption helpers — MASTER_KEY must be set in the environment.
"""

import os

from cryptography.fernet import Fernet, InvalidToken

_raw_key = os.environ.get("MASTER_KEY", "")
if not _raw_key:
    raise RuntimeError(
        "MASTER_KEY environment variable is not set. "
        'Generate one with: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"'
    )

try:
    _fernet = Fernet(_raw_key.encode())
except Exception as exc:
    raise RuntimeError(f"MASTER_KEY is invalid: {exc}") from exc


def encrypt_bytes(data: bytes) -> bytes:
    return _fernet.encrypt(data)


def decrypt_bytes(token: bytes) -> bytes:
    try:
        return _fernet.decrypt(token)
    except InvalidToken as exc:
        raise ValueError(
            "Decryption failed — token is invalid or MASTER_KEY has changed."
        ) from exc


def encrypt_str(plaintext: str) -> bytes:
    return encrypt_bytes(plaintext.encode("utf-8"))


def decrypt_str(token: bytes) -> str:
    return decrypt_bytes(token).decode("utf-8")
