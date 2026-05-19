"""
Validate PEM private keys and certificates at upload time.
"""

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate


def normalize_passphrase(password: str | None) -> str:
    """Strip whitespace — must match on upload, API preflight, and worker signing."""
    return (password or "").strip()


def _public_key_pem(public_key) -> bytes:
    return public_key.public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo,
    )


def validate_pem_key_pair(key_bytes: bytes, cert_bytes: bytes, password: str) -> None:
    """
    Parse and verify a PEM private key / certificate pair.

    Raises ValueError with a user-facing message on invalid PEM, wrong
    passphrase, or a key that does not match the certificate.
    """
    if not key_bytes.strip():
        raise ValueError("Private key file is empty.")
    if not cert_bytes.strip():
        raise ValueError("Certificate file is empty.")

    password = normalize_passphrase(password)
    pw = password.encode("utf-8") if password else None

    try:
        private_key = load_pem_private_key(key_bytes, password=pw)
    except TypeError:
        raise ValueError(
            "Private key is encrypted; provide the correct passphrase."
        ) from None
    except ValueError as exc:
        msg = str(exc).lower()
        if "password" in msg or "decrypt" in msg:
            raise ValueError(
                "Incorrect key passphrase. Re-upload your private key in Profile Settings "
                "and enter the passphrase that unlocks the PEM file."
            ) from None
        raise ValueError(f"Invalid private key: {exc}") from None

    try:
        certificate = load_pem_x509_certificate(cert_bytes)
    except ValueError as exc:
        raise ValueError(f"Invalid certificate PEM: {exc}") from None

    cert_pub = certificate.public_key()
    key_pub = private_key.public_key()

    if _public_key_pem(cert_pub) != _public_key_pem(key_pub):
        raise ValueError("Private key does not match the uploaded certificate.")
