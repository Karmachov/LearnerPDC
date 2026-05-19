
import io
from PIL import Image

try:
    img = Image.open(io.BytesIO(b""))
    print("Success")
except Exception as e:
    print(f"Pillow error: {e}")

try:
    # Simulating what endesive might be doing
    data = b""
    if not isinstance(data, bytes):
        raise ValueError("Invalid image format: not bytes")
    # Some libraries might raise ValueError if bytes are empty
    if len(data) == 0:
        raise ValueError("Invalid image format: bytes")
    print("Success 2")
except Exception as e:
    print(f"Simulated error: {e}")
