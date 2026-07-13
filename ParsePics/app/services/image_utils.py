"""Image conversion helpers."""

from __future__ import annotations

import io
from pathlib import Path

from PIL import Image


def bytes_to_png(data: bytes, dest: Path) -> Path:
    """Convert arbitrary image bytes to a PNG file at *dest*."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(io.BytesIO(data)) as img:
        # Preserve transparency when present; otherwise RGB is enough.
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
        else:
            img = img.convert("RGB")
        img.save(dest, format="PNG")
    return dest


def image_label(index: int) -> str:
    """Return img_001 style label (1-based index)."""
    return f"img_{index:03d}"
