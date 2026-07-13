"""Filesystem helpers for temp uploads and result output."""

from __future__ import annotations

import shutil
import uuid
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RESULT_DIR = PROJECT_ROOT / "result"
TEMP_DIR = PROJECT_ROOT / "temp"

ALLOWED_EXTENSIONS = {".docx", ".pptx"}


def ensure_dirs() -> None:
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def make_temp_dir() -> Path:
    ensure_dirs()
    path = TEMP_DIR / uuid.uuid4().hex
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_upload(data: bytes, filename: str, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / Path(filename).name
    path.write_bytes(data)
    return path


def prepare_output_dir(base_name: str) -> Path:
    """Create (or recreate) result/{base_name}/extracted_images."""
    ensure_dirs()
    out = RESULT_DIR / base_name
    if out.exists():
        shutil.rmtree(out)
    images = out / "extracted_images"
    images.mkdir(parents=True, exist_ok=True)
    return out


def cleanup_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)


def to_project_relative(path: Path) -> str:
    """Return path relative to project root when possible."""
    try:
        return str(path.resolve().relative_to(PROJECT_ROOT.resolve())).replace("\\", "/")
    except ValueError:
        return str(path).replace("\\", "/")
