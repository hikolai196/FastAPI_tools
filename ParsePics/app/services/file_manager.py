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


def make_run_id() -> str:
    """Short unique id for a processing run output folder."""
    return uuid.uuid4().hex[:8]


def prepare_output_dir(base_name: str, run_id: str | None = None) -> Path:
    """Create result/{base_name}_{run_id}/extracted_images without wiping prior runs."""
    ensure_dirs()
    rid = run_id or make_run_id()
    out = RESULT_DIR / f"{base_name}_{rid}"
    # Extremely unlikely collision; append another segment rather than delete.
    while out.exists():
        out = RESULT_DIR / f"{base_name}_{rid}_{uuid.uuid4().hex[:4]}"
    images = out / "extracted_images"
    images.mkdir(parents=True, exist_ok=True)
    return out


def resolve_result_dir(folder_name: str) -> Path | None:
    """Resolve a result folder by name; reject path traversal."""
    name = Path(folder_name).name
    if not name or name in {".", ".."}:
        return None
    candidate = (RESULT_DIR / name).resolve()
    try:
        candidate.relative_to(RESULT_DIR.resolve())
    except ValueError:
        return None
    if candidate.is_dir():
        return candidate
    return None

def cleanup_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)


def to_project_relative(path: Path) -> str:
    """Return path relative to project root when possible."""
    try:
        return str(path.resolve().relative_to(PROJECT_ROOT.resolve())).replace("\\", "/")
    except ValueError:
        return str(path).replace("\\", "/")
