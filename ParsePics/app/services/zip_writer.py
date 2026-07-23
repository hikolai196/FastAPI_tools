"""Create downloadable zip archives of result folders."""

from __future__ import annotations

import zipfile
from pathlib import Path


def create_result_zip(result_dir: Path, zip_path: Path | None = None) -> Path:
    """
    Zip everything under result_dir into a sibling .zip (or *zip_path*).

    Archive member paths are relative to result_dir's parent so the folder
    name is preserved inside the zip.
    """
    if not result_dir.is_dir():
        raise FileNotFoundError(f"Result directory not found: {result_dir}")

    dest = zip_path or result_dir.with_suffix(".zip")
    dest.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(result_dir.rglob("*")):
            if path.is_file():
                arcname = path.relative_to(result_dir.parent)
                zf.write(path, arcname.as_posix())
    return dest
