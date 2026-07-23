"""FastAPI entrypoint for document image extraction."""

from __future__ import annotations

import asyncio
import logging
import sys
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# Allow `uvicorn app.api.main:app` from project root
PROJECT_ROOT = Path(__file__).resolve().parents[2]
FRONTEND_DIR = PROJECT_ROOT / "frontend"
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.services.file_manager import (  # noqa: E402
    cleanup_dir,
    ensure_dirs,
    is_allowed_file,
    make_temp_dir,
    resolve_result_dir,
    save_upload,
)
from app.services.processor import process_files  # noqa: E402
from app.services.schemas import ProcessResponse  # noqa: E402
from app.services.zip_writer import create_result_zip  # noqa: E402

logger = logging.getLogger(__name__)

# Upload guards (bytes / counts)
MAX_FILES_PER_REQUEST = 20
MAX_FILE_BYTES = 50 * 1024 * 1024  # 50 MB per file
MAX_TOTAL_BYTES = 200 * 1024 * 1024  # 200 MB total


@asynccontextmanager
async def lifespan(_app: FastAPI):
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    )
    ensure_dirs()
    yield


app = FastAPI(title="ParsePics", version="1.1.0", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(FRONTEND_DIR / "index.html")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/download/{folder_name}")
def download_result(folder_name: str):
    """Download a previously processed result folder as a zip archive."""
    result_dir = resolve_result_dir(folder_name)
    if result_dir is None:
        return JSONResponse(
            status_code=404,
            content={"detail": "Result not found"},
        )

    zip_path = result_dir.with_suffix(".zip")
    if not zip_path.is_file():
        try:
            create_result_zip(result_dir, zip_path)
        except Exception as exc:
            logger.exception("Failed to zip %s", result_dir)
            return JSONResponse(
                status_code=500,
                content={"detail": f"Failed to create zip: {exc}"},
            )

    return FileResponse(
        path=zip_path,
        media_type="application/zip",
        filename=f"{result_dir.name}.zip",
    )


@app.post("/process", response_model=ProcessResponse)
async def process(files: list[UploadFile] = File(...)) -> ProcessResponse | JSONResponse:
    if not files:
        return JSONResponse(
            status_code=400,
            content={"status": "error", "results": [], "detail": "No files uploaded"},
        )

    if len(files) > MAX_FILES_PER_REQUEST:
        return JSONResponse(
            status_code=400,
            content={
                "status": "error",
                "results": [],
                "detail": f"Too many files (max {MAX_FILES_PER_REQUEST})",
            },
        )

    temp_dir = make_temp_dir()
    saved: list[Path] = []
    skipped: list[dict] = []
    total_bytes = 0

    try:
        for upload in files:
            name = upload.filename or "unknown"
            if not is_allowed_file(name):
                skipped.append(
                    {
                        "file_name": name,
                        "base_name": Path(name).stem,
                        "status": "error",
                        "image_count": 0,
                        "skipped_image_count": 0,
                        "skip_reasons": [],
                        "error": "Only .docx and .pptx are allowed",
                    }
                )
                continue

            data = await upload.read()
            size = len(data)
            if size > MAX_FILE_BYTES:
                skipped.append(
                    {
                        "file_name": name,
                        "base_name": Path(name).stem,
                        "status": "error",
                        "image_count": 0,
                        "skipped_image_count": 0,
                        "skip_reasons": [],
                        "error": f"File exceeds {MAX_FILE_BYTES // (1024 * 1024)} MB limit",
                    }
                )
                continue

            total_bytes += size
            if total_bytes > MAX_TOTAL_BYTES:
                skipped.append(
                    {
                        "file_name": name,
                        "base_name": Path(name).stem,
                        "status": "error",
                        "image_count": 0,
                        "skipped_image_count": 0,
                        "skip_reasons": [],
                        "error": f"Total upload exceeds {MAX_TOTAL_BYTES // (1024 * 1024)} MB limit",
                    }
                )
                continue

            saved.append(save_upload(data, name, temp_dir))

        results = [
            r.model_dump()
            for r in await asyncio.to_thread(process_files, saved)
        ]
        results.extend(skipped)

        statuses = {r["status"] for r in results}
        if not results:
            overall = "error"
        elif statuses == {"success"}:
            overall = "success"
        elif "success" in statuses:
            overall = "partial"
        else:
            overall = "error"

        return ProcessResponse(status=overall, results=results)
    finally:
        cleanup_dir(temp_dir)
