"""FastAPI entrypoint for document image extraction."""

from __future__ import annotations

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
    save_upload,
)
from app.services.processor import process_files  # noqa: E402
from app.services.schemas import ProcessResponse  # noqa: E402


@asynccontextmanager
async def lifespan(_app: FastAPI):
    ensure_dirs()
    yield


app = FastAPI(title="ParsePics", version="1.0.0", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(FRONTEND_DIR / "index.html")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/process", response_model=ProcessResponse)
async def process(files: list[UploadFile] = File(...)) -> JSONResponse:
    if not files:
        return JSONResponse(
            status_code=400,
            content={"status": "error", "results": [], "detail": "No files uploaded"},
        )

    temp_dir = make_temp_dir()
    saved: list[Path] = []
    skipped: list[dict] = []

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
                        "error": "Only .docx and .pptx are allowed",
                    }
                )
                continue
            data = await upload.read()
            saved.append(save_upload(data, name, temp_dir))

        results = [r.model_dump() for r in process_files(saved)]
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

        return JSONResponse(content={"status": overall, "results": results})
    finally:
        cleanup_dir(temp_dir)
