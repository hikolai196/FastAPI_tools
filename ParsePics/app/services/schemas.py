"""Shared response / result models."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

OverallStatus = Literal["success", "partial", "error"]
FileStatus = Literal["success", "error"]


class FileResult(BaseModel):
    file_name: str
    base_name: str
    status: FileStatus
    image_count: int = 0
    skipped_image_count: int = 0
    skip_reasons: list[str] = Field(default_factory=list)
    run_id: str | None = None
    output_dir: str | None = None
    modified_file: str | None = None
    excel_file: str | None = None
    images_dir: str | None = None
    download_url: str | None = None
    error: str | None = None


class ProcessResponse(BaseModel):
    status: OverallStatus = Field(description="overall status: success | partial | error")
    results: list[FileResult]