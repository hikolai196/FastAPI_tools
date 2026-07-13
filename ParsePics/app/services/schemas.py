"""Shared response / result models."""

from __future__ import annotations

from pydantic import BaseModel, Field


class FileResult(BaseModel):
    file_name: str
    base_name: str
    status: str
    image_count: int = 0
    output_dir: str | None = None
    modified_file: str | None = None
    excel_file: str | None = None
    images_dir: str | None = None
    error: str | None = None


class ProcessResponse(BaseModel):
    status: str = Field(description="overall status: success | partial | error")
    results: list[FileResult]
