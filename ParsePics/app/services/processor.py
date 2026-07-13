"""Orchestrate per-file document image extraction."""

from __future__ import annotations

from pathlib import Path

from app.services.excel_writer import write_photo_excel
from app.services.extractors.docx_extractor import process_docx
from app.services.extractors.pptx_extractor import process_pptx
from app.services.file_manager import (
    prepare_output_dir,
    to_project_relative,
)
from app.services.schemas import FileResult


def process_file(source: Path) -> FileResult:
    """Process a single .docx / .pptx file into result/{base_name}/."""
    file_name = source.name
    suffix = source.suffix.lower()
    base_name = source.stem

    if suffix not in {".docx", ".pptx"}:
        return FileResult(
            file_name=file_name,
            base_name=base_name,
            status="error",
            error=f"Unsupported file type: {suffix}",
        )

    try:
        out_dir = prepare_output_dir(base_name)
        images_dir = out_dir / "extracted_images"

        if suffix == ".docx":
            modified = out_dir / f"{base_name}_extracted.docx"
            entries = process_docx(source, images_dir, modified)
        else:
            modified = out_dir / f"{base_name}_extracted.pptx"
            entries = process_pptx(source, images_dir, modified)

        excel_path = out_dir / f"{base_name}_photoatacks.xlsx"
        write_photo_excel(excel_path, entries)

        return FileResult(
            file_name=file_name,
            base_name=base_name,
            status="success",
            image_count=len(entries),
            output_dir=to_project_relative(out_dir),
            modified_file=to_project_relative(modified),
            excel_file=to_project_relative(excel_path),
            images_dir=to_project_relative(images_dir),
        )
    except Exception as exc:
        return FileResult(
            file_name=file_name,
            base_name=base_name,
            status="error",
            error=str(exc),
        )


def process_files(sources: list[Path]) -> list[FileResult]:
    """Process multiple files; one failure does not stop the rest."""
    return [process_file(path) for path in sources]
