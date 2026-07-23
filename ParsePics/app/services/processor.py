"""Orchestrate per-file document image extraction."""

from __future__ import annotations

from pathlib import Path

from app.services.excel_writer import write_photo_excel
from app.services.extractors.docx_extractor import process_docx
from app.services.extractors.pptx_extractor import process_pptx
from app.services.file_manager import (
    make_run_id,
    prepare_output_dir,
    to_project_relative,
)
from app.services.schemas import FileResult
from app.services.zip_writer import create_result_zip


def process_file(source: Path) -> FileResult:
    """Process a single .docx / .pptx file into result/{base_name}_{run_id}/."""
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

    run_id = make_run_id()
    try:
        out_dir = prepare_output_dir(base_name, run_id=run_id)
        # Folder name may gain an extra suffix on collision; keep the real stem.
        folder_name = out_dir.name
        actual_run_id = folder_name[len(base_name) + 1 :] if folder_name.startswith(f"{base_name}_") else run_id
        images_dir = out_dir / "extracted_images"

        if suffix == ".docx":
            modified = out_dir / f"{base_name}_extracted.docx"
            extract = process_docx(source, images_dir, modified)
        else:
            modified = out_dir / f"{base_name}_extracted.pptx"
            extract = process_pptx(source, images_dir, modified)

        excel_path = out_dir / f"{base_name}_photoattacks.xlsx"
        write_photo_excel(excel_path, extract.entries)
        create_result_zip(out_dir)

        return FileResult(
            file_name=file_name,
            base_name=base_name,
            status="success",
            image_count=len(extract.entries),
            skipped_image_count=len(extract.skip_reasons),
            skip_reasons=list(extract.skip_reasons),
            run_id=actual_run_id,
            output_dir=to_project_relative(out_dir),
            modified_file=to_project_relative(modified),
            excel_file=to_project_relative(excel_path),
            images_dir=to_project_relative(images_dir),
            download_url=f"/download/{folder_name}",
        )
    except Exception as exc:
        return FileResult(
            file_name=file_name,
            base_name=base_name,
            status="error",
            run_id=run_id,
            error=str(exc),
        )


def process_files(sources: list[Path]) -> list[FileResult]:
    """Process multiple files; one failure does not stop the rest."""
    return [process_file(path) for path in sources]
