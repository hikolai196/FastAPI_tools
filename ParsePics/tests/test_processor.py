"""Unit / integration tests for ParsePics core + API."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.services.file_manager import RESULT_DIR, ensure_dirs
from app.services.image_utils import bytes_to_png, image_label
from app.services.processor import process_file, process_files
from tests.conftest import (
    create_docx_no_images,
    create_docx_with_images,
    create_pptx_no_images,
    create_pptx_with_images,
    excel_has_images,
    make_png_bytes,
)


@pytest.fixture(autouse=True)
def _clean_result(tmp_path, monkeypatch):
    """Redirect result/temp under tmp to keep the repo clean."""
    import app.services.file_manager as fm

    result = tmp_path / "result"
    temp = tmp_path / "temp"
    result.mkdir()
    temp.mkdir()
    monkeypatch.setattr(fm, "RESULT_DIR", result)
    monkeypatch.setattr(fm, "TEMP_DIR", temp)
    monkeypatch.setattr(fm, "PROJECT_ROOT", tmp_path)
    ensure_dirs()
    yield result


def test_image_label():
    assert image_label(1) == "img_001"
    assert image_label(12) == "img_012"


def test_bytes_to_png(tmp_path: Path):
    dest = tmp_path / "out.png"
    bytes_to_png(make_png_bytes((10, 20, 30)), dest)
    assert dest.exists()
    assert dest.stat().st_size > 0


def test_process_docx_with_images(_clean_result, tmp_path: Path):
    src = create_docx_with_images(
        tmp_path / "report.docx",
        [(255, 0, 0), (0, 255, 0), (0, 0, 255)],
    )
    result = process_file(src)
    assert result.status == "success"
    assert result.image_count == 3

    out = _clean_result / "report"
    assert (out / "extracted_images" / "img_001.png").exists()
    assert (out / "extracted_images" / "img_002.png").exists()
    assert (out / "extracted_images" / "img_003.png").exists()
    assert (out / "report_extracted.docx").exists()
    excel = out / "report_photoatacks.xlsx"
    assert excel.exists()
    assert excel_has_images(excel) == 3

    wb = load_workbook(excel)
    ws = wb.active
    assert ws["A1"].value == "name"
    assert ws["B1"].value == "pic"
    assert ws["A2"].value == "img_001"
    assert ws["A4"].value == "img_003"


def test_process_docx_no_images(_clean_result, tmp_path: Path):
    src = create_docx_no_images(tmp_path / "empty.docx")
    result = process_file(src)
    assert result.status == "success"
    assert result.image_count == 0
    out = _clean_result / "empty"
    assert (out / "empty_extracted.docx").exists()
    excel = out / "empty_photoatacks.xlsx"
    assert excel.exists()
    assert excel_has_images(excel) == 0
    wb = load_workbook(excel)
    assert wb.active["A1"].value == "name"


def test_process_pptx_with_images(_clean_result, tmp_path: Path):
    src = create_pptx_with_images(
        tmp_path / "deck.pptx",
        [(255, 0, 0), (0, 255, 0)],
    )
    result = process_file(src)
    assert result.status == "success"
    assert result.image_count == 2

    out = _clean_result / "deck"
    assert (out / "extracted_images" / "img_001.png").exists()
    assert (out / "extracted_images" / "img_002.png").exists()
    assert (out / "deck_extracted.pptx").exists()
    excel = out / "deck_photoatacks.xlsx"
    assert excel.exists()
    assert excel_has_images(excel) == 2

    # Modified pptx should contain textboxes with labels, not the original pictures
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(str(out / "deck_extracted.pptx"))
    texts = []
    pics = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    pics += 1
            except NotImplementedError:
                pass
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
    assert pics == 0
    assert "img_001" in texts
    assert "img_002" in texts


def test_process_pptx_no_images(_clean_result, tmp_path: Path):
    src = create_pptx_no_images(tmp_path / "blank.pptx")
    result = process_file(src)
    assert result.status == "success"
    assert result.image_count == 0
    assert (_clean_result / "blank" / "blank_extracted.pptx").exists()
    assert (_clean_result / "blank" / "blank_photoatacks.xlsx").exists()


def test_process_files_isolates_failures(_clean_result, tmp_path: Path):
    good = create_docx_with_images(tmp_path / "good.docx", [(1, 2, 3)])
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a real docx")
    results = process_files([good, bad])
    assert len(results) == 2
    assert results[0].status == "success"
    assert results[1].status == "error"


def test_health_and_process_api(_clean_result, tmp_path: Path):
    from app.api.main import app

    client = TestClient(app)
    assert client.get("/health").json() == {"status": "ok"}

    home = client.get("/")
    assert home.status_code == 200
    assert "ParsePics" in home.text
    assert client.get("/static/styles.css").status_code == 200
    assert client.get("/static/app.js").status_code == 200

    docx = create_docx_with_images(tmp_path / "api_doc.docx", [(9, 9, 9)])
    pptx = create_pptx_with_images(tmp_path / "api_ppt.pptx", [(8, 8, 8)])

    with docx.open("rb") as f1, pptx.open("rb") as f2:
        resp = client.post(
            "/process",
            files=[
                ("files", ("api_doc.docx", f1, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
                ("files", ("api_ppt.pptx", f2, "application/vnd.openxmlformats-officedocument.presentationml.presentation")),
            ],
        )

    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "success"
    assert len(body["results"]) == 2
    assert all(r["status"] == "success" for r in body["results"])
    assert body["results"][0]["image_count"] == 1
    assert body["results"][1]["image_count"] == 1


def test_api_rejects_unsupported(_clean_result, tmp_path: Path):
    from app.api.main import app

    client = TestClient(app)
    junk = tmp_path / "note.txt"
    junk.write_text("hello")
    with junk.open("rb") as f:
        resp = client.post(
            "/process",
            files=[("files", ("note.txt", f, "text/plain"))],
        )
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "error"
    assert body["results"][0]["status"] == "error"
