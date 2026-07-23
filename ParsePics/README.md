# ParsePics

Extract images from `.docx` / `.pptx`, save as PNG, replace image slots with labels (`img_001`…), and build an Excel mapping with embedded pictures.

## Setup

```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# macOS / Linux
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
uvicorn app.api.main:app --reload --host 127.0.0.1 --port 8000
```

Then open the UI: [http://127.0.0.1:8000/](http://127.0.0.1:8000/)

API:

- Health: `GET /health`
- Process: `POST /process` (multipart field name: `files`)
- Download: `GET /download/{folder_name}` (zip of that result run)

Limits: up to **20 files** per request, **50 MB** per file, **200 MB** total.

Example:

```bash
curl -X POST "http://127.0.0.1:8000/process" -F "files=@sample.docx" -F "files=@deck.pptx"
```

## Output layout

Each run gets a unique folder so re-uploading the same filename does not wipe earlier results:

```
result/{document_name}_{run_id}/
├── extracted_images/
│   ├── img_001.png
│   ├── img_002.png
│   └── ...
├── {document_name}_extracted.docx   # or .pptx
└── {document_name}_photoattacks.xlsx

result/{document_name}_{run_id}.zip   # same contents, ready to download
```

Excel columns: `name` (e.g. `img_001`), `pic` (embedded PNG).

## Tests

```bash
pytest -q
```

## Notes

- Frontend is plain HTML/CSS/JS under `frontend/`, served by FastAPI.
- Only `.docx` and `.pptx` are accepted.
- One file failing does not stop others.
- A single bad image is skipped (with a reason in the API/UI); the rest of that document continues.
- Documents with zero images still produce a modified file and a header-only Excel.
- For floating images in Word, if precise replacement is hard, the label is inserted into the nearest paragraph.
- In PowerPoint, each picture is removed and a textbox with the same label is created at the same position.
- Prefer binding to `127.0.0.1` for local use; there is no authentication.

---

## How to experience the changes

After starting the app (`uvicorn …` above), open [http://127.0.0.1:8000/](http://127.0.0.1:8000/).

### 1. Unique output folders (no overwrite)

1. Upload a file named e.g. `report.docx` and click **開始處理**.
2. Note the **輸出目錄** path — it looks like `result/report_a1b2c3d4/`.
3. Upload the same `report.docx` again and process once more.
4. You get a **second** folder (`result/report_…` with a different id). The first run’s files remain on disk.

### 2. Download ZIP from the UI

1. Process any `.docx` or `.pptx` that contains images.
2. In the result card, click **下載 ZIP**.
3. The browser downloads `{document_name}_{run_id}.zip` containing the extracted PNGs, modified Office file, and Excel mapping.
4. Or call the API: `GET /download/report_a1b2c3d4` (use the folder name from the result).

### 3. Upload limits

1. Try selecting more than 20 files, or a file larger than 50 MB — the UI rejects them with a clear status message.
2. Oversized files sent to the API are returned as per-file `error` rows (still HTTP 200 for mixed batches); too many files returns HTTP 400.

### 4. Skip reasons for bad images

1. Process a normal document — **略過圖片** stays hidden when nothing was skipped.
2. If a document has a broken/unreadable image blob, processing still succeeds for the good images; the card shows a **partial** badge, a skip count, and reason lines (also in the JSON `skip_reasons` field).

### 5. Excel filename fix

1. After a successful run, open the result folder or ZIP.
2. The Excel file is named `{document_name}_photoattacks.xlsx` (corrected spelling).
