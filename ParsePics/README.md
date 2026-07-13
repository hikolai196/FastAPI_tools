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
uvicorn app.api.main:app --reload --host 0.0.0.0 --port 8000
```

Then open the UI: [http://127.0.0.1:8000/](http://127.0.0.1:8000/)

API:

- Health: `GET /health`
- Process: `POST /process` (multipart field name: `files`)

Example:

```bash
curl -X POST "http://127.0.0.1:8000/process" -F "files=@sample.docx" -F "files=@deck.pptx"
```

## Output layout

```
result/{document_name}/
├── extracted_images/
│   ├── img_001.png
│   ├── img_002.png
│   └── ...
├── {document_name}_extracted.docx   # or .pptx
└── {document_name}_photoatacks.xlsx
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
- A single bad image is skipped; the rest of that document continues.
- Documents with zero images still produce a modified file and a header-only Excel.
- For floating images in Word, if precise replacement is hard, the label is inserted into the nearest paragraph.
- In PowerPoint, each picture is removed and a textbox with the same label is created at the same position.
