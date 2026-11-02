# Document Search

Offline Windows document search. Phase 1 provides safe text extraction for
DOCX, XLSX, XLS, PPTX, PDF, and TXT files.

## Setup

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Extraction API

```python
from document_search.extraction import extract_document

result = extract_document(r"C:\資料\報告書.pdf")
if result.ok:
    for chunk in result.chunks:
        print(chunk.text, chunk.metadata)
else:
    print(result.status, result.error)
```

Extraction is read-only. Unsupported formats and corrupt files are represented
by `ExtractionResult` values with `unsupported` or `error` status.

## Index and search

```python
from document_search.indexing import DocumentIndex

index = DocumentIndex(r"C:\検索データ\index.db")
report = index.index_directory(r"C:\資料")
results = index.search("東京")
```

Indexing is incremental and removes records for files deleted from a scanned
directory. Searches match normalized filenames and content, including one- and
two-character Japanese substrings.

## Run the GUI

```powershell
$env:PYTHONPATH = "src"
.\.venv\Scripts\python.exe -m document_search
```

The Japanese interface supports folder selection, incremental index updates,
progress and cancellation, and filename/content results with matched snippets.
Indexing, folder scanning, and search queries run outside the UI thread so the
interface remains responsive.

## Tests

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```
