from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from .models import ExtractedChunk, ExtractionResult
from .legacy import LegacyOfficeUnavailable, read_doc, read_ppt
from .readers import read_docx, read_pdf, read_pptx, read_txt, read_xls, read_xlsx


Reader = Callable[[Path], list[ExtractedChunk]]

READERS: dict[str, Reader] = {
    ".docx": read_docx,
    ".doc": read_doc,
    ".xlsx": read_xlsx,
    ".xls": read_xls,
    ".pptx": read_pptx,
    ".ppt": read_ppt,
    ".pdf": read_pdf,
    ".txt": read_txt,
}

def extract_document(path: str | Path) -> ExtractionResult:
    document_path = Path(path)
    suffix = document_path.suffix.lower()
    document_type = suffix.removeprefix(".") or "unknown"

    reader = READERS.get(suffix)
    if reader is None:
        return ExtractionResult(
            path=document_path,
            document_type=document_type,
            status="unsupported",
            error=f"Unsupported document type: {suffix or '(none)'}",
        )

    if not document_path.is_file():
        return ExtractionResult(
            path=document_path,
            document_type=document_type,
            status="error",
            error=f"File not found: {document_path}",
        )

    try:
        chunks = tuple(reader(document_path))
    except LegacyOfficeUnavailable as exc:
        return ExtractionResult(
            path=document_path,
            document_type=document_type,
            status="unsupported",
            error=str(exc),
        )
    except Exception as exc:
        return ExtractionResult(
            path=document_path,
            document_type=document_type,
            status="error",
            error=f"{type(exc).__name__}: {exc}",
        )

    return ExtractionResult(
        path=document_path,
        document_type=document_type,
        status="ok",
        chunks=chunks,
    )
