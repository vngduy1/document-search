from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pypdf import PdfReader
import xlrd

from .models import ExtractedChunk


def read_txt(path: Path) -> list[ExtractedChunk]:
    data = path.read_bytes()
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = data.decode("utf-16")
    else:
        text = _decode_text(data)
    return [ExtractedChunk(text=text, metadata={"kind": "file"})]


def _decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp932"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeError("TXT encoding is not supported (expected UTF-8, UTF-16, or CP932)")


def read_docx(path: Path) -> list[ExtractedChunk]:
    document = Document(path)
    chunks: list[ExtractedChunk] = []
    for number, paragraph in enumerate(document.paragraphs, start=1):
        if paragraph.text.strip():
            chunks.append(
                ExtractedChunk(paragraph.text, {"kind": "paragraph", "paragraph": number})
            )
    for table_number, table in enumerate(document.tables, start=1):
        for row_number, row in enumerate(table.rows, start=1):
            text = "\t".join(cell.text for cell in row.cells).strip()
            if text:
                chunks.append(
                    ExtractedChunk(
                        text,
                        {"kind": "table_row", "table": table_number, "row": row_number},
                    )
                )
    return chunks


def read_xlsx(path: Path) -> list[ExtractedChunk]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        chunks: list[ExtractedChunk] = []
        for sheet in workbook.worksheets:
            for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                text = _join_cells(row)
                if text:
                    chunks.append(
                        ExtractedChunk(
                            text,
                            {"kind": "sheet_row", "sheet": sheet.title, "row": row_number},
                        )
                    )
        return chunks
    finally:
        workbook.close()


def read_xls(path: Path) -> list[ExtractedChunk]:
    workbook = xlrd.open_workbook(path, on_demand=True)
    try:
        chunks: list[ExtractedChunk] = []
        for sheet in workbook.sheets():
            for row_index in range(sheet.nrows):
                text = _join_cells(sheet.row_values(row_index))
                if text:
                    chunks.append(
                        ExtractedChunk(
                            text,
                            {"kind": "sheet_row", "sheet": sheet.name, "row": row_index + 1},
                        )
                    )
        return chunks
    finally:
        workbook.release_resources()


def _join_cells(values: Any) -> str:
    return "\t".join(_cell_text(value) for value in values if value is not None).strip()


def _cell_text(value: Any) -> str:
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_pptx(path: Path) -> list[ExtractedChunk]:
    presentation = Presentation(path)
    chunks: list[ExtractedChunk] = []
    for slide_number, slide in enumerate(presentation.slides, start=1):
        parts: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                parts.extend(
                    paragraph.text
                    for paragraph in shape.text_frame.paragraphs
                    if paragraph.text.strip()
                )
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    row_text = "\t".join(cell.text for cell in row.cells).strip()
                    if row_text:
                        parts.append(row_text)
        text = "\n".join(parts)
        if text:
            chunks.append(ExtractedChunk(text, {"kind": "slide", "slide": slide_number}))
    return chunks


def read_pdf(path: Path) -> list[ExtractedChunk]:
    reader = PdfReader(path)
    chunks: list[ExtractedChunk] = []
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            chunks.append(ExtractedChunk(text, {"kind": "page", "page": page_number}))
    return chunks
