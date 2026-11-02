from pathlib import Path

from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, StreamObject
import xlwt

from document_search.extraction import extract_document


def test_txt_handles_japanese_unicode_path_and_cp932(tmp_path: Path) -> None:
    path = tmp_path / "日本語の資料.txt"
    path.write_bytes("検索テスト\n東京".encode("cp932"))

    result = extract_document(path)

    assert result.ok
    assert result.chunks[0].text == "検索テスト\n東京"
    assert result.chunks[0].metadata == {"kind": "file"}


def test_docx_extracts_paragraphs_and_table_rows(tmp_path: Path) -> None:
    path = tmp_path / "報告書.docx"
    document = Document()
    document.add_paragraph("日本語の段落")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "項目"
    table.cell(0, 1).text = "値"
    document.save(path)

    result = extract_document(path)

    assert result.ok
    assert [chunk.text for chunk in result.chunks] == ["日本語の段落", "項目\t値"]
    assert result.chunks[1].metadata == {"kind": "table_row", "table": 1, "row": 1}


def test_xlsx_preserves_sheet_and_row_metadata(tmp_path: Path) -> None:
    path = tmp_path / "売上.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "東京支店"
    sheet.append(["商品", "数量"])
    sheet.append(["りんご", 2])
    workbook.save(path)

    result = extract_document(path)

    assert result.ok
    assert result.chunks[1].text == "りんご\t2"
    assert result.chunks[1].metadata == {"kind": "sheet_row", "sheet": "東京支店", "row": 2}


def test_xls_preserves_sheet_and_row_metadata(tmp_path: Path) -> None:
    path = tmp_path / "旧売上.xls"
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("大阪支店")
    sheet.write(0, 0, "商品")
    sheet.write(1, 0, "みかん")
    sheet.write(1, 1, 3)
    workbook.save(str(path))

    result = extract_document(path)

    assert result.ok
    assert result.chunks[1].text == "みかん\t3"
    assert result.chunks[1].metadata == {"kind": "sheet_row", "sheet": "大阪支店", "row": 2}


def test_pptx_preserves_slide_metadata(tmp_path: Path) -> None:
    path = tmp_path / "発表.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    box.text_frame.text = "第一スライド"
    presentation.save(path)

    result = extract_document(path)

    assert result.ok
    assert result.chunks[0].text == "第一スライド"
    assert result.chunks[0].metadata == {"kind": "slide", "slide": 1}


def test_pdf_preserves_page_metadata(tmp_path: Path) -> None:
    path = tmp_path / "ページ.pdf"
    writer = PdfWriter()
    page = writer.add_blank_page(width=100, height=100)
    font = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
        }
    )
    page[NameObject("/Resources")] = DictionaryObject(
        {NameObject("/Font"): DictionaryObject({NameObject("/F1"): font})}
    )
    content = StreamObject()
    content.set_data(b"BT /F1 12 Tf 10 50 Td (Page text) Tj ET")
    page[NameObject("/Contents")] = content
    with path.open("wb") as stream:
        writer.write(stream)

    result = extract_document(path)

    assert result.ok
    assert result.chunks[0].text.strip() == "Page text"
    assert result.chunks[0].metadata == {"kind": "page", "page": 1}


def test_unknown_format_is_unsupported(tmp_path: Path) -> None:
    result = extract_document(tmp_path / "画像.png")
    assert result.status == "unsupported"
    assert result.error


def test_missing_supported_file_returns_error(tmp_path: Path) -> None:
    result = extract_document(tmp_path / "ない.doc")
    assert result.status == "error"
    assert "File not found" in result.error


def test_corrupt_supported_files_return_error(tmp_path: Path) -> None:
    for extension in ("docx", "xlsx", "pptx", "pdf"):
        path = tmp_path / f"破損.{extension}"
        path.write_bytes(b"not a valid document")

        result = extract_document(path)

        assert result.status == "error"
        assert result.chunks == ()
        assert result.error
