from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, StreamObject

from document_search.indexing import DocumentIndex


def test_indexes_and_searches_japanese_substrings_and_filenames(tmp_path: Path) -> None:
    documents = tmp_path / "資料"
    documents.mkdir()
    content_path = documents / "報告書.txt"
    content_path.write_text("東京都の検索システム", encoding="utf-8")
    filename_path = documents / "大阪計画.txt"
    filename_path.write_text("別の内容", encoding="utf-8")
    original = content_path.read_bytes()
    original_mtime = content_path.stat().st_mtime_ns
    index = DocumentIndex(tmp_path / "index.db")

    report = index.index_directory(documents)

    assert report.indexed == 2
    assert report.failed == 0
    assert [result.filename for result in index.search("京")] == ["報告書.txt"]
    assert index.search("東京")[0].matched_in == ("content",)
    assert "東京都の検索システム" in index.search("東京")[0].snippet
    assert index.search("検索シス")[0].matched_in == ("content",)
    assert index.search("大阪")[0].matched_in == ("filename",)
    assert content_path.read_bytes() == original
    assert content_path.stat().st_mtime_ns == original_mtime


def test_search_normalizes_width_and_case(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    (documents / "Unicode.txt").write_text("ＡＢＣ１２３", encoding="utf-8")
    index = DocumentIndex(tmp_path / "index.db")
    index.index_directory(documents)

    results = index.search("abc12")

    assert [result.filename for result in results] == ["Unicode.txt"]


def test_incremental_update_and_deleted_file_cleanup(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    changed = documents / "更新.txt"
    deleted = documents / "削除.txt"
    changed.write_text("旧内容", encoding="utf-8")
    deleted.write_text("削除対象", encoding="utf-8")
    index = DocumentIndex(tmp_path / "index.db")
    first = index.index_directory(documents)

    second = index.index_directory(documents)
    changed.write_text("新しい内容です", encoding="utf-8")
    new_mtime = changed.stat().st_mtime_ns + 1_000_000_000
    os.utime(changed, ns=(new_mtime, new_mtime))
    deleted.unlink()
    third = index.index_directory(documents)

    assert first.indexed == 2
    assert second.unchanged == 2
    assert second.indexed == 0
    assert third.indexed == 1
    assert third.deleted == 1
    assert index.search("旧内容") == []
    assert [result.filename for result in index.search("新")] == ["更新.txt"]
    assert index.search("削除対象") == []


def test_corrupt_file_is_recorded_without_stopping_indexing(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    (documents / "正常.txt").write_text("正常な内容", encoding="utf-8")
    (documents / "破損資料.pdf").write_bytes(b"not a pdf")
    index = DocumentIndex(tmp_path / "index.db")

    report = index.index_directory(documents)
    corrupt_results = index.search("破損")

    assert report.indexed == 2
    assert report.failed == 1
    assert len(report.issues) == 1
    assert index.search("正常")[0].status == "ok"
    assert corrupt_results[0].status == "error"
    assert corrupt_results[0].matched_in == ("filename",)


def test_skips_excluded_directories(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    for directory_name in (".venv", "build", "dist"):
        directory = documents / directory_name
        directory.mkdir()
        (directory / "除外.txt").write_text("見つからない", encoding="utf-8")
    (documents / "対象.txt").write_text("見つかる", encoding="utf-8")
    index = DocumentIndex(documents / "search-index.db")

    report = index.index_directory(documents)

    assert report.indexed == 1
    assert index.search("見つからない") == []
    assert [result.filename for result in index.search("見つかる")] == ["対象.txt"]


def test_progress_and_cancellation_do_not_remove_unseen_records(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    first_path = documents / "a.txt"
    deleted_path = documents / "b.txt"
    first_path.write_text("保持中", encoding="utf-8")
    deleted_path.write_text("削除済みだが未同期", encoding="utf-8")
    index = DocumentIndex(tmp_path / "index.db")
    index.index_directory(documents)
    deleted_path.unlink()
    first_path.write_text("更新中", encoding="utf-8")
    cancel = False
    progress: list[tuple[int, int, str]] = []

    def on_progress(current: int, total: int, path: Path) -> None:
        nonlocal cancel
        progress.append((current, total, path.name))
        if current == 1:
            cancel = True

    report = index.index_directory(
        documents,
        progress_callback=on_progress,
        should_cancel=lambda: cancel,
    )

    assert report.cancelled
    assert report.deleted == 0
    assert progress[0][:2] == (0, 1)
    assert progress[-1] == (1, 1, "a.txt")
    assert index.search("削除済み")[0].filename == "b.txt"


def test_search_returns_excel_slide_and_pdf_locations(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "東京支店"
    sheet.append(["見出し"])
    sheet.append(["対象語"])
    workbook.save(documents / "売上.xlsx")

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    box.text_frame.text = "対象語"
    presentation.save(documents / "発表.pptx")

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
    content.set_data(b"BT /F1 12 Tf 10 50 Td (Target text) Tj ET")
    page[NameObject("/Contents")] = content
    with (documents / "資料.pdf").open("wb") as stream:
        writer.write(stream)

    index = DocumentIndex(tmp_path / "index.db")
    index.index_directory(documents)

    japanese_results = {result.document_type: result for result in index.search("対象語")}
    assert japanese_results["xlsx"].locations == ("シート「東京支店」 / 行 2",)
    assert japanese_results["pptx"].locations == ("スライド 1",)
    assert index.search("Target")[0].locations == ("ページ 1",)


def test_search_filters_by_document_type(tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    (documents / "共通.txt").write_text("内容", encoding="utf-8")
    (documents / "共通.pdf").write_bytes(b"not a pdf")
    index = DocumentIndex(tmp_path / "index.db")
    index.index_directory(documents)

    results = index.search("共通", document_types={"txt"})

    assert [result.document_type for result in results] == ["txt"]
