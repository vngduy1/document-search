from __future__ import annotations

from pathlib import Path
from threading import Event

from PySide6.QtCore import QThread, Qt

from document_search.gui import DocumentSearchWindow
from document_search.gui import main_window as main_window_module
from document_search.gui.highlighting import normalized_match_spans
from document_search.indexing import DocumentIndex, IndexingReport


def test_japanese_ui_and_search_results(qtbot, tmp_path: Path) -> None:
    documents = tmp_path / "資料"
    documents.mkdir()
    (documents / "東京報告.txt").write_text("日本語の検索対象です", encoding="utf-8")
    database = tmp_path / "index.db"
    DocumentIndex(database).index_directory(documents)
    window = DocumentSearchWindow(database)
    qtbot.addWidget(window)
    window.show()

    assert window.windowTitle() == "ドキュメント検索"
    assert window.index_button.text() == "インデックス作成 / 更新"
    assert [
        window.results_table.horizontalHeaderItem(column).text() for column in range(5)
    ] == ["ファイル名", "パス", "種類", "一致箇所", "一致した内容"]

    window.search_edit.setText("検索")
    qtbot.waitUntil(lambda: window.results_table.rowCount() == 1)

    assert window.results_table.item(0, 0).text() == "東京報告.txt"
    assert window.results_table.item(0, 3).text() == "本文"
    assert "日本語の検索対象" in window.results_table.item(0, 4).text()
    assert window.status_label.text() == "検索結果: 1件。"
    window.results_table.selectRow(0)
    qtbot.waitUntil(lambda: bool(window.preview_edit.toPlainText()))
    assert "日本語の検索対象" in window.preview_edit.toPlainText()
    assert len(window.preview_edit.extraSelections()) == 1


def test_indexing_runs_off_the_ui_thread(qtbot, tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    database = tmp_path / "index.db"
    window = DocumentSearchWindow(database)
    qtbot.addWidget(window)
    window.show()
    window.folder_edit.setText(str(documents))
    release_worker = Event()
    called_threads: list[QThread] = []
    ui_thread = QThread.currentThread()

    def fake_index_directory(root, *, progress_callback, should_cancel):
        called_threads.append(QThread.currentThread())
        progress_callback(0, 1, Path(root))
        release_worker.wait(2)
        return IndexingReport(0, 0, 0, 0)

    window._index.index_directory = fake_index_directory
    qtbot.mouseClick(window.index_button, Qt.MouseButton.LeftButton)

    qtbot.waitUntil(lambda: bool(called_threads))
    assert called_threads[0] is not ui_thread
    assert window.cancel_button.isEnabled()
    assert not window.index_button.isEnabled()

    release_worker.set()
    qtbot.waitUntil(lambda: window._thread is None)
    assert window.index_button.isEnabled()
    assert "完了しました" in window.status_label.text()


def test_cancel_button_requests_worker_cancellation(qtbot, tmp_path: Path) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    window = DocumentSearchWindow(tmp_path / "index.db")
    qtbot.addWidget(window)
    window.show()
    window.folder_edit.setText(str(documents))
    worker_started = Event()

    def cancellable_index(root, *, progress_callback, should_cancel):
        worker_started.set()
        progress_callback(0, 1, Path(root))
        while not should_cancel():
            worker_started.wait(0.01)
        return IndexingReport(0, 0, 0, 0, cancelled=True)

    window._index.index_directory = cancellable_index
    qtbot.mouseClick(window.index_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(worker_started.is_set)
    qtbot.mouseClick(window.cancel_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: window._thread is None)

    assert "キャンセルしました" in window.status_label.text()


def test_search_runs_off_the_ui_thread(qtbot, tmp_path: Path) -> None:
    window = DocumentSearchWindow(tmp_path / "index.db")
    qtbot.addWidget(window)
    window.show()
    release_worker = Event()
    called_threads: list[QThread] = []
    ui_thread = QThread.currentThread()

    def fake_search(query, *, document_types=None):
        called_threads.append(QThread.currentThread())
        release_worker.wait(2)
        return []

    window._index.search = fake_search
    window.search_edit.setText("検索")

    qtbot.waitUntil(lambda: bool(called_threads))
    assert called_threads[0] is not ui_thread
    assert window.status_label.text() == "検索中…"

    release_worker.set()
    qtbot.waitUntil(lambda: window.status_label.text() == "検索結果: 0件。")


def test_normalized_keyword_highlighting() -> None:
    assert normalized_match_spans("ＡＢＣとabcと検索検索", "abc") == [(0, 3), (4, 7)]
    assert normalized_match_spans("検索検索", "検索") == [(0, 2), (2, 4)]


def test_filter_sort_copy_tooltip_and_open(qtbot, tmp_path: Path, monkeypatch) -> None:
    documents = tmp_path / "docs"
    documents.mkdir()
    first = documents / "a資料.txt"
    second = documents / "b資料.txt"
    first.write_text("共通語", encoding="utf-8")
    second.write_text("共通語", encoding="utf-8")
    (documents / "共通資料.pdf").write_bytes(b"not a pdf")
    database = tmp_path / "index.db"
    DocumentIndex(database).index_directory(documents)
    window = DocumentSearchWindow(database)
    qtbot.addWidget(window)
    window.show()
    window.search_edit.setText("共通")
    qtbot.waitUntil(lambda: window.results_table.rowCount() == 3)

    window.type_filter.setCurrentIndex(window.type_filter.findData("txt"))
    qtbot.waitUntil(lambda: window.results_table.rowCount() == 2)
    window.results_table.sortItems(0, Qt.SortOrder.DescendingOrder)

    assert window.results_table.item(0, 0).text() == "b資料.txt"
    assert window.results_table.isSortingEnabled()
    assert window.results_table.item(0, 1).toolTip() == str(second.resolve())

    window.results_table.selectRow(0)
    window._copy_selected_path()
    assert main_window_module.QApplication.clipboard().text() == str(second.resolve())

    opened_urls = []

    class FakeDesktopServices:
        @staticmethod
        def openUrl(url):
            opened_urls.append(url)
            return True

    monkeypatch.setattr(main_window_module, "QDesktopServices", FakeDesktopServices)
    window._open_result(0, 0)
    assert Path(opened_urls[0].toLocalFile()) == second.resolve()
