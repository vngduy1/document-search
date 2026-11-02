from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import QThread, QThreadPool, QTimer, Qt, QUrl, Slot
from PySide6.QtGui import QCloseEvent, QColor, QDesktopServices, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from document_search.indexing import DocumentIndex, IndexingReport, SearchResult

from .highlighting import normalized_match_spans
from .workers import IndexWorker, SearchWorker


class DocumentSearchWindow(QMainWindow):
    def __init__(self, database_path: str | Path | None = None) -> None:
        super().__init__()
        self.setWindowTitle("ドキュメント検索")
        self.resize(1100, 750)
        self._index = DocumentIndex(database_path or _default_database_path())
        self._thread: QThread | None = None
        self._worker: IndexWorker | None = None
        self._closing = False
        self._search_pool = QThreadPool.globalInstance()
        self._search_generation = 0
        self._build_ui()

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(200)
        self._search_timer.timeout.connect(self._run_search)
        self.search_edit.textChanged.connect(self._queue_search)
        self.type_filter.currentIndexChanged.connect(self._queue_search)

    def _build_ui(self) -> None:
        central = QWidget(self)
        layout = QVBoxLayout(central)

        folder_layout = QGridLayout()
        folder_layout.addWidget(QLabel("対象フォルダー"), 0, 0)
        self.folder_edit = QLineEdit()
        self.folder_edit.setReadOnly(True)
        self.folder_edit.setPlaceholderText("検索対象のフォルダーを選択")
        folder_layout.addWidget(self.folder_edit, 0, 1)
        self.folder_button = QPushButton("フォルダーを選択…")
        self.folder_button.clicked.connect(self._select_folder)
        folder_layout.addWidget(self.folder_button, 0, 2)
        layout.addLayout(folder_layout)

        action_layout = QHBoxLayout()
        self.index_button = QPushButton("インデックス作成 / 更新")
        self.index_button.clicked.connect(self._start_indexing)
        action_layout.addWidget(self.index_button)
        self.cancel_button = QPushButton("キャンセル")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self._cancel_indexing)
        action_layout.addWidget(self.cancel_button)
        action_layout.addStretch()
        layout.addLayout(action_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("待機中")
        layout.addWidget(self.progress_bar)

        search_layout = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("ファイル名または内容を検索")
        self.search_edit.setClearButtonEnabled(True)
        search_layout.addWidget(self.search_edit)
        search_layout.addWidget(QLabel("種類"))
        self.type_filter = QComboBox()
        for label, value in (
            ("すべての種類", None),
            ("Word (DOCX)", "docx"),
            ("Excel (XLSX)", "xlsx"),
            ("Excel (XLS)", "xls"),
            ("PowerPoint (PPTX)", "pptx"),
            ("PDF", "pdf"),
            ("テキスト (TXT)", "txt"),
            ("旧式 Word (DOC)", "doc"),
            ("旧式 PowerPoint (PPT)", "ppt"),
        ):
            self.type_filter.addItem(label, value)
        search_layout.addWidget(self.type_filter)
        layout.addLayout(search_layout)

        self.results_table = QTableWidget(0, 5)
        self.results_table.setHorizontalHeaderLabels(
            ["ファイル名", "パス", "種類", "一致箇所", "一致した内容"]
        )
        self.results_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSortingEnabled(True)
        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._show_result_menu)
        self.results_table.cellDoubleClicked.connect(self._open_result)
        self.results_table.itemSelectionChanged.connect(self._show_selected_preview)
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.results_table)

        self.preview_location_label = QLabel("プレビュー")
        layout.addWidget(self.preview_location_label)
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        self.preview_edit.setPlaceholderText("検索結果を選択すると内容を表示します。")
        self.preview_edit.setMaximumHeight(190)
        layout.addWidget(self.preview_edit)

        self.status_label = QLabel("フォルダーを選択してインデックスを作成してください。")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        self.setCentralWidget(central)

    @Slot()
    def _select_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "対象フォルダーを選択")
        if folder:
            self.folder_edit.setText(folder)
            self._set_status("「インデックス作成 / 更新」を押してください。")

    @Slot()
    def _start_indexing(self) -> None:
        if self._thread is not None:
            return
        folder_text = self.folder_edit.text().strip()
        if not folder_text or not Path(folder_text).is_dir():
            self._set_status("有効な対象フォルダーを選択してください。", error=True)
            return

        self._set_indexing_state(True)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFormat("ファイルを確認中…")
        self._set_status("インデックスを更新しています…")

        thread = QThread(self)
        worker = IndexWorker(self._index, Path(folder_text))
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_index_finished)
        worker.failed.connect(self._on_index_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        thread.finished.connect(self._thread_finished)
        self._thread = thread
        self._worker = worker
        thread.start()

    @Slot()
    def _cancel_indexing(self) -> None:
        if self._worker is not None:
            self._worker.request_cancel()
            self.cancel_button.setEnabled(False)
            self._set_status("キャンセルしています…")

    @Slot(int, int, str)
    def _on_progress(self, current: int, total: int, path: str) -> None:
        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(current)
        filename = Path(path).name
        self.progress_bar.setFormat(f"{current} / {total}  {filename}")

    @Slot(object)
    def _on_index_finished(self, report: IndexingReport) -> None:
        self._set_indexing_state(False)
        if report.cancelled:
            self._set_status(
                f"キャンセルしました。更新: {report.indexed}件、未変更: {report.unchanged}件。"
            )
        else:
            message = (
                f"完了しました。更新: {report.indexed}件、"
                f"未変更: {report.unchanged}件、削除: {report.deleted}件。"
            )
            if report.failed:
                message += f" エラー: {report.failed}件。"
            self._set_status(message, error=bool(report.failed))
        self.status_label.setToolTip("\n".join(report.issues[:10]))
        if self.search_edit.text().strip():
            self._search_generation += 1
            self._run_search()

    @Slot(str)
    def _on_index_failed(self, message: str) -> None:
        self._set_indexing_state(False)
        self._set_status(f"インデックス更新に失敗しました: {message}", error=True)

    @Slot()
    def _thread_finished(self) -> None:
        thread = self._thread
        self._worker = None
        self._thread = None
        if thread is not None:
            thread.deleteLater()
        if self._closing:
            self.close()

    @Slot()
    def _queue_search(self) -> None:
        self._search_generation += 1
        self._search_timer.start()

    @Slot()
    def _run_search(self) -> None:
        query = self.search_edit.text().strip()
        self.results_table.setRowCount(0)
        self.preview_edit.clear()
        self.preview_location_label.setText("プレビュー")
        if not query:
            return
        self._set_status("検索中…")
        selected_type = self.type_filter.currentData()
        document_types = {selected_type} if selected_type else None
        worker = SearchWorker(
            self._index,
            query,
            self._search_generation,
            document_types,
        )
        worker.signals.finished.connect(self._on_search_finished)
        worker.signals.failed.connect(self._on_search_failed)
        self._search_pool.start(worker)

    @Slot(int, str, object)
    def _on_search_finished(self, request_id: int, query: str, results: object) -> None:
        if request_id != self._search_generation or query != self.search_edit.text().strip():
            return
        results = list(results)
        self.results_table.setSortingEnabled(False)
        self.results_table.setRowCount(len(results))
        for row, result in enumerate(results):
            values = (
                result.filename,
                str(result.path),
                result.document_type.upper(),
                "、".join(result.locations),
                result.snippet,
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, result)
                if column == 1:
                    item.setToolTip(str(result.path))
                self.results_table.setItem(row, column, item)
        self.results_table.setSortingEnabled(True)
        self._set_status(f"検索結果: {len(results)}件。")

    @Slot(int, str, str)
    def _on_search_failed(self, request_id: int, query: str, message: str) -> None:
        if request_id == self._search_generation and query == self.search_edit.text().strip():
            self._set_status(f"検索に失敗しました: {message}", error=True)

    @Slot()
    def _show_selected_preview(self) -> None:
        result = self._selected_result()
        if result is None:
            self.preview_edit.clear()
            self.preview_location_label.setText("プレビュー")
            return
        self.preview_location_label.setText(
            f"プレビュー — {'、'.join(result.locations)}"
        )
        self.preview_edit.setPlainText(result.preview)
        selections: list[QTextEdit.ExtraSelection] = []
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QColor("#fff176"))
        for start, end in normalized_match_spans(result.preview, self.search_edit.text()):
            selection = QTextEdit.ExtraSelection()
            cursor = QTextCursor(self.preview_edit.document())
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            selection.cursor = cursor
            selection.format = highlight_format
            selections.append(selection)
        self.preview_edit.setExtraSelections(selections)

    @Slot(object)
    def _show_result_menu(self, position: object) -> None:
        item = self.results_table.itemAt(position)
        if item is None:
            return
        self.results_table.selectRow(item.row())
        menu = QMenu(self)
        menu.addAction("フルパスをコピー", self._copy_selected_path)
        menu.addAction("ファイルを開く", lambda: self._open_result(item.row(), 0))
        menu.exec(self.results_table.viewport().mapToGlobal(position))

    @Slot()
    def _copy_selected_path(self) -> None:
        result = self._selected_result()
        if result is not None:
            QApplication.clipboard().setText(str(result.path))
            self._set_status("フルパスをコピーしました。")

    @Slot(int, int)
    def _open_result(self, row: int, _column: int) -> None:
        result = self._result_at_row(row)
        if result is None:
            return
        if not result.path.exists():
            self._set_status("ファイルが見つかりません。", error=True)
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(result.path))):
            self._set_status("ファイルを開けませんでした。", error=True)

    def _selected_result(self) -> SearchResult | None:
        return self._result_at_row(self.results_table.currentRow())

    def _result_at_row(self, row: int) -> SearchResult | None:
        if row < 0:
            return None
        item = self.results_table.item(row, 0)
        if item is None:
            return None
        result = item.data(Qt.ItemDataRole.UserRole)
        return result if isinstance(result, SearchResult) else None

    def _set_indexing_state(self, indexing: bool) -> None:
        self.folder_button.setEnabled(not indexing)
        self.index_button.setEnabled(not indexing)
        self.cancel_button.setEnabled(indexing)

    def _set_status(self, message: str, *, error: bool = False) -> None:
        self.status_label.setText(message)
        self.status_label.setStyleSheet("color: #b00020;" if error else "")
        if not error:
            self.status_label.setToolTip("")

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._worker is not None:
            self._closing = True
            self._worker.request_cancel()
            event.ignore()
            return
        event.accept()


def _default_database_path() -> Path:
    local_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_data) if local_data else Path.home() / "AppData" / "Local"
    return base / "DocumentSearch" / "index.db"
