from __future__ import annotations

from pathlib import Path
from threading import Event
from collections.abc import Collection

from PySide6.QtCore import QObject, QRunnable, Signal, Slot

from document_search.indexing import DocumentIndex


class IndexWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, index: DocumentIndex, root: Path) -> None:
        super().__init__()
        self._index = index
        self._root = root
        self._cancel_requested = Event()

    def request_cancel(self) -> None:
        self._cancel_requested.set()

    @Slot()
    def run(self) -> None:
        try:
            report = self._index.index_directory(
                self._root,
                progress_callback=self._report_progress,
                should_cancel=self._cancel_requested.is_set,
            )
        except Exception as exc:
            self.failed.emit(f"{type(exc).__name__}: {exc}")
            return
        self.finished.emit(report)

    def _report_progress(self, current: int, total: int, path: Path) -> None:
        self.progress.emit(current, total, str(path))


class SearchSignals(QObject):
    finished = Signal(int, str, object)
    failed = Signal(int, str, str)


class SearchWorker(QRunnable):
    def __init__(
        self,
        index: DocumentIndex,
        query: str,
        request_id: int,
        document_types: Collection[str] | None = None,
    ) -> None:
        super().__init__()
        self._index = index
        self._query = query
        self._request_id = request_id
        self._document_types = document_types
        self.signals = SearchSignals()

    @Slot()
    def run(self) -> None:
        try:
            results = self._index.search(
                self._query, document_types=self._document_types
            )
        except Exception as exc:
            self.signals.failed.emit(
                self._request_id, self._query, f"{type(exc).__name__}: {exc}"
            )
            return
        self.signals.finished.emit(self._request_id, self._query, results)
