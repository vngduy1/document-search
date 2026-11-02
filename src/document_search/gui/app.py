from __future__ import annotations

import sys
import sqlite3
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from .main_window import DocumentSearchWindow


def main(arguments: list[str] | None = None) -> int:
    arguments = list(sys.argv if arguments is None else arguments)
    if "--self-test" in arguments:
        return _sqlite_self_test()

    smoke_test = "--smoke-test-gui" in arguments
    application = QApplication([arguments[0]])
    temporary_directory = TemporaryDirectory(prefix="document-search-") if smoke_test else None
    database_path = (
        Path(temporary_directory.name) / "smoke-test.db" if temporary_directory else None
    )
    window = DocumentSearchWindow(database_path)
    window.show()
    if smoke_test:
        QTimer.singleShot(750, application.quit)
    exit_code = application.exec()
    if temporary_directory:
        temporary_directory.cleanup()
    return exit_code


def _sqlite_self_test() -> int:
    if sys.platform == "win32":
        import pythoncom  # noqa: F401
        import pywintypes  # noqa: F401
        from win32com import client  # noqa: F401

    connection = sqlite3.connect(":memory:")
    try:
        connection.execute("CREATE VIRTUAL TABLE test_fts USING fts5(text, tokenize='trigram')")
        connection.execute("INSERT INTO test_fts(text) VALUES (?)", ("日本語検索",))
        match = connection.execute(
            "SELECT count(*) FROM test_fts WHERE test_fts MATCH ?", ("日本語",)
        ).fetchone()[0]
    finally:
        connection.close()
    return 0 if match == 1 else 1
