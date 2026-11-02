from __future__ import annotations

import os
from collections.abc import Callable
from pathlib import Path


DOCUMENT_EXTENSIONS = frozenset(
    {".docx", ".xlsx", ".xls", ".pptx", ".pdf", ".txt", ".doc", ".ppt"}
)
SKIPPED_DIRECTORIES = frozenset(
    {".git", ".mypy_cache", ".pytest_cache", ".pytest-tmp", ".venv", "__pycache__", "build", "dist"}
)


def discover_documents(
    root: Path, should_cancel: Callable[[], bool] | None = None
) -> tuple[list[Path], list[str], bool]:
    documents: list[Path] = []
    issues: list[str] = []

    def record_walk_error(error: OSError) -> None:
        issues.append(f"{error.filename or root}: {error}")

    for directory, subdirectories, filenames in os.walk(
        root, topdown=True, onerror=record_walk_error, followlinks=False
    ):
        if should_cancel and should_cancel():
            return documents, issues, True
        subdirectories[:] = [
            name for name in subdirectories if name.casefold() not in SKIPPED_DIRECTORIES
        ]
        directory_path = Path(directory)
        for filename in filenames:
            if should_cancel and should_cancel():
                return documents, issues, True
            path = directory_path / filename
            if path.suffix.casefold() in DOCUMENT_EXTENSIONS:
                documents.append(path.resolve())

    documents.sort(key=lambda path: os.path.normcase(str(path)))
    return documents, issues, False
