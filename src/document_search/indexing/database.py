from __future__ import annotations

from datetime import datetime, timezone
from collections.abc import Callable, Collection
from contextlib import closing
import json
import os
from pathlib import Path
import sqlite3
import unicodedata

from document_search.extraction import ExtractionResult, extract_document

from .models import IndexingReport, MatchLocation, SearchResult
from .scanner import discover_documents


SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY,
    path TEXT NOT NULL UNIQUE COLLATE NOCASE,
    root_path TEXT NOT NULL COLLATE NOCASE,
    filename TEXT NOT NULL,
    filename_normalized TEXT NOT NULL,
    document_type TEXT NOT NULL,
    size INTEGER NOT NULL,
    mtime_ns INTEGER NOT NULL,
    extraction_status TEXT NOT NULL,
    extraction_error TEXT,
    indexed_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS documents_root_path_idx ON documents(root_path);

CREATE TABLE IF NOT EXISTS chunks (
    id INTEGER PRIMARY KEY,
    document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    position INTEGER NOT NULL,
    text TEXT NOT NULL,
    text_normalized TEXT NOT NULL,
    metadata_json TEXT NOT NULL,
    UNIQUE(document_id, position)
);

CREATE INDEX IF NOT EXISTS chunks_document_id_idx ON chunks(document_id);

CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
    filename,
    content,
    tokenize='trigram'
);
"""


class DocumentIndex:
    def __init__(self, database_path: str | Path) -> None:
        self.database_path = Path(database_path).resolve()
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as connection:
            connection.executescript(SCHEMA)
            connection.commit()

    def index_directory(
        self,
        root: str | Path,
        *,
        progress_callback: Callable[[int, int, Path], None] | None = None,
        should_cancel: Callable[[], bool] | None = None,
    ) -> IndexingReport:
        root_path = Path(root).resolve()
        if not root_path.is_dir():
            return IndexingReport(0, 0, 0, 1, (f"Not a directory: {root_path}",))

        paths, discovery_issues, cancelled = discover_documents(root_path, should_cancel)
        issues = list(discovery_issues)
        indexed = unchanged = failed = 0
        if cancelled:
            return IndexingReport(0, 0, 0, 0, tuple(issues), True)
        if progress_callback:
            progress_callback(0, len(paths), root_path)

        with closing(self._connect()) as connection, connection:
            existing_rows = connection.execute(
                "SELECT id, path, size, mtime_ns FROM documents WHERE root_path = ? COLLATE NOCASE",
                (str(root_path),),
            ).fetchall()
            existing = {os.path.normcase(row["path"]): row for row in existing_rows}
            seen: set[str] = set()

            for current, path in enumerate(paths, start=1):
                if should_cancel and should_cancel():
                    cancelled = True
                    break
                path_text = str(path)
                path_key = os.path.normcase(path_text)
                seen.add(path_key)
                try:
                    state = path.stat()
                except OSError as exc:
                    failed += 1
                    issues.append(f"{path}: {exc}")
                    if progress_callback:
                        progress_callback(current, len(paths), path)
                    continue

                old = existing.get(path_key)
                if old and old["size"] == state.st_size and old["mtime_ns"] == state.st_mtime_ns:
                    unchanged += 1
                    if progress_callback:
                        progress_callback(current, len(paths), path)
                    continue

                result = extract_document(path)
                self._store_document(connection, root_path, state, result)
                indexed += 1
                if result.status == "error":
                    failed += 1
                    issues.append(f"{path}: {result.error}")
                if progress_callback:
                    progress_callback(current, len(paths), path)

            if should_cancel and should_cancel():
                cancelled = True
            deleted_ids = (
                [] if cancelled else [row["id"] for key, row in existing.items() if key not in seen]
            )
            for document_id in deleted_ids:
                connection.execute("DELETE FROM documents_fts WHERE rowid = ?", (document_id,))
                connection.execute("DELETE FROM documents WHERE id = ?", (document_id,))

        return IndexingReport(
            indexed, unchanged, len(deleted_ids), failed, tuple(issues), cancelled
        )

    def search(
        self,
        query: str,
        *,
        limit: int = 100,
        document_types: Collection[str] | None = None,
    ) -> list[SearchResult]:
        normalized_query = _normalize(query)
        if not normalized_query or limit <= 0:
            return []
        normalized_types = None
        if document_types is not None:
            normalized_types = sorted(
                {value.casefold().removeprefix(".") for value in document_types if value}
            )
            if not normalized_types:
                return []

        exact_filter = """
            instr(d.filename_normalized, :query) > 0
            OR EXISTS (
                SELECT 1 FROM chunks c
                WHERE c.document_id = d.id
                AND instr(c.text_normalized, :query) > 0
            )
        """
        if len(normalized_query) >= 3:
            source = "documents d JOIN documents_fts f ON f.rowid = d.id"
            candidate_filter = "documents_fts MATCH :fts_query AND"
        else:
            source = "documents d"
            candidate_filter = ""

        rows = self._search_rows(
            source,
            candidate_filter,
            exact_filter,
            normalized_query,
            limit,
            normalized_types,
        )
        return [self._make_search_result(row, normalized_query) for row in rows]

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def _store_document(
        self,
        connection: sqlite3.Connection,
        root_path: Path,
        state: os.stat_result,
        result: ExtractionResult,
    ) -> None:
        path_text = str(result.path.resolve())
        connection.execute(
            """
            INSERT INTO documents (
                path, root_path, filename, filename_normalized, document_type,
                size, mtime_ns, extraction_status, extraction_error, indexed_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(path) DO UPDATE SET
                root_path = excluded.root_path,
                filename = excluded.filename,
                filename_normalized = excluded.filename_normalized,
                document_type = excluded.document_type,
                size = excluded.size,
                mtime_ns = excluded.mtime_ns,
                extraction_status = excluded.extraction_status,
                extraction_error = excluded.extraction_error,
                indexed_at = excluded.indexed_at
            """,
            (
                path_text,
                str(root_path),
                result.path.name,
                _normalize(result.path.name),
                result.document_type,
                state.st_size,
                state.st_mtime_ns,
                result.status,
                result.error,
                datetime.now(timezone.utc).isoformat(),
            ),
        )
        document_id = connection.execute(
            "SELECT id FROM documents WHERE path = ? COLLATE NOCASE", (path_text,)
        ).fetchone()["id"]

        connection.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
        connection.execute("DELETE FROM documents_fts WHERE rowid = ?", (document_id,))

        for position, chunk in enumerate(result.chunks):
            connection.execute(
                """
                INSERT INTO chunks (
                    document_id, position, text, text_normalized, metadata_json
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    document_id,
                    position,
                    chunk.text,
                    _normalize(chunk.text),
                    json.dumps(chunk.metadata, ensure_ascii=False, sort_keys=True),
                ),
            )

        content = "\n".join(_normalize(chunk.text) for chunk in result.chunks)
        connection.execute(
            "INSERT INTO documents_fts(rowid, filename, content) VALUES (?, ?, ?)",
            (document_id, _normalize(result.path.name), content),
        )

    def _search_rows(
        self,
        source: str,
        candidate_filter: str,
        exact_filter: str,
        normalized_query: str,
        limit: int,
        document_types: list[str] | None,
    ) -> list[sqlite3.Row]:
        type_filter = ""
        if document_types:
            placeholders = []
            for index in range(len(document_types)):
                placeholders.append(f":type_{index}")
            type_filter = f"AND d.document_type IN ({', '.join(placeholders)})"
        sql = f"""
            SELECT
                d.id,
                d.path,
                d.filename,
                d.document_type,
                d.extraction_status,
                instr(d.filename_normalized, :query) > 0 AS filename_match,
                EXISTS (
                    SELECT 1 FROM chunks c
                    WHERE c.document_id = d.id
                    AND instr(c.text_normalized, :query) > 0
                ) AS content_match,
                COALESCE((
                    SELECT c.text FROM chunks c
                    WHERE c.document_id = d.id
                    AND instr(c.text_normalized, :query) > 0
                    ORDER BY c.position
                    LIMIT 1
                ), '') AS matched_text,
                COALESCE((
                    SELECT group_concat(m.metadata_json, char(30))
                    FROM (
                        SELECT c.metadata_json
                        FROM chunks c
                        WHERE c.document_id = d.id
                        AND instr(c.text_normalized, :query) > 0
                        ORDER BY c.position
                        LIMIT 10
                    ) AS m
                ), '') AS matched_metadata
            FROM {source}
            WHERE {candidate_filter} ({exact_filter}) {type_filter}
            ORDER BY filename_match DESC, d.filename COLLATE NOCASE, d.path COLLATE NOCASE
            LIMIT :limit
        """
        parameters = {
            "query": normalized_query,
            "fts_query": _quote_fts_phrase(normalized_query),
            "limit": limit,
        }
        for index, document_type in enumerate(document_types or []):
            parameters[f"type_{index}"] = document_type
        with closing(self._connect()) as connection:
            return connection.execute(sql, parameters).fetchall()

    @staticmethod
    def _make_search_result(row: sqlite3.Row, normalized_query: str) -> SearchResult:
        matched_in: list[MatchLocation] = []
        if row["filename_match"]:
            matched_in.append("filename")
        if row["content_match"]:
            matched_in.append("content")
        matched_text = row["matched_text"]
        return SearchResult(
            path=Path(row["path"]),
            filename=row["filename"],
            document_type=row["document_type"],
            status=row["extraction_status"],
            matched_in=tuple(matched_in),
            snippet=_make_snippet(matched_text, normalized_query) or row["filename"],
            preview=_make_preview(matched_text, normalized_query) or row["filename"],
            locations=_make_locations(row["matched_metadata"], tuple(matched_in)),
        )


def _normalize(value: str) -> str:
    return unicodedata.normalize("NFKC", value).casefold()


def _quote_fts_phrase(value: str) -> str:
    return f'"{value.replace(chr(34), chr(34) * 2)}"'


def _make_snippet(text: str, normalized_query: str, *, width: int = 100) -> str:
    compact = " ".join(text.split())
    if not compact:
        return ""
    position = _normalize(compact).find(normalized_query)
    if position < 0:
        return compact[:width]
    start = max(0, position - width // 3)
    end = min(len(compact), start + width)
    prefix = "…" if start else ""
    suffix = "…" if end < len(compact) else ""
    return f"{prefix}{compact[start:end]}{suffix}"


def _make_preview(text: str, normalized_query: str, *, width: int = 4000) -> str:
    if len(text) <= width:
        return text
    position = _normalize(text).find(normalized_query)
    start = max(0, position - width // 3) if position >= 0 else 0
    end = min(len(text), start + width)
    prefix = "…\n" if start else ""
    suffix = "\n…" if end < len(text) else ""
    return f"{prefix}{text[start:end]}{suffix}"


def _make_locations(serialized: str, matched_in: tuple[MatchLocation, ...]) -> tuple[str, ...]:
    locations: list[str] = []
    if "filename" in matched_in:
        locations.append("ファイル名")
    for value in filter(None, serialized.split(chr(30))):
        try:
            metadata = json.loads(value)
        except (TypeError, json.JSONDecodeError):
            continue
        location = _format_location(metadata)
        if location not in locations:
            locations.append(location)
    if "content" in matched_in and not any(location != "ファイル名" for location in locations):
        locations.append("本文")
    return tuple(locations)


def _format_location(metadata: dict[str, object]) -> str:
    kind = metadata.get("kind")
    if kind == "sheet_row":
        return f"シート「{metadata.get('sheet', '')}」 / 行 {metadata.get('row', '')}"
    if kind == "slide":
        return f"スライド {metadata.get('slide', '')}"
    if kind == "page":
        return f"ページ {metadata.get('page', '')}"
    if kind == "paragraph":
        return f"段落 {metadata.get('paragraph', '')}"
    if kind == "table_row":
        return f"表 {metadata.get('table', '')} / 行 {metadata.get('row', '')}"
    return "本文"
