from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal


MatchLocation = Literal["filename", "content"]


@dataclass(frozen=True, slots=True)
class IndexingReport:
    indexed: int
    unchanged: int
    deleted: int
    failed: int
    issues: tuple[str, ...] = ()
    cancelled: bool = False


@dataclass(frozen=True, slots=True)
class SearchResult:
    path: Path
    filename: str
    document_type: str
    status: str
    matched_in: tuple[MatchLocation, ...]
    snippet: str
    preview: str
    locations: tuple[str, ...]
