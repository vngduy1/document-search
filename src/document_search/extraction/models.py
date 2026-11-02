from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal


ExtractionStatus = Literal["ok", "unsupported", "error"]


@dataclass(frozen=True, slots=True)
class ExtractedChunk:
    text: str
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class ExtractionResult:
    path: Path
    document_type: str
    status: ExtractionStatus
    chunks: tuple[ExtractedChunk, ...] = ()
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.status == "ok"
