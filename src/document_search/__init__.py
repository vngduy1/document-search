"""Offline document search."""

from .extraction import ExtractedChunk, ExtractionResult, extract_document
from .indexing import DocumentIndex, IndexingReport, SearchResult

__all__ = [
    "DocumentIndex",
    "ExtractedChunk",
    "ExtractionResult",
    "IndexingReport",
    "SearchResult",
    "extract_document",
]
