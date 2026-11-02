from __future__ import annotations

from bisect import bisect_left, bisect_right
import unicodedata


def normalized_match_spans(text: str, query: str) -> list[tuple[int, int]]:
    normalized_query = _normalize(query)
    if not text or not normalized_query:
        return []

    normalized_text = _normalize(text)
    boundaries = [len(_normalize(text[:index])) for index in range(len(text) + 1)]
    spans: list[tuple[int, int]] = []
    search_from = 0
    while True:
        position = normalized_text.find(normalized_query, search_from)
        if position < 0:
            break
        normalized_end = position + len(normalized_query)
        start = max(0, bisect_right(boundaries, position) - 1)
        end = bisect_left(boundaries, normalized_end)
        if end < len(boundaries) and boundaries[end] == normalized_end:
            end = bisect_right(boundaries, normalized_end) - 1
        end = max(start + 1, min(end, len(text)))
        spans.append((start, end))
        search_from = normalized_end
    return spans


def _normalize(value: str) -> str:
    return unicodedata.normalize("NFKC", value).casefold()
