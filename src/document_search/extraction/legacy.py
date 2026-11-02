from __future__ import annotations

from pathlib import Path
from typing import Any

from .models import ExtractedChunk


MACROS_FORCE_DISABLED = 3


class LegacyOfficeUnavailable(RuntimeError):
    """Raised when safe Microsoft Office automation is unavailable."""


def read_doc(path: Path) -> list[ExtractedChunk]:
    pythoncom, client = _load_com()
    pythoncom.CoInitialize()
    application = document = None
    try:
        application = _start_application(client, "Word.Application")
        application.Visible = False
        application.DisplayAlerts = 0
        _disable_macros(application, "Word")
        try:
            application.Options.UpdateLinksAtOpen = False
        except Exception:
            pass
        document = application.Documents.Open(
            FileName=str(path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        text = str(document.Content.Text)
        if not text.strip():
            return []
        return [ExtractedChunk(text, {"kind": "document"})]
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                pass
            document = None
        if application is not None:
            try:
                application.Quit(SaveChanges=0)
            except Exception:
                pass
            application = None
        pythoncom.CoUninitialize()


def read_ppt(path: Path) -> list[ExtractedChunk]:
    pythoncom, client = _load_com()
    pythoncom.CoInitialize()
    application = presentation = None
    try:
        application = _start_application(client, "PowerPoint.Application")
        _disable_macros(application, "PowerPoint")
        try:
            application.DisplayAlerts = 1
        except Exception:
            pass
        try:
            application.Options.UpdateLinksAtOpen = False
        except Exception:
            pass
        presentation = application.Presentations.Open(
            str(path), ReadOnly=True, Untitled=False, WithWindow=False
        )
        chunks: list[ExtractedChunk] = []
        for slide_number in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides.Item(slide_number)
            parts: list[str] = []
            for shape_number in range(1, slide.Shapes.Count + 1):
                parts.extend(_shape_text(shape=slide.Shapes.Item(shape_number)))
            text = "\n".join(part for part in parts if part.strip())
            if text:
                chunks.append(
                    ExtractedChunk(text, {"kind": "slide", "slide": slide_number})
                )
        return chunks
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
            presentation = None
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
            application = None
        pythoncom.CoUninitialize()


def _load_com() -> tuple[Any, Any]:
    try:
        import pythoncom
        from win32com import client
    except ImportError as exc:
        raise LegacyOfficeUnavailable("Microsoft Office COM support is not installed") from exc
    return pythoncom, client


def _start_application(client: Any, program_id: str) -> Any:
    try:
        return client.DispatchEx(program_id)
    except Exception as exc:
        raise LegacyOfficeUnavailable(f"{program_id} is not available") from exc


def _disable_macros(application: Any, application_name: str) -> None:
    try:
        application.AutomationSecurity = MACROS_FORCE_DISABLED
        if application.AutomationSecurity != MACROS_FORCE_DISABLED:
            raise RuntimeError("Office did not accept the requested security level")
    except Exception as exc:
        raise LegacyOfficeUnavailable(
            f"{application_name} macro disabling is unavailable; file was not opened"
        ) from exc


def _shape_text(shape: Any) -> list[str]:
    parts: list[str] = []
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            parts.append(str(shape.TextFrame.TextRange.Text))
    except Exception:
        pass
    try:
        if shape.HasTable:
            table = shape.Table
            for row in range(1, table.Rows.Count + 1):
                cells = []
                for column in range(1, table.Columns.Count + 1):
                    cells.append(str(table.Cell(row, column).Shape.TextFrame.TextRange.Text))
                parts.append("\t".join(cells))
    except Exception:
        pass
    try:
        group_items = shape.GroupItems
        for item_number in range(1, group_items.Count + 1):
            parts.extend(_shape_text(group_items.Item(item_number)))
    except Exception:
        pass
    return parts
