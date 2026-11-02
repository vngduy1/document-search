from __future__ import annotations

import os
from pathlib import Path

import pytest

from document_search.extraction import extract_document
from document_search.extraction import legacy


class _FakePythonCom:
    def __init__(self) -> None:
        self.initialized = False

    def CoInitialize(self) -> None:
        self.initialized = True

    def CoUninitialize(self) -> None:
        self.initialized = False


class _FakeDocument:
    def __init__(self) -> None:
        self.Content = type("Content", (), {"Text": "旧式の日本語文書"})()
        self.close_arguments = None

    def Close(self, **kwargs) -> None:
        self.close_arguments = kwargs


class _FakeDocuments:
    def __init__(self, application) -> None:
        self.application = application
        self.open_arguments = None
        self.document = _FakeDocument()

    def Open(self, **kwargs):
        assert self.application.AutomationSecurity == legacy.MACROS_FORCE_DISABLED
        self.open_arguments = kwargs
        return self.document


class _FakeWord:
    def __init__(self) -> None:
        self.AutomationSecurity = 0
        self.Documents = _FakeDocuments(self)
        self.quit_arguments = None

    def Quit(self, **kwargs) -> None:
        self.quit_arguments = kwargs


class _FakeClient:
    def __init__(self, application) -> None:
        self.application = application

    def DispatchEx(self, program_id: str):
        assert program_id == "Word.Application"
        return self.application


def test_word_com_forces_macro_disable_and_read_only(monkeypatch, tmp_path: Path) -> None:
    pythoncom = _FakePythonCom()
    application = _FakeWord()
    monkeypatch.setattr(legacy, "_load_com", lambda: (pythoncom, _FakeClient(application)))

    chunks = legacy.read_doc(tmp_path / "旧式.doc")

    assert chunks[0].text == "旧式の日本語文書"
    assert application.AutomationSecurity == legacy.MACROS_FORCE_DISABLED
    assert application.Documents.open_arguments["ReadOnly"] is True
    assert application.Documents.open_arguments["AddToRecentFiles"] is False
    assert application.Documents.document.close_arguments == {"SaveChanges": 0}
    assert application.quit_arguments == {"SaveChanges": 0}
    assert not pythoncom.initialized


def test_legacy_extraction_is_unsupported_when_safe_com_is_unavailable(
    monkeypatch, tmp_path: Path
) -> None:
    path = tmp_path / "旧式.doc"
    path.write_bytes(b"legacy")

    def unavailable():
        raise legacy.LegacyOfficeUnavailable("Office is unavailable")

    monkeypatch.setattr(legacy, "_load_com", unavailable)
    result = extract_document(path)

    assert result.status == "unsupported"
    assert result.error == "Office is unavailable"


def test_legacy_file_is_not_opened_when_macro_security_cannot_be_forced(
    monkeypatch, tmp_path: Path
) -> None:
    class RejectingDocuments:
        def Open(self, **_kwargs):
            raise AssertionError("file must not be opened")

    class RejectingWord:
        Visible = False
        DisplayAlerts = 0
        Documents = RejectingDocuments()

        @property
        def AutomationSecurity(self):
            return 0

        @AutomationSecurity.setter
        def AutomationSecurity(self, _value):
            pass

        def Quit(self, **_kwargs):
            pass

    pythoncom = _FakePythonCom()
    application = RejectingWord()
    monkeypatch.setattr(legacy, "_load_com", lambda: (pythoncom, _FakeClient(application)))

    with pytest.raises(legacy.LegacyOfficeUnavailable, match="macro disabling"):
        legacy.read_doc(tmp_path / "unsafe.doc")

    assert not pythoncom.initialized


@pytest.mark.office_com
@pytest.mark.skipif(
    os.environ.get("RUN_OFFICE_COM_TESTS") != "1",
    reason="set RUN_OFFICE_COM_TESTS=1 to run installed-Office integration",
)
def test_installed_office_extracts_legacy_doc_and_ppt_read_only(tmp_path: Path) -> None:
    import pythoncom
    from win32com import client

    doc_path = tmp_path / "旧式日本語.doc"
    ppt_path = tmp_path / "旧式日本語.ppt"
    pythoncom.CoInitialize()
    word = document = powerpoint = presentation = None
    try:
        word = client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = legacy.MACROS_FORCE_DISABLED
        document = word.Documents.Add()
        document.Content.Text = "旧式Wordの日本語"
        document.SaveAs2(str(doc_path), FileFormat=0, AddToRecentFiles=False)
        document.Close(SaveChanges=0)
        document = None
        word.Quit(SaveChanges=0)
        word = None

        powerpoint = client.DispatchEx("PowerPoint.Application")
        powerpoint.AutomationSecurity = legacy.MACROS_FORCE_DISABLED
        powerpoint.DisplayAlerts = 1
        presentation = powerpoint.Presentations.Add(WithWindow=False)
        slide = presentation.Slides.Add(1, 12)
        text_box = slide.Shapes.AddTextbox(1, 20, 20, 500, 60)
        text_box.TextFrame.TextRange.Text = "旧式PowerPointの日本語"
        presentation.SaveAs(str(ppt_path), 1)
        text_box = None
        slide = None
        presentation.Close()
        presentation = None
        powerpoint.Quit()
        powerpoint = None
    finally:
        if document is not None:
            document.Close(SaveChanges=0)
        if word is not None:
            word.Quit(SaveChanges=0)
        if presentation is not None:
            presentation.Close()
        if powerpoint is not None:
            powerpoint.Quit()
        pythoncom.CoUninitialize()

    original_doc = (doc_path.read_bytes(), doc_path.stat().st_mtime_ns)
    original_ppt = (ppt_path.read_bytes(), ppt_path.stat().st_mtime_ns)

    doc_result = extract_document(doc_path)
    ppt_result = extract_document(ppt_path)

    assert doc_result.ok
    assert "旧式Wordの日本語" in doc_result.chunks[0].text
    assert ppt_result.ok
    assert "旧式PowerPointの日本語" in ppt_result.chunks[0].text
    assert ppt_result.chunks[0].metadata == {"kind": "slide", "slide": 1}
    assert (doc_path.read_bytes(), doc_path.stat().st_mtime_ns) == original_doc
    assert (ppt_path.read_bytes(), ppt_path.stat().st_mtime_ns) == original_ppt
