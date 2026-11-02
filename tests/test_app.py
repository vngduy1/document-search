from document_search.gui.app import _sqlite_self_test


def test_sqlite_self_test_verifies_fts5_trigram() -> None:
    assert _sqlite_self_test() == 0
