import pytest

from word_document_server.tools import live_v2_tools


def setup_function():
    live_v2_tools._sessions.clear()
    live_v2_tools._handles.clear()


def test_register_session_and_resolve_filename():
    session_id = live_v2_tools._register_session({
        "document": "Contract.docx",
        "full_path": r"C:\Docs\Contract.docx",
    })

    assert session_id.startswith("word_")
    assert live_v2_tools._resolve_filename(session_id=session_id) == "Contract.docx"
    assert live_v2_tools._resolve_filename(session_id=session_id, filename="Other.docx") == "Other.docx"


def test_store_handle_creates_selection_target():
    session_id = live_v2_tools._register_session({"document": "Contract.docx"})
    match = {"start": 10, "end": 18, "text": "ACME Corp", "context": "Hello ACME Corp"}

    enriched = live_v2_tools._store_handle(session_id, match, 0)

    assert enriched["handle"] == "match_1"
    assert enriched["target"] == {"kind": "selection", "start": 10, "end": 18}
    assert live_v2_tools._resolve_target(session_id=session_id, handle="match_1") == enriched["target"]


def test_resolve_target_validates_unknown_handle():
    session_id = live_v2_tools._register_session({"document": "Contract.docx"})

    with pytest.raises(ValueError, match="Unknown handle"):
        live_v2_tools._resolve_target(session_id=session_id, handle="match_404")


def test_target_range_requires_selection_start_and_end():
    assert live_v2_tools._target_range({"kind": "selection", "start": "1", "end": "5"}) == (1, 5)

    with pytest.raises(ValueError, match="target.kind"):
        live_v2_tools._target_range({"kind": "block", "start": 1, "end": 5})

    with pytest.raises(ValueError, match="requires start and end"):
        live_v2_tools._target_range({"kind": "selection", "start": 1})
