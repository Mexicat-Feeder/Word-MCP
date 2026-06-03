import pytest
import json
import asyncio

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


def test_save_requires_output_path_for_unsaved_sessions():
    session_id = live_v2_tools._register_session({"document": "Document1", "full_path": ""})

    result = json.loads(asyncio.run(live_v2_tools.word_v2_save(session_id)))

    assert "unsaved document" in result["error"]


def test_open_without_path_attaches_active_document(monkeypatch):
    async def fake_list_open():
        return json.dumps({
            "success": True,
            "count": 2,
            "documents": [
                {"index": 1, "name": "Background.docx", "full_path": r"C:\Docs\Background.docx", "active": False},
                {"index": 2, "name": "Source.docx", "full_path": r"C:\Docs\Source.docx", "active": True, "pages": 3, "saved": True},
            ],
        })

    async def fail_create_document(**_kwargs):
        raise AssertionError("word_v2_open() must not create a blank document when an active document exists")

    monkeypatch.setattr(live_v2_tools.live_read_tools, "word_live_list_open", fake_list_open)
    monkeypatch.setattr(live_v2_tools.live_tools, "word_live_create_document", fail_create_document)

    result = json.loads(asyncio.run(live_v2_tools.word_v2_open()))

    assert result["success"] is True
    assert result["document"] == "Source.docx"
    assert result["full_path"] == r"C:\Docs\Source.docx"
    assert result["session_id"].startswith("word_")
    assert live_v2_tools._resolve_filename(result["session_id"]) == "Source.docx"


def test_open_action_list_returns_documents_without_session(monkeypatch):
    async def fake_list_open():
        return json.dumps({
            "success": True,
            "count": 1,
            "documents": [
                {"index": 1, "name": "Source.docx", "full_path": r"C:\Docs\Source.docx", "active": True},
            ],
        })

    monkeypatch.setattr(live_v2_tools.live_read_tools, "word_live_list_open", fake_list_open)

    result = json.loads(asyncio.run(live_v2_tools.word_v2_open(action="list")))

    assert result["success"] is True
    assert result["count"] == 1
    assert result["documents"][0]["active"] is True
    assert "session_id" not in result
    assert "word_v2_open()" in result["usage"]


def test_open_action_attach_uses_listed_document_identity(monkeypatch):
    async def fake_list_open():
        return json.dumps({
            "success": True,
            "count": 2,
            "documents": [
                {"index": 1, "name": "Background.docx", "full_path": r"C:\Docs\Background.docx", "active": False},
                {"index": 2, "name": "Source.docx", "full_path": r"C:\Docs\Source.docx", "active": True},
            ],
        })

    monkeypatch.setattr(live_v2_tools.live_read_tools, "word_live_list_open", fake_list_open)

    result = json.loads(asyncio.run(live_v2_tools.word_v2_open(action="attach", path="1")))

    assert result["success"] is True
    assert result["document"] == "Background.docx"
    assert live_v2_tools._sessions[result["session_id"]]["full_path"] == r"C:\Docs\Background.docx"


def test_open_action_new_explicitly_creates_blank_document(monkeypatch):
    async def fake_create_document(visible=True):
        return json.dumps({
            "success": True,
            "document": "Document1",
            "full_path": "",
            "message": "New blank document 'Document1' created successfully in live session",
            "visible": visible,
        })

    def fake_apply_template(filename):
        return {"success": True, "template": "default_plain", "document": filename}

    monkeypatch.setattr(live_v2_tools.live_tools, "word_live_create_document", fake_create_document)
    monkeypatch.setattr(live_v2_tools, "_apply_default_template_live", fake_apply_template)

    result = json.loads(asyncio.run(live_v2_tools.word_v2_open(action="new", visible=False)))

    assert result["success"] is True
    assert result["document"] == "Document1"
    assert result["template"] == "default_plain"
    assert result["template_result"]["success"] is True
    assert result["session_id"].startswith("word_")


def test_blueprint_validation_rejects_unknown_blocks_and_bad_tables():
    errors = live_v2_tools._validate_blueprint({
        "blocks": [
            {"type": "callout", "text": "Nope"},
            {"type": "table", "rows": [["A"], ["B", "C"]]},
        ],
    })

    assert any("unsupported type" in error for error in errors)
    assert any("same non-zero width" in error for error in errors)


def test_blueprint_expected_counts():
    counts = live_v2_tools._blueprint_expected_counts({
        "blocks": [
            {"type": "heading", "text": "Title", "level": 1},
            {"type": "paragraph", "text": "Body"},
            {"type": "list", "items": ["One", "Two"]},
            {"type": "table", "rows": [["A", "B"]]},
            {"type": "image_placeholder", "label": "screen"},
            {"type": "page_break"},
        ],
    })

    assert counts == {
        "paragraphs": 5,
        "tables": 1,
        "image_placeholders": 1,
        "blocks": 6,
    }


def test_layout_invalid_action_does_not_require_session():
    result = json.loads(asyncio.run(live_v2_tools.word_v2_layout(session_id="missing", action="bad")))

    assert result["error"] == "Invalid action"
    assert "page_setup" in result["valid_actions"]


def test_blueprint_create_dispatches_blocks_in_order(monkeypatch):
    calls = []

    async def fake_open(action="open", visible=True, **_kwargs):
        calls.append(("open", action, visible))
        session_id = live_v2_tools._register_session({
            "document": "Document1",
            "full_path": "",
            "template": "default_plain",
        })
        return json.dumps({"success": True, "document": "Document1", "session_id": session_id})

    async def fake_layout(**kwargs):
        calls.append(("layout", kwargs["action"]))
        return json.dumps({"success": True, "session_id": kwargs["session_id"], "action": kwargs["action"]})

    async def fake_edit(**kwargs):
        calls.append(("edit", kwargs["action"], tuple(kwargs.get("paragraphs") or []), kwargs.get("style")))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    async def fake_table(**kwargs):
        calls.append(("table", kwargs["action"], kwargs["rows"], kwargs["cols"]))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    async def fake_get_content(session_id, action="info", **_kwargs):
        return json.dumps({"success": True, "session_id": session_id, "paragraphs": 10, "tables": 1})

    monkeypatch.setattr(live_v2_tools, "word_v2_open", fake_open)
    monkeypatch.setattr(live_v2_tools, "word_v2_layout", fake_layout)
    monkeypatch.setattr(live_v2_tools, "word_v2_edit", fake_edit)
    monkeypatch.setattr(live_v2_tools, "word_v2_table", fake_table)
    monkeypatch.setattr(live_v2_tools, "word_v2_get_content", fake_get_content)

    result = json.loads(asyncio.run(live_v2_tools.word_v2_blueprint(
        action="create",
        blueprint={
            "page_setup": {"size": "letter", "orientation": "portrait"},
            "properties": {"title": "Demo"},
            "blocks": [
                {"type": "heading", "text": "Title", "level": 1},
                {"type": "paragraph", "text": "Body"},
                {"type": "table", "rows": [["A", "B"], ["1", "2"]]},
                {"type": "page_break"},
                {"type": "image_placeholder", "label": "Chart"},
            ],
        },
    )))

    assert result["success"] is True
    assert calls[:5] == [
        ("open", "new", True),
        ("layout", "page_setup"),
        ("layout", "properties"),
        ("edit", "insert_paragraphs", ("Title",), "Heading 1"),
        ("edit", "insert_paragraphs", ("Body",), "Normal"),
    ]
    assert ("table", "create", 2, 2) in calls
    assert ("layout", "page_break") in calls
    assert result["warnings"][0]["warning"].startswith("image_placeholder")
