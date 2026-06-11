import pytest
import json
import asyncio

from word_document_server.tools import live_api_tools


def setup_function():
    live_api_tools._sessions.clear()
    live_api_tools._handles.clear()


def test_register_session_and_resolve_filename():
    session_id = live_api_tools._register_session({
        "document": "Contract.docx",
        "full_path": r"C:\Docs\Contract.docx",
    })

    assert session_id.startswith("word_")
    assert live_api_tools._resolve_filename(session_id=session_id) == r"C:\Docs\Contract.docx"
    assert live_api_tools._resolve_filename(session_id=session_id, filename="Other.docx") == "Other.docx"


def test_resolve_filename_falls_back_to_document_name_for_unsaved_sessions():
    session_id = live_api_tools._register_session({
        "document": "Document1",
        "full_path": "",
    })

    assert live_api_tools._resolve_filename(session_id=session_id) == "Document1"


def test_close_save_flag_accepts_no_without_prompting():
    assert live_api_tools.live_tools._word_close_save_flag("no") == 0
    assert live_api_tools.live_tools._word_close_save_flag("discard") == 0
    assert live_api_tools.live_tools._word_close_save_flag("save") == -2
    assert live_api_tools.live_tools._word_close_save_flag("prompt") == -1

    with pytest.raises(ValueError, match="Unknown save_changes"):
        live_api_tools.live_tools._word_close_save_flag("maybe")


def test_store_handle_creates_selection_target():
    session_id = live_api_tools._register_session({"document": "Contract.docx"})
    match = {"start": 10, "end": 18, "text": "ACME Corp", "context": "Hello ACME Corp"}

    enriched = live_api_tools._store_handle(session_id, match, 0)

    assert enriched["handle"] == "match_1"
    assert enriched["target"] == {"kind": "selection", "start": 10, "end": 18}
    assert live_api_tools._resolve_target(session_id=session_id, handle="match_1") == enriched["target"]


def test_resolve_target_validates_unknown_handle():
    session_id = live_api_tools._register_session({"document": "Contract.docx"})

    with pytest.raises(ValueError, match="Unknown handle"):
        live_api_tools._resolve_target(session_id=session_id, handle="match_404")


def test_target_range_requires_selection_start_and_end():
    assert live_api_tools._target_range({"kind": "selection", "start": "1", "end": "5"}) == (1, 5)

    with pytest.raises(ValueError, match="target.kind"):
        live_api_tools._target_range({"kind": "block", "start": 1, "end": 5})

    with pytest.raises(ValueError, match="requires start and end"):
        live_api_tools._target_range({"kind": "selection", "start": 1})


def test_save_requires_output_path_for_unsaved_sessions():
    session_id = live_api_tools._register_session({"document": "Document1", "full_path": ""})

    result = json.loads(asyncio.run(live_api_tools.word_save(session_id)))

    assert "unsaved document" in result["error"]


def test_open_without_path_creates_new_document(monkeypatch):
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

    monkeypatch.setattr(live_api_tools.live_tools, "word_live_create_document", fake_create_document)
    monkeypatch.setattr(live_api_tools, "_apply_default_template_live", fake_apply_template)

    result = json.loads(asyncio.run(live_api_tools.word_open()))

    assert result["success"] is True
    assert result["document"] == "Document1"
    assert result["full_path"] == ""
    assert result["visible"] is True
    assert result["template"] == "default_plain"
    assert "No path was provided" in result["message"]
    assert result["session_id"].startswith("word_")


def test_open_action_attach_without_path_attaches_active_document(monkeypatch):
    async def fake_list_open():
        return json.dumps({
            "success": True,
            "count": 2,
            "documents": [
                {"index": 1, "name": "Background.docx", "full_path": r"C:\Docs\Background.docx", "active": False},
                {"index": 2, "name": "Source.docx", "full_path": r"C:\Docs\Source.docx", "active": True, "pages": 3, "saved": True},
            ],
        })

    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_list_open", fake_list_open)

    result = json.loads(asyncio.run(live_api_tools.word_open(action="attach")))

    assert result["success"] is True
    assert result["document"] == "Source.docx"
    assert result["full_path"] == r"C:\Docs\Source.docx"
    assert result["session_id"].startswith("word_")
    assert live_api_tools._resolve_filename(result["session_id"]) == r"C:\Docs\Source.docx"


def test_open_action_list_returns_documents_without_session(monkeypatch):
    existing_session = live_api_tools._register_session({
        "document": "Existing.docx",
        "full_path": r"C:\Docs\Existing.docx",
    })

    async def fake_list_open():
        return json.dumps({
            "success": True,
            "count": 1,
            "documents": [
                {"index": 1, "name": "Source.docx", "full_path": r"C:\Docs\Source.docx", "active": True},
            ],
        })

    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_list_open", fake_list_open)

    result = json.loads(asyncio.run(live_api_tools.word_open(action="list")))

    assert result["success"] is True
    assert result["count"] == 1
    assert result["documents"][0]["active"] is True
    assert result["session_count"] == 1
    assert result["sessions"][0]["session_id"] == existing_session
    assert "session_id" not in result
    assert "word_open(action='attach')" in result["usage"]


def test_open_action_sessions_returns_registered_sessions():
    session_id = live_api_tools._register_session({
        "document": "Source.docx",
        "full_path": r"C:\Docs\Source.docx",
    })
    live_api_tools._handles[session_id]["match_1"] = {
        "target": {"kind": "selection", "start": 1, "end": 5}
    }

    result = json.loads(asyncio.run(live_api_tools.word_open(action="sessions")))

    assert result["success"] is True
    assert result["count"] == 1
    assert result["sessions"][0]["session_id"] == session_id
    assert result["sessions"][0]["filename"] == "Source.docx"
    assert result["sessions"][0]["full_path"] == r"C:\Docs\Source.docx"
    assert result["sessions"][0]["handle_count"] == 1


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

    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_list_open", fake_list_open)

    result = json.loads(asyncio.run(live_api_tools.word_open(action="attach", path="1")))

    assert result["success"] is True
    assert result["document"] == "Background.docx"
    assert live_api_tools._sessions[result["session_id"]]["full_path"] == r"C:\Docs\Background.docx"


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

    monkeypatch.setattr(live_api_tools.live_tools, "word_live_create_document", fake_create_document)
    monkeypatch.setattr(live_api_tools, "_apply_default_template_live", fake_apply_template)

    result = json.loads(asyncio.run(live_api_tools.word_open(action="new", visible=False)))

    assert result["success"] is True
    assert result["document"] == "Document1"
    assert result["template"] == "default_plain"
    assert result["template_result"]["success"] is True
    assert result["session_id"].startswith("word_")


def test_get_content_snapshot_and_diff_delegate_to_live_read_tools(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Source.docx",
        "full_path": r"C:\Docs\Source.docx",
    })
    calls = []

    async def fake_snapshot(filename):
        calls.append(("snapshot", filename))
        return json.dumps({"success": True, "snapshot_timestamp": 123})

    async def fake_diff(filename):
        calls.append(("diff", filename))
        return json.dumps({"success": True, "changes": [{"paragraph_index": 1}]})

    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_take_snapshot", fake_snapshot)
    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_get_diff", fake_diff)

    snapshot = json.loads(asyncio.run(live_api_tools.word_get_content(session_id, action="snapshot")))
    diff = json.loads(asyncio.run(live_api_tools.word_get_content(session_id, action="diff")))

    assert snapshot["success"] is True
    assert snapshot["session_id"] == session_id
    assert diff["changes"] == [{"paragraph_index": 1}]
    assert calls == [
        ("snapshot", r"C:\Docs\Source.docx"),
        ("diff", r"C:\Docs\Source.docx"),
    ]


def test_search_rejects_overlong_find_text_with_actionable_usage():
    session_id = live_api_tools._register_session({"document": "Source.docx"})

    result = json.loads(asyncio.run(live_api_tools.word_search(
        session_id=session_id,
        find_text="x" * 256,
    )))

    assert "Word Find limit" in result["error"]
    assert "word_search" in result["alternatives"][0]
    assert "paragraph_index" in result["usage"]


def test_replace_with_overlong_find_text_returns_actionable_usage():
    session_id = live_api_tools._register_session({"document": "Source.docx"})

    result = json.loads(asyncio.run(live_api_tools.word_edit(
        session_id=session_id,
        action="replace",
        find_text="x" * 256,
        text="replacement",
    )))

    assert "Word Find limit" in result["error"]
    assert "paragraph_index" in result["usage"]


def test_replace_can_target_one_based_paragraph_index(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Source.docx",
        "full_path": r"C:\Docs\Source.docx",
    })
    calls = []

    def fake_resolve_paragraph_target(filename, paragraph_index):
        calls.append(("target", filename, paragraph_index))
        return {"kind": "selection", "start": 10, "end": 30}

    async def fake_delete(filename, start, end, track_changes):
        calls.append(("delete", filename, start, end, track_changes))
        return json.dumps({"success": True, "deleted_text": "old paragraph"})

    async def fake_insert(filename, text, position, bookmark, track_changes):
        calls.append(("insert", filename, text, position, bookmark, track_changes))
        return json.dumps({"success": True, "inserted_text": text})

    monkeypatch.setattr(live_api_tools, "_resolve_paragraph_target", fake_resolve_paragraph_target)
    monkeypatch.setattr(live_api_tools.live_tools, "word_live_delete_text", fake_delete)
    monkeypatch.setattr(live_api_tools.live_tools, "word_live_insert_text", fake_insert)

    result = json.loads(asyncio.run(live_api_tools.word_edit(
        session_id=session_id,
        action="replace",
        paragraph_index=3,
        text="new paragraph",
        track_changes=True,
    )))

    assert result["success"] is True
    assert calls == [
        ("target", r"C:\Docs\Source.docx", 3),
        ("delete", r"C:\Docs\Source.docx", 10, 30, True),
        ("insert", r"C:\Docs\Source.docx", "new paragraph", "10", None, True),
    ]


def test_insert_paragraphs_uses_one_based_paragraph_index(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Source.docx",
        "full_path": r"C:\Docs\Source.docx",
    })
    calls = []

    async def fake_insert_paragraphs(filename, paragraphs, target_text, target_paragraph_index, position, style, track_changes):
        calls.append((filename, paragraphs, target_text, target_paragraph_index, position, style, track_changes))
        return json.dumps({"success": True, "inserted_count": len(paragraphs)})

    monkeypatch.setattr(live_api_tools.live_tools, "word_live_insert_paragraphs", fake_insert_paragraphs)

    result = json.loads(asyncio.run(live_api_tools.word_edit(
        session_id=session_id,
        action="insert_paragraphs",
        paragraphs=["Inserted"],
        paragraph_index=2,
        position="after",
        style="Normal",
        track_changes=True,
    )))

    assert result["success"] is True
    assert calls == [
        (r"C:\Docs\Source.docx", ["Inserted"], None, 1, "after", "Normal", True),
    ]


def test_comment_create_without_target_returns_usage_before_live_call(monkeypatch):
    session_id = live_api_tools._register_session({"document": "Source.docx"})

    async def fail_add_comment(*_args, **_kwargs):
        raise AssertionError("word_comment should validate missing targets before calling live add_comment")

    monkeypatch.setattr(live_api_tools.live_read_tools, "word_live_add_comment", fail_add_comment)

    result = json.loads(asyncio.run(live_api_tools.word_comment(
        session_id=session_id,
        action="create",
        text="Needs clarification.",
    )))

    assert result["error"] == "Comment target required"
    assert "paragraph_index" in result["usage"]


def test_blueprint_validation_rejects_unknown_blocks_and_bad_tables():
    errors = live_api_tools._validate_blueprint({
        "blocks": [
            {"type": "callout", "text": "Nope"},
            {"type": "table", "rows": [["A"], ["B", "C"]]},
        ],
    })

    assert any("unsupported type" in error for error in errors)
    assert any("same non-zero width" in error for error in errors)


def test_normalize_page_setup_accepts_custom_reference_size():
    normalized = live_api_tools._normalize_page_setup({
        "size": "custom",
        "width": 612,
        "height": 792,
        "orientation": "portrait",
    })

    assert normalized["size"] == "custom"
    assert normalized["width"] == 612
    assert normalized["height"] == 792


def test_blueprint_expected_counts():
    counts = live_api_tools._blueprint_expected_counts({
        "blocks": [
            {"type": "title_page", "title": "Cover"},
            {"type": "toc", "levels": 3},
            {"type": "heading", "text": "Title", "level": 1},
            {"type": "paragraph", "text": "Body"},
            {"type": "list", "items": ["One", "Two"]},
            {"type": "table", "rows": [["A", "B"]]},
            {"type": "image", "path": r"C:\Docs\screen.png"},
            {"type": "image_placeholder", "label": "screen"},
            {"type": "page_break"},
        ],
    })

    assert counts == {
        "paragraphs": 7,
        "tables": 1,
        "images": 1,
        "image_placeholders": 1,
        "blocks": 9,
    }


def test_natural_sort_orders_numbered_asset_names(tmp_path):
    names = ["pic 10.png", "pic 2.png", "pic 1.png", "notes.txt"]
    for name in names:
        (tmp_path / name).write_text("x", encoding="utf-8")

    assert [path.split("\\")[-1].split("/")[-1] for path in live_api_tools._asset_paths(str(tmp_path))] == [
        "pic 1.png",
        "pic 2.png",
        "pic 10.png",
    ]


def test_page_role_classifies_title_toc_and_body_pages():
    assert live_api_tools._page_role(1, 3, {2}) == "title_page"
    assert live_api_tools._page_role(2, 3, {2}) == "toc"
    assert live_api_tools._page_role(3, 3, {2}) == "body"


def test_asset_mapping_skips_title_page_images():
    images = [
        {"image_index": 1, "page": 1, "range_start": 10},
        {"image_index": 2, "page": 3, "range_start": 20},
        {"image_index": 3, "page": 3, "range_start": 30},
    ]

    warnings = live_api_tools._map_assets_to_body_images(
        images,
        [r"C:\Assets\pic 1.png", r"C:\Assets\pic 2.png"],
        first_body_page=3,
        toc_pages={2},
    )

    assert warnings == []
    assert "path" not in images[0]
    assert images[1]["path"] == r"C:\Assets\pic 1.png"
    assert images[2]["path"] == r"C:\Assets\pic 2.png"


def test_split_asset_paths_separates_and_orders_title_assets():
    title_assets, body_assets = live_api_tools._split_asset_paths([
        r"C:\Assets\2.1 Initial Setup pic 1.png",
        r"C:\Assets\Title page pic 3.png",
        r"C:\Assets\Ttitle page pic 2 logo.png",
        r"C:\Assets\Title Page pic 1.png",
        r"C:\Assets\2.2 AMD Adrenaline Installation pic 1.png",
    ])

    assert title_assets == [
        r"C:\Assets\Title Page pic 1.png",
        r"C:\Assets\Ttitle page pic 2 logo.png",
        r"C:\Assets\Title page pic 3.png",
    ]
    assert body_assets == [
        r"C:\Assets\2.1 Initial Setup pic 1.png",
        r"C:\Assets\2.2 AMD Adrenaline Installation pic 1.png",
    ]


def test_title_asset_mapping_adds_paths_to_title_page_images():
    images = [
        {"image_index": 1, "page": 1, "range_start": 10},
        {"image_index": 2, "page": 1, "range_start": 20},
        {"image_index": 3, "page": 3, "range_start": 30},
    ]

    warnings = live_api_tools._map_assets_to_title_images(
        images,
        [r"C:\Assets\Title Page pic 1.png", r"C:\Assets\Title page pic 2.png"],
        first_body_page=3,
        toc_pages={2},
    )

    assert warnings == []
    assert images[0]["path"] == r"C:\Assets\Title Page pic 1.png"
    assert images[1]["path"] == r"C:\Assets\Title page pic 2.png"
    assert "path" not in images[2]


def test_picture_shapes_are_emitted_as_image_records():
    images = live_api_tools._picture_shapes_as_images([
        {"shape_index": 1, "shape_type": 13, "page": 3, "width_pt": 396.0, "height_pt": 231.1, "wrap_type": 4},
        {"shape_index": 2, "shape_type": 17, "page": 1, "width_pt": 100.0, "height_pt": 50.0},
    ])

    assert images == [{
        "type": "image",
        "image_index": 1,
        "source": "floating_shape",
        "name": None,
        "page": 3,
        "paragraph_index": None,
        "range_start": 0,
        "width_pt": 396.0,
        "height_pt": 231.1,
        "wrapping": "infront",
        "wrap_type": 4,
        "left_pt": None,
        "top_pt": None,
        "relative_horizontal_position": None,
        "relative_vertical_position": None,
        "paragraph_after": True,
    }]


def test_picture_shapes_fall_back_to_name_when_type_is_unreadable():
    images = live_api_tools._picture_shapes_as_images([
        {"shape_index": 1, "name": "Picture 9", "shape_type": None, "page": 3, "width_pt": 396.0, "height_pt": 231.1},
        {"shape_index": 2, "name": "Text Box 1", "shape_type": None, "page": 1, "width_pt": 100.0, "height_pt": 50.0},
    ])

    assert len(images) == 1
    assert images[0]["source"] == "floating_shape"
    assert images[0]["width_pt"] == 396.0


def test_inline_image_inspection_tolerates_unreadable_title():
    class ParagraphFormat:
        Alignment = 1

    class Range:
        Start = 42

        def Information(self, _code):
            return 3
    Range.ParagraphFormat = ParagraphFormat()

    class InlineShape:
        Width = 396.04
        Height = 231.05

        @property
        def Title(self):
            raise RuntimeError("COM title failure")
    InlineShape.Range = Range()

    class Collection:
        Count = 1

        def __call__(self, _index):
            return InlineShape()

    class Doc:
        InlineShapes = Collection()

    images = live_api_tools._inspect_inline_images_live(Doc(), [(7, 40, 50)])

    assert images == [{
        "type": "image",
        "image_index": 1,
        "source": "inline",
        "name": "InlineShape 1",
        "page": 3,
        "paragraph_index": 7,
        "range_start": 42,
        "width_pt": 396.0,
        "height_pt": 231.1,
        "wrapping": "inline",
        "alignment": "center",
        "paragraph_after": True,
    }]


def test_shape_inspection_tolerates_unreadable_type():
    class Anchor:
        Start = 42

        def Information(self, _code):
            return 3

    class WrapFormat:
        Type = 4

    class Shape:
        Name = "Picture 9"
        Width = 396.04
        Height = 231.05
        Left = -999995.0
        Top = 55.55
        RelativeHorizontalPosition = 0
        RelativeVerticalPosition = 2

        @property
        def Type(self):
            raise RuntimeError("COM type failure")
    Shape.Anchor = Anchor()
    Shape.WrapFormat = WrapFormat()

    class Collection:
        Count = 1

        def __call__(self, _index):
            return Shape()

    class Doc:
        Shapes = Collection()

    shapes = live_api_tools._inspect_shapes_live(Doc(), [(7, 40, 50)])

    assert shapes == [{
        "type": "shape",
        "shape_index": 1,
        "name": "Picture 9",
        "page": 3,
        "paragraph_index": 7,
        "range_start": 42,
        "width_pt": 396.0,
        "height_pt": 231.1,
        "left_pt": -999995.0,
        "top_pt": 55.5,
        "shape_type": None,
        "auto_shape_type": None,
        "wrap_type": 4,
        "relative_horizontal_position": 0,
        "relative_vertical_position": 2,
    }]


def test_blueprint_list_format_maps_inspected_numbering():
    assert live_api_tools._blueprint_list_format({
        "kind": "word_list",
        "list_string": "3.",
        "list_value": 3,
        "level": 1,
    }) == {
        "list_type": "number",
        "level": 0,
        "continue_previous": True,
    }
    assert live_api_tools._blueprint_list_format({
        "kind": "word_list",
        "list_string": "",
        "list_value": 1,
        "level": 2,
    }) == {
        "list_type": "bullet",
        "level": 1,
        "continue_previous": False,
    }
    assert live_api_tools._blueprint_list_format({"kind": "none"}) is None


def test_inspected_pages_flatten_for_blueprint_create():
    normalized = live_api_tools._normalize_blueprint({
        "document": {
            "pages": [
                {"page": 1, "blocks": [{"type": "title_page", "title": "Cover"}]},
                {"page": 2, "blocks": [{"type": "toc"}]},
                {"page": 3, "blocks": [{"type": "heading", "text": "1. Overview", "level": 1}]},
            ],
        },
    })

    assert [block["type"] for block in normalized["blocks"]] == ["title_page", "toc", "heading"]


def test_inspected_document_blocks_take_precedence_over_pages():
    normalized = live_api_tools._normalize_blueprint({
        "document": {
            "blocks": [{"type": "heading", "text": "1. Overview", "level": 1}],
            "pages": [
                {"page": 1, "blocks": [{"type": "title_page", "title": "Cover"}]},
            ],
        },
    })

    assert [block["type"] for block in normalized["blocks"]] == ["heading"]


def test_page_blueprint_image_blocks_use_anchor_paragraph_metadata():
    pages, blocks, warnings = live_api_tools._build_page_blueprint(
        paragraph_records=[
            {
                "type": "heading",
                "text": "1. Overview",
                "style": "Heading 1",
                "page": 3,
                "paragraph_index": 10,
                "range_start": 100,
                "in_table": False,
            }
        ],
        image_records=[
            {
                "type": "image",
                "page": 3,
                "paragraph_index": 10,
                "range_start": 101,
                "path": r"C:\Assets\pic 1.png",
                "width_pt": 396,
                "height_pt": 231,
            }
        ],
        table_records=[],
        shape_records=[],
        page_count=3,
        first_body_page=3,
        toc_pages=set(),
    )

    image_block = next(block for block in blocks if block["type"] == "image")
    assert warnings == []
    assert pages[2]["role"] == "body"
    assert image_block["type"] == "image"
    assert image_block["anchor_paragraph_index"] == 10
    assert "paragraph_index" not in image_block


def test_layout_invalid_action_does_not_require_session():
    result = json.loads(asyncio.run(live_api_tools.word_layout(session_id="missing", action="bad")))

    assert result["error"] == "Invalid action"
    assert "page_setup" in result["valid_actions"]


def test_table_delete_column_requires_col_before_live_call(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Doc.docx",
        "full_path": r"C:\Docs\Doc.docx",
    })

    async def fail_modify_table(*_args, **_kwargs):
        raise AssertionError("word_table should validate delete_column col before COM call")

    monkeypatch.setattr(live_api_tools.live_tools, "word_live_modify_table", fail_modify_table)

    result = json.loads(asyncio.run(live_api_tools.word_table(
        session_id=session_id,
        action="delete_column",
    )))

    assert result["error"] == "delete_column requires col"
    assert "col=2" in result["usage"]


def test_layout_break_alias_inserts_page_break(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Doc.docx",
        "full_path": r"C:\Docs\Doc.docx",
    })
    calls = []

    def fake_insert_break(filename, break_kind, position="end", paragraph_index=None, break_type="next_page"):
        calls.append((filename, break_kind, position, paragraph_index, break_type))
        return {"success": True, "break": break_kind}

    monkeypatch.setattr(live_api_tools, "_layout_insert_break_live", fake_insert_break)

    result = json.loads(asyncio.run(live_api_tools.word_layout(
        session_id=session_id,
        action="break",
        break_type="page",
        position="end",
    )))

    assert result["success"] is True
    assert result["break"] == "page"
    assert calls == [(r"C:\Docs\Doc.docx", "page", "end", None, "next_page")]


def test_mutations_accept_full_public_tool_names(monkeypatch):
    calls = []

    async def fake_edit(**kwargs):
        calls.append(kwargs)
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    monkeypatch.setattr(live_api_tools, "word_edit", fake_edit)

    result = json.loads(asyncio.run(live_api_tools.word_mutations(
        session_id="word_123",
        action="apply",
        operations=[
            {"tool": "word_edit", "action": "insert", "text": "Hello"},
            {"tool": "mcp_word_word_edit", "action": "insert", "text": "World"},
        ],
    )))

    assert result["success"] is True
    assert [entry["tool"] for entry in result["results"]] == ["edit", "edit"]
    assert [entry["requested_tool"] for entry in result["results"]] == [
        "word_edit",
        "mcp_word_word_edit",
    ]
    assert calls == [
        {"action": "insert", "text": "Hello", "session_id": "word_123"},
        {"action": "insert", "text": "World", "session_id": "word_123"},
    ]


def test_media_insert_delegates_to_live_image_tool(monkeypatch):
    session_id = live_api_tools._register_session({
        "document": "Doc.docx",
        "full_path": r"C:\Docs\Doc.docx",
    })
    calls = []

    async def fake_insert_image(**kwargs):
        calls.append(kwargs)
        return json.dumps({"success": True, "image": "screen.png"})

    monkeypatch.setattr(live_api_tools.live_tools, "word_live_insert_image", fake_insert_image)

    result = json.loads(asyncio.run(live_api_tools.word_media(
        session_id=session_id,
        action="insert_image",
        path=r"C:\Assets\screen.png",
        width_pt=360,
        alignment="center",
        wrapping="infront",
        left_pt=12.5,
        top_pt=44.0,
        relative_horizontal_position=0,
        relative_vertical_position=2,
    )))

    assert result["success"] is True
    assert result["session_id"] == session_id
    assert calls == [{
        "filename": r"C:\Docs\Doc.docx",
        "image_path": r"C:\Assets\screen.png",
        "paragraph_index": None,
        "position": "end",
        "width_inches": None,
        "height_inches": None,
        "width_pt": 360,
        "height_pt": None,
        "alignment": "center",
        "wrapping": "infront",
        "border_style": None,
        "border_width_pt": None,
        "border_color": None,
        "link_to_file": False,
        "paragraph_after": True,
        "left_pt": 12.5,
        "top_pt": 44.0,
        "relative_horizontal_position": 0,
        "relative_vertical_position": 2,
    }]


def test_blueprint_validation_requires_image_path():
    errors = live_api_tools._validate_blueprint({
        "blocks": [{"type": "image"}],
    })

    assert errors == ["block 0 image requires path"]


def test_apply_blueprint_paragraph_replays_inspected_list_format(monkeypatch):
    calls = []
    paragraph_counts = iter([5, 6])

    async def fake_edit(**kwargs):
        calls.append(("edit", kwargs["paragraphs"], kwargs["style"]))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    async def fake_get_content(session_id, action="info", **_kwargs):
        return json.dumps({"success": True, "session_id": session_id, "paragraphs": next(paragraph_counts)})

    async def fake_format(**kwargs):
        calls.append((
            "format",
            kwargs["action"],
            kwargs["start_paragraph"],
            kwargs["end_paragraph"],
            kwargs["list_type"],
            kwargs["level"],
            kwargs["continue_previous"],
        ))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    monkeypatch.setattr(live_api_tools, "word_edit", fake_edit)
    monkeypatch.setattr(live_api_tools, "word_get_content", fake_get_content)
    monkeypatch.setattr(live_api_tools, "word_format", fake_format)

    result = asyncio.run(live_api_tools._apply_blueprint_block("word_123", {
        "type": "paragraph",
        "text": "Second item",
        "style": "List Paragraph",
        "numbering": {
            "kind": "word_list",
            "list_string": "2.",
            "list_value": 2,
            "level": 1,
        },
    }))

    assert result["result"]["success"] is True
    assert result["format_result"]["success"] is True
    assert calls == [
        ("edit", ["Second item"], "List Paragraph"),
        ("format", "list", 5, 5, "number", 0, True),
    ]


def test_apply_title_page_replays_simple_shapes(monkeypatch):
    calls = []

    async def fake_edit(**kwargs):
        calls.append(("edit", kwargs["paragraphs"], kwargs["style"]))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    async def fake_layout(**kwargs):
        calls.append(("layout", kwargs["action"]))
        return json.dumps({"success": True, "session_id": kwargs["session_id"]})

    def fake_insert_shape(session_id, shape):
        calls.append(("shape", session_id, shape["name"], shape.get("text")))
        return {"success": True, "name": shape["name"]}

    monkeypatch.setattr(live_api_tools, "word_edit", fake_edit)
    monkeypatch.setattr(live_api_tools, "word_layout", fake_layout)
    monkeypatch.setattr(live_api_tools, "_insert_title_shape_live", fake_insert_shape)

    result = asyncio.run(live_api_tools._apply_blueprint_block("word_123", {
        "type": "title_page",
        "text": "",
        "page_break_after": True,
        "shapes": [
            {"name": "Text Box 1", "text": "AI DAY IN THE LIFE", "width_pt": 500, "height_pt": 90},
            {"name": "Rectangle 2", "auto_shape_type": 1, "width_pt": 600, "height_pt": 56},
        ],
    }))

    assert result["result"]["success"] is True
    assert result["shape_results"] == [
        {"success": True, "name": "Text Box 1"},
        {"success": True, "name": "Rectangle 2"},
    ]
    assert calls == [
        ("shape", "word_123", "Text Box 1", "AI DAY IN THE LIFE"),
        ("shape", "word_123", "Rectangle 2", None),
        ("layout", "page_break"),
    ]


def test_blueprint_create_dispatches_blocks_in_order(monkeypatch):
    calls = []

    async def fake_open(action="open", visible=True, **_kwargs):
        calls.append(("open", action, visible))
        session_id = live_api_tools._register_session({
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

    async def fake_media(**kwargs):
        calls.append((
            "media",
            kwargs["action"],
            kwargs["path"],
            kwargs["width_pt"],
            kwargs["alignment"],
            kwargs.get("wrapping"),
            kwargs.get("left_pt"),
            kwargs.get("top_pt"),
            kwargs.get("relative_horizontal_position"),
            kwargs.get("relative_vertical_position"),
        ))
        return json.dumps({"success": True, "session_id": kwargs["session_id"], "image": "screen.png"})

    async def fake_get_content(session_id, action="info", **_kwargs):
        return json.dumps({"success": True, "session_id": session_id, "paragraphs": 10, "tables": 1})

    monkeypatch.setattr(live_api_tools, "word_open", fake_open)
    monkeypatch.setattr(live_api_tools, "word_layout", fake_layout)
    monkeypatch.setattr(live_api_tools, "word_edit", fake_edit)
    monkeypatch.setattr(live_api_tools, "word_table", fake_table)
    monkeypatch.setattr(live_api_tools, "word_media", fake_media)
    monkeypatch.setattr(live_api_tools, "word_get_content", fake_get_content)

    result = json.loads(asyncio.run(live_api_tools.word_blueprint(
        action="create",
        blueprint={
            "page_setup": {"size": "letter", "orientation": "portrait"},
            "properties": {"title": "Demo"},
            "blocks": [
                {"type": "heading", "text": "Title", "level": 1},
                {"type": "paragraph", "text": "Body"},
                {"type": "table", "rows": [["A", "B"], ["1", "2"]]},
                {"type": "page_break"},
                {
                    "type": "image",
                    "path": r"C:\Assets\screen.png",
                    "width_pt": 360,
                    "alignment": "center",
                    "wrap_type": 4,
                    "left_pt": -999995.0,
                    "top_pt": 55.5,
                    "relative_horizontal_position": 0,
                    "relative_vertical_position": 2,
                },
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
    assert ("media", "insert_image", r"C:\Assets\screen.png", 360, "center", "infront", -999995.0, 55.5, 0, 2) in calls
    assert result["warnings"][0]["warning"].startswith("image_placeholder")
