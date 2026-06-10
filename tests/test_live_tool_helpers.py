from word_document_server.tools.live_tools import (
    _coerce_int_list,
    _coerce_table_entry,
    _normalize_paragraph_insert_items,
    _paragraph_insert_payload,
    _word_specials_to_text,
)


def test_word_specials_to_text_converts_find_replace_tokens():
    assert _word_specials_to_text("A^pB^tC^mD^sE") == "A\rB\tC\x0cD\u00a0E"


def test_paragraph_insert_payload_keeps_paragraphs_separate():
    payload = _paragraph_insert_payload(["First paragraph", "Second paragraph"])

    assert payload == "First paragraph\rSecond paragraph\r"


def test_paragraph_insert_payload_normalizes_existing_newlines():
    payload = _paragraph_insert_payload(["First\nline", "Second\r\nline\r"])

    assert payload == "First\rline\rSecond\rline\r"


def test_paragraph_insert_payload_extracts_text_from_agent_dicts():
    payload = _paragraph_insert_payload([
        {"style": "Heading 1", "text": "This is a test."},
        {"style": "Normal", "text": "Second paragraph."},
    ])

    assert payload == "This is a test.\rSecond paragraph.\r"


def test_normalize_paragraph_insert_items_preserves_per_item_style():
    items = _normalize_paragraph_insert_items([
        {"style": "Heading 1", "text": "Title"},
        "Body",
    ], fallback_style="Normal")

    assert items == [
        {"text": "Title", "style": "Heading 1"},
        {"text": "Body", "style": "Normal"},
    ]


def test_coerce_int_list_accepts_string_numbers():
    assert _coerce_int_list(["1", 2, "3"]) == [1, 2, 3]


def test_coerce_table_entry_accepts_nested_list_and_json_string():
    assert _coerce_table_entry([1, 0, "#DDDDDD"], 3, "cell_shading") == [
        1,
        0,
        "#DDDDDD",
    ]
    assert _coerce_table_entry('[1, 1, true]', 3, "cell_bold") == [1, 1, True]
