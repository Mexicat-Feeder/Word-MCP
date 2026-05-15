from word_document_server.tools.live_tools import (
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
