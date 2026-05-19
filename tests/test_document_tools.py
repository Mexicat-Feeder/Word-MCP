import asyncio
from pathlib import Path

from word_document_server.tools.document_tools import create_document


def test_create_document_creates_missing_parent_directory(tmp_path: Path):
    target = tmp_path / "missing" / "nested" / "created.docx"

    result = asyncio.run(create_document(str(target), title="Created", author="Tester"))

    assert "created successfully" in result
    assert target.exists()
