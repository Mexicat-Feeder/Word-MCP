import pytest

from word_document_server.core import word_com


class BrokenDocument:
    @property
    def Name(self):
        raise RuntimeError("<unknown>.Name")

    @property
    def FullName(self):
        raise RuntimeError("<unknown>.FullName")


class FakeDocument:
    def __init__(self, name, full_name):
        self.Name = name
        self.FullName = full_name


class FakeDocuments:
    def __init__(self, docs):
        self._docs = docs
        self.Count = len(docs)

    def __call__(self, index):
        return self._docs[index - 1]


class FakeApp:
    def __init__(self, docs, active=None):
        self.Documents = FakeDocuments(docs)
        self.ActiveDocument = active if active is not None else (docs[0] if docs else None)


def test_has_readable_documents_rejects_broken_doc_proxy():
    app = FakeApp([BrokenDocument()])

    assert word_com._has_readable_documents(app) is False


def test_has_readable_documents_accepts_any_named_doc():
    app = FakeApp([BrokenDocument(), FakeDocument("Source.docx", r"C:\Docs\Source.docx")])

    assert word_com._has_readable_documents(app) is True


def test_find_document_skips_broken_docs_and_matches_full_path():
    doc = FakeDocument("Source.docx", r"C:\Docs\Source.docx")
    app = FakeApp([BrokenDocument(), doc])

    assert word_com.find_document(app, r"C:\Docs\Source.docx") is doc


def test_find_document_uses_first_readable_when_active_is_broken():
    doc = FakeDocument("Source.docx", r"C:\Docs\Source.docx")
    app = FakeApp([BrokenDocument(), doc], active=BrokenDocument())

    assert word_com.find_document(app) is doc


def test_find_document_reports_unreadable_entries():
    app = FakeApp([BrokenDocument()])

    with pytest.raises(ValueError, match="<unreadable:1>"):
        word_com.find_document(app, "Missing.docx")
