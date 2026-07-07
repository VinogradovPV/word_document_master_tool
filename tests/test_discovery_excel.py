from word_document_master_tool.core.models import SourceKind
from word_document_master_tool.filesystem.discovery import DocumentDiscovery


def test_discovery_finds_word_and_excel_sources(tmp_path):
    (tmp_path / "a.docx").write_text("word")
    (tmp_path / "b.xls").write_text("excel")
    (tmp_path / "c.xlsx").write_text("excel")
    (tmp_path / "~$temp.xlsx").write_text("temp")
    (tmp_path / "image.png").write_text("png")

    items = DocumentDiscovery().find_documents(str(tmp_path))

    names = {item.file_name: item for item in items}
    assert set(names) == {"a.docx", "b.xls", "c.xlsx"}
    assert names["a.docx"].source_kind == SourceKind.WORD
    assert names["b.xls"].source_kind == SourceKind.EXCEL
    assert names["c.xlsx"].source_kind == SourceKind.EXCEL
