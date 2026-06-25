from word_document_master_tool.filesystem.discovery import find_word_documents
from word_document_master_tool.filesystem.paths import (
    build_safe_file_name,
    can_write_to_folder,
    ensure_unique_file_path,
)


def test_find_word_documents(tmp_path):
    # Создаем тестовые файлы
    (tmp_path / "test1.docx").write_text("content")
    (tmp_path / "test2.doc").write_text("content")
    (tmp_path / "~$temp.docx").write_text("content")
    (tmp_path / "image.png").write_text("content")

    items = find_word_documents(str(tmp_path))
    assert len(items) == 2
    names = [item.file_name for item in items]
    assert "test1.docx" in names
    assert "test2.doc" in names
    assert "~$temp.docx" not in names


def test_build_safe_file_name():
    assert build_safe_file_name("file/name*.docx") == "file_name_.docx"
    assert build_safe_file_name('con:test"?.rtf') == "con_test__.rtf"


def test_ensure_unique_file_path(tmp_path):
    file_path = tmp_path / "test.docx"
    file_path.write_text("content")

    unique_path = ensure_unique_file_path(str(file_path))
    assert unique_path != str(file_path)
    assert "test_001.docx" in unique_path


def test_can_write_to_folder(tmp_path):
    assert can_write_to_folder(str(tmp_path)) is True
    # На Linux сложно имитировать отсутствие прав без sudo/chmod, 
    # но базовый тест на существующую папку проходит.
