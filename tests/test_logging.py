from datetime import datetime

from word_document_master_tool.core.logging_service import (
    ProcessingLogger,
    setup_application_logging,
)
from word_document_master_tool.core.models import DocumentItem, DocumentStatus


def test_logging_utf8(tmp_path):
    logger = ProcessingLogger(str(tmp_path))
    log_file = logger.log_path

    item = DocumentItem(
        order_index=1,
        file_path="test.docx",
        file_name="тест_кириллица.docx",
        extension="docx",
        size_bytes=100,
        modified_at=datetime.now(),
        status=DocumentStatus.OK,
    )

    logger.log_item(item, superscript_pages="1, 3", message="Успешно")

    # Проверяем, что файл в UTF-8 и содержит кириллицу
    with open(log_file, encoding="utf-8") as f:
        content = f.read()
        assert "тест_кириллица.docx" in content
        assert "Успешно" in content
        assert "Надстрочные символы, стр." in content


def test_application_logging(tmp_path):
    setup_application_logging(str(tmp_path))
    import logging

    logging.info("Test message")
    log_file = tmp_path / "application.log"
    assert log_file.exists()
    with open(log_file, encoding="utf-8") as f:
        assert "Test message" in f.read()
