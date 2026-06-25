from datetime import datetime

from word_document_master_tool.core.logging_service import LoggingService
from word_document_master_tool.core.models import DocumentItem, DocumentStatus


def test_logging_utf8(tmp_path):
    log_service = LoggingService(str(tmp_path))
    log_file = log_service.log_file
    
    item = DocumentItem(
        order_index=1,
        file_path="test.docx",
        file_name="тест_кириллица.docx",
        extension="docx",
        size_bytes=100,
        modified_at=datetime.now(),
        status=DocumentStatus.OK
    )
    
    log_service.log_item(item, superscript_pages="1, 3", message="Успешно")
    
    # Проверяем, что файл в UTF-8 и содержит кириллицу
    with open(log_file, encoding="utf-8") as f:
        content = f.read()
        assert "тест_кириллица.docx" in content
        assert "Успешно" in content
        assert "Надстрочные символы, стр." in content
