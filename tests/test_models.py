from word_document_master_tool.core.models import ToolSettings


def test_tool_settings_validation_no_errors():
    settings = ToolSettings(output_folder="C:/test")
    assert len(settings.validate_errors()) == 0


def test_page_numbering_conflict():
    settings = ToolSettings(output_folder="C:/test")
    settings.page_numbering.enabled = True
    settings.page_numbering.continuous = True
    settings.page_numbering.restart_each_document = True
    errors = settings.validate_errors()
    assert any("Сквозная нумерация страниц" in e for e in errors)


def test_headers_footers_conflict():
    settings = ToolSettings(output_folder="C:/test")
    settings.page_numbering.enabled = True
    settings.page_numbering.remove_headers_footers = True
    settings.page_numbering.preserve_headers_footers = True
    errors = settings.validate_errors()
    assert any("Очистка и сохранение колонтитулов" in e for e in errors)


def test_footnote_conflict():
    settings = ToolSettings(output_folder="C:/test")
    settings.footnotes.enabled = True
    settings.footnotes.continuous = True
    settings.footnotes.restart_each_document = True
    errors = settings.validate_errors()
    assert any("Сквозная нумерация сносок" in e for e in errors)
