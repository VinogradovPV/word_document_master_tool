from word_document_master_tool.core.execution_plan import build_execution_plan
from word_document_master_tool.core.models import ToolSettings


def test_excel_files_are_visible_in_execution_plan() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"
    settings.pdf.export_excel_sources = True

    plan = build_execution_plan(settings, selected_count=3, word_count=1, excel_count=2)

    assert "Excel-файлов выбрано: 2" in plan.lines()
    assert "PDF из Excel: да" in plan.lines()


def test_excel_pdf_disabled_warns_that_excel_files_are_skipped() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"
    settings.pdf.export_excel_sources = False

    plan = build_execution_plan(settings, selected_count=1, word_count=0, excel_count=1)

    assert any("PDF из Excel выключен" in warning for warning in plan.warnings)
    assert any("Word-настройки" in warning for warning in plan.warnings)


def test_excel_only_pdf_does_not_require_word() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"
    settings.pdf.export_excel_sources = True

    plan = build_execution_plan(
        settings,
        selected_count=1,
        word_count=0,
        excel_count=1,
        word_available=False,
    )

    assert not any("Word COM" in error for error in plan.errors)


def test_combined_pdf_allows_excel_individual_pdf() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"
    settings.pdf.export_excel_sources = True
    settings.pdf.merge_generated_pdfs = True

    plan = build_execution_plan(settings, selected_count=1, word_count=0, excel_count=1)

    assert not any("отдельных PDF" in error for error in plan.errors)
