from word_document_master_tool.core.execution_plan import build_execution_plan
from word_document_master_tool.core.models import ToolSettings


def test_combined_pdf_without_individual_pdfs_is_error() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"
    settings.pdf.export_sources = False
    settings.pdf.export_processed_copies = False
    settings.pdf.export_merged = False
    settings.pdf.merge_generated_pdfs = True

    plan = build_execution_plan(settings, selected_count=1)

    assert any("общий PDF" in error for error in plan.errors)


def test_pdf_without_pdf_folder_is_error() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.export_sources = True
    settings.pdf.output_folder = ""

    plan = build_execution_plan(settings, selected_count=1)

    assert any("папка PDF" in error for error in plan.errors)


def test_source_pdf_warns_that_numbering_is_ignored() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.export_sources = True
    settings.pdf.output_folder = "C:/pdf"
    settings.page_numbering.enabled = True

    plan = build_execution_plan(settings, selected_count=1)

    assert any("PDF исходников" in warning for warning in plan.warnings)


def test_no_selected_documents_is_error() -> None:
    settings = ToolSettings(source_folder="C:/src", output_folder="C:/out")
    settings.pdf.output_folder = "C:/pdf"

    plan = build_execution_plan(settings, selected_count=0)

    assert any("Не выбраны документы" in error for error in plan.errors)
