import os

from pypdf import PdfWriter

from word_document_master_tool.core.models import DocumentItem, SourceKind, ToolSettings
from word_document_master_tool.pdf.pdf_merge_service import PdfMergeService


def create_dummy_pdf(path):
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    with open(path, "wb") as f:
        writer.write(f)
    writer.close()


def test_merge_pdfs(tmp_path):
    pdf1 = tmp_path / "1.pdf"
    pdf2 = tmp_path / "2.pdf"
    create_dummy_pdf(str(pdf1))
    create_dummy_pdf(str(pdf2))

    settings = ToolSettings(output_folder=str(tmp_path), output_file_name="test")
    service = PdfMergeService(settings)

    items = [
        DocumentItem(1, "1.docx", "1.docx", "docx", 0, None, pdf_path=str(pdf1)),
        DocumentItem(2, "2.docx", "2.docx", "docx", 0, None, pdf_path=str(pdf2)),
    ]

    result = service.merge_pdfs(items)
    assert result.success
    assert os.path.exists(result.output_path)
    assert "test_combined.pdf" in result.output_path
    assert result.merged_count == 2


def test_merge_pdfs_keeps_table_order_for_word_and_excel(tmp_path):
    pdf1 = tmp_path / "word.pdf"
    pdf2 = tmp_path / "excel.pdf"
    create_dummy_pdf(str(pdf1))
    create_dummy_pdf(str(pdf2))

    settings = ToolSettings(output_folder=str(tmp_path), output_file_name="mixed")
    settings.pdf.output_folder = str(tmp_path)
    service = PdfMergeService(settings)

    items = [
        DocumentItem(1, "a.docx", "a.docx", "docx", 0, None, pdf_path=str(pdf1)),
        DocumentItem(
            2,
            "b.xlsx",
            "b.xlsx",
            "xlsx",
            0,
            None,
            source_kind=SourceKind.EXCEL,
            pdf_path=str(pdf2),
        ),
    ]

    result = service.merge_pdfs(items)

    assert result.success
    assert result.merged_count == 2
    assert result.skipped_paths == []


def test_merge_pdfs_reports_selected_sources_without_pdf(tmp_path):
    settings = ToolSettings(output_folder=str(tmp_path), output_file_name="empty")
    settings.pdf.output_folder = str(tmp_path)
    service = PdfMergeService(settings)

    item = DocumentItem(
        1,
        "b.xlsx",
        "b.xlsx",
        "xlsx",
        0,
        None,
        source_kind=SourceKind.EXCEL,
    )

    result = service.merge_pdfs([item])

    assert result.success is False
    assert item.file_path in result.skipped_paths
