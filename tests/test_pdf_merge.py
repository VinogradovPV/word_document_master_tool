import os

from pypdf import PdfWriter

from word_document_master_tool.core.models import DocumentItem, ToolSettings
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
