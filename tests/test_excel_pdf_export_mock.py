from datetime import datetime

from word_document_master_tool.core.models import (
    DocumentItem,
    DocumentStatus,
    SourceKind,
    ToolSettings,
)
from word_document_master_tool.excel.excel_pdf_export_service import ExcelPdfExportService


class FakeWorkbook:
    def __init__(self, *, fail_export: bool = False):
        self.fail_export = fail_export
        self.export_kwargs = None
        self.close_save_changes = None

    def ExportAsFixedFormat(self, **kwargs):
        if self.fail_export:
            raise RuntimeError("export failed")
        self.export_kwargs = kwargs

    def Close(self, SaveChanges=False):
        self.close_save_changes = SaveChanges


class FakeExcelApp:
    def __init__(self, workbook: FakeWorkbook):
        self.workbook = workbook
        self.entered = False
        self.exited = False
        self.opened_path = ""
        self.opened_read_only = None

    def __enter__(self):
        self.entered = True
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.exited = True

    def open_workbook(self, file_path: str, read_only: bool = True):
        self.opened_path = file_path
        self.opened_read_only = read_only
        return self.workbook

    def export_workbook_as_pdf(self, workbook, output_path: str, **kwargs) -> bool:
        workbook.ExportAsFixedFormat(Filename=output_path, **kwargs)
        return True

    @staticmethod
    def close_workbook(workbook) -> None:
        workbook.Close(SaveChanges=False)


def _excel_item(path) -> DocumentItem:
    return DocumentItem(
        order_index=1,
        file_path=str(path),
        file_name=path.name,
        extension=path.suffix.lstrip("."),
        size_bytes=path.stat().st_size,
        modified_at=datetime.fromtimestamp(path.stat().st_mtime),
        source_kind=SourceKind.EXCEL,
    )


def test_excel_pdf_export_uses_read_only_workbook_and_closes_without_save(tmp_path):
    source = tmp_path / "report.xlsx"
    source.write_text("excel")
    output = tmp_path / "pdf"
    settings = ToolSettings(output_folder=str(tmp_path))
    settings.pdf.output_folder = str(output)
    item = _excel_item(source)
    workbook = FakeWorkbook()
    created_apps = []

    def factory():
        app = FakeExcelApp(workbook)
        created_apps.append(app)
        return app

    service = ExcelPdfExportService(settings, excel_app_factory=factory)

    assert service.export_source_to_pdf(item) is True

    app = created_apps[0]
    assert app.entered is True
    assert app.exited is True
    assert app.opened_path == str(source)
    assert app.opened_read_only is True
    assert workbook.export_kwargs["Filename"].endswith("report_excel.pdf")
    assert workbook.close_save_changes is False
    assert item.file_path == str(source)
    assert item.status == DocumentStatus.OK
    assert item.pdf_path.endswith("report_excel.pdf")


def test_excel_pdf_export_marks_error_and_closes_workbook(tmp_path):
    source = tmp_path / "broken.xls"
    source.write_text("excel")
    settings = ToolSettings(output_folder=str(tmp_path))
    settings.pdf.output_folder = str(tmp_path / "pdf")
    item = _excel_item(source)
    workbook = FakeWorkbook(fail_export=True)
    service = ExcelPdfExportService(settings)

    assert service.export_source_to_pdf(item, excel_app=FakeExcelApp(workbook)) is False

    assert item.status == DocumentStatus.ERROR
    assert "export failed" in item.error_message
    assert item.file_path == str(source)
    assert workbook.close_save_changes is False
