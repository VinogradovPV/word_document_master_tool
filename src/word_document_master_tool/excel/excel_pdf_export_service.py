import contextlib
import logging
import os

from ..core.logging_service import ProcessingLogger
from ..core.models import DocumentItem, DocumentStatus, SourceKind, ToolSettings
from ..filesystem.paths import build_safe_file_name
from .excel_app import ExcelApp


class ExcelPdfExportService:
    """
    Сервис экспорта Excel-исходников в PDF через Microsoft Excel COM.
    """

    def __init__(self, settings: ToolSettings, excel_app_factory=ExcelApp):
        self.settings = settings
        self.excel_app_factory = excel_app_factory

    def export_source_to_pdf(
        self,
        item: DocumentItem,
        excel_app: ExcelApp | None = None,
        log_service: ProcessingLogger | None = None,
    ) -> bool:
        if not item.is_selected or item.source_kind != SourceKind.EXCEL:
            return False

        if excel_app is not None:
            return self._export_with_app(item, excel_app, log_service)

        try:
            with self.excel_app_factory() as app:
                return self._export_with_app(item, app, log_service)
        except Exception as exc:
            self._mark_error(item, exc, log_service)
            return False

    def _export_with_app(
        self,
        item: DocumentItem,
        excel_app: ExcelApp,
        log_service: ProcessingLogger | None,
    ) -> bool:
        workbook = None
        pdf_path = self._get_output_path(item)
        try:
            workbook = excel_app.open_workbook(item.file_path, read_only=True)
            if workbook is None:
                raise RuntimeError("Не удалось открыть книгу Excel")

            success = excel_app.export_workbook_as_pdf(
                workbook,
                pdf_path,
                include_properties=self.settings.pdf.include_properties,
                ignore_print_areas=False,
            )
            if not success:
                raise RuntimeError("ExportAsFixedFormat вернул ошибку")

            item.pdf_path = pdf_path
            item.status = DocumentStatus.OK
            item.error_message = ""
            logging.info(f"Excel exported to PDF: {item.file_name}")
            if log_service is not None:
                log_service.log_item(item, message="Excel exported to PDF")
            return True
        except Exception as exc:
            self._mark_error(item, exc, log_service)
            return False
        finally:
            if workbook is not None:
                with contextlib.suppress(Exception):
                    excel_app.close_workbook(workbook)

    def _get_output_path(self, item: DocumentItem) -> str:
        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        os.makedirs(output_folder, exist_ok=True)
        base_name = build_safe_file_name(os.path.splitext(item.file_name)[0])
        return os.path.join(output_folder, f"{base_name}_excel.pdf")

    @staticmethod
    def _mark_error(
        item: DocumentItem,
        exc: Exception,
        log_service: ProcessingLogger | None,
    ) -> None:
        item.status = DocumentStatus.ERROR
        item.error_message = str(exc)
        logging.error(f"Excel PDF export failed for {item.file_name}: {exc}")
        if log_service is not None:
            log_service.log_item(item, message=str(exc))
