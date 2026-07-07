import logging

from ..core.logging_service import ProcessingLogger
from ..core.models import DocumentItem, DocumentStatus, SourceKind, ToolSettings
from ..excel.excel_app import ExcelApp
from ..excel.excel_pdf_export_service import ExcelPdfExportService
from ..word.word_app import WordApp
from .pdf_export_service import PdfExportService


class PdfExportCoordinator:
    """
    Диспетчер PDF-экспорта для смешанных источников Word и Excel.
    """

    def __init__(
        self,
        settings: ToolSettings,
        word_app_factory=WordApp,
        excel_app_factory=ExcelApp,
        word_pdf_service: PdfExportService | None = None,
        excel_pdf_service: ExcelPdfExportService | None = None,
    ):
        self.settings = settings
        self.word_app_factory = word_app_factory
        self.excel_app_factory = excel_app_factory
        self.word_pdf_service = word_pdf_service or PdfExportService(settings)
        self.excel_pdf_service = excel_pdf_service or ExcelPdfExportService(
            settings, excel_app_factory
        )

    def export_sources_to_pdf(
        self,
        items: list[DocumentItem],
        log_service: ProcessingLogger | None = None,
    ) -> None:
        selected_items = [item for item in items if item.is_selected]
        if self.settings.pdf.export_sources:
            self._export_word_sources(selected_items, log_service)
        if self.settings.pdf.export_excel_sources:
            self._export_excel_sources(selected_items, log_service)

    def _export_word_sources(
        self,
        items: list[DocumentItem],
        log_service: ProcessingLogger | None,
    ) -> None:
        word_items = [item for item in items if item.source_kind == SourceKind.WORD]
        if not word_items:
            return

        try:
            with self.word_app_factory() as word:
                if not word.app:
                    raise RuntimeError("Microsoft Word COM недоступен")
                for item in word_items:
                    self.word_pdf_service.export_source_to_pdf(item, word)
                    if log_service is not None:
                        message = item.error_message or "Word exported to PDF"
                        log_service.log_item(item, message=message)
        except Exception as exc:
            for item in word_items:
                self._mark_error(item, exc, log_service)

    def _export_excel_sources(
        self,
        items: list[DocumentItem],
        log_service: ProcessingLogger | None,
    ) -> None:
        excel_items = [item for item in items if item.source_kind == SourceKind.EXCEL]
        if not excel_items:
            return

        try:
            with self.excel_app_factory() as excel:
                for item in excel_items:
                    self.excel_pdf_service.export_source_to_pdf(
                        item, excel_app=excel, log_service=log_service
                    )
        except Exception as exc:
            for item in excel_items:
                self._mark_error(item, exc, log_service)

    @staticmethod
    def _mark_error(
        item: DocumentItem,
        exc: Exception,
        log_service: ProcessingLogger | None,
    ) -> None:
        item.status = DocumentStatus.ERROR
        item.error_message = str(exc)
        logging.error(f"PDF export failed for {item.file_name}: {exc}")
        if log_service is not None:
            log_service.log_item(item, message=str(exc))
