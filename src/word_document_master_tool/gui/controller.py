import logging
import threading
from collections.abc import Callable

from ..core.logging_service import ProcessingLogger, setup_application_logging
from ..core.models import SourceKind, ToolSettings
from ..pdf.pdf_export_coordinator import PdfExportCoordinator
from ..pdf.pdf_export_service import PdfExportService
from ..pdf.pdf_merge_service import PdfMergeService
from ..word.merge_service import WordMergeService
from ..word.source_processing_service import SourceProcessingService
from ..word.split_service import WordSplitService
from ..word.word_app import WordApp
from .state import GuiState


class AppController:
    """
    Контроллер приложения, связывающий GUI и бизнес-логику.
    """

    def __init__(self, state: GuiState):
        self.state = state
        self._is_running = False

    def run_in_thread(self, task: Callable, on_complete: Callable | None = None):
        """Запускает задачу в отдельном потоке."""
        if self._is_running:
            return

        def wrapper():
            self._is_running = True
            try:
                task()
            except Exception as e:
                logging.error(f"Error in background task: {e}")
            finally:
                self._is_running = False
                if on_complete:
                    on_complete()

        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()

    def process_files(self, settings: ToolSettings, progress_callback: Callable[[int, int], None]):
        """Обработка выбранных файлов (правки, комментарии, экспорт в PDF)."""
        selected_items = [item for item in self.state.documents if item.is_selected]
        if not selected_items:
            return

        setup_application_logging(settings.output_folder)
        logger = ProcessingLogger(settings.output_folder)

        total = len(selected_items)
        if self._word_processing_requested(settings):
            source_service = SourceProcessingService(settings)
            source_service.process_copies(selected_items, log_service=logger)

        if settings.pdf.export_sources or settings.pdf.export_excel_sources:
            pdf_coordinator = PdfExportCoordinator(settings)
            pdf_coordinator.export_sources_to_pdf(selected_items, log_service=logger)

        if settings.pdf.merge_generated_pdfs:
            pdf_merge_service = PdfMergeService(settings)
            result = pdf_merge_service.merge_pdfs(selected_items)
            if not result.success:
                logging.error(f"PDF merge failed: {result.error_message}")

        progress_callback(total, total)

    def merge_documents(
        self, settings: ToolSettings, progress_callback: Callable[[int, int], None]
    ):
        """Слияние выбранных документов."""
        selected_items = [item for item in self.state.documents if item.is_selected]
        if not selected_items:
            return

        setup_application_logging(settings.output_folder)
        logger = ProcessingLogger(settings.output_folder)

        progress_callback(0, 100)
        try:
            if settings.pdf.export_sources or settings.pdf.export_excel_sources:
                pdf_coordinator = PdfExportCoordinator(settings)
                pdf_coordinator.export_sources_to_pdf(selected_items, log_service=logger)

            word_items = [item for item in selected_items if item.source_kind == SourceKind.WORD]
            merged_path = ""
            if word_items:
                merge_service = WordMergeService(settings)
                merged_path = merge_service.merge_documents(word_items)

            if settings.pdf.export_merged and merged_path:
                pdf_service = PdfExportService(settings)
                with WordApp() as word:
                    if not word.app:
                        raise RuntimeError("Word is not available for PDF export.")
                    pdf_service.export_processed_copy(merged_path, word)

            if settings.pdf.merge_generated_pdfs:
                pdf_merge_service = PdfMergeService(settings)
                result = pdf_merge_service.merge_pdfs(selected_items)
                if not result.success:
                    logging.error(f"PDF merge failed: {result.error_message}")

        except Exception as e:
            logging.error(f"Merge failed: {e}")
        finally:
            progress_callback(100, 100)

    def split_documents(
        self, settings: ToolSettings, progress_callback: Callable[[int, int], None]
    ):
        """Разделение документа по маркерам."""
        setup_application_logging(settings.output_folder)
        with WordApp() as word:
            if not word.app:
                logging.error("Word is not available for splitting.")
                return

            split_service = WordSplitService(word, settings)
            progress_callback(0, 100)
            try:
                # Разделяем файл по маркерам
                split_service.split_by_markers(settings.output_folder)
            except Exception as e:
                logging.error(f"Split failed: {e}")
            finally:
                progress_callback(100, 100)

    @staticmethod
    def _word_processing_requested(settings: ToolSettings) -> bool:
        return any(
            [
                settings.source_processing.accept_revisions,
                settings.source_processing.disable_track_changes,
                settings.source_processing.remove_comments,
                settings.page_numbering.enabled,
                settings.footnotes.enabled,
                settings.pdf.export_processed_copies,
            ]
        )
