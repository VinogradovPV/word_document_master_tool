import logging
import threading
from collections.abc import Callable

from ..core.models import DocumentStatus, ToolSettings
from ..pdf.pdf_export_service import PdfExportService
from ..pdf.pdf_merge_service import PdfMergeService
from ..word.merge_service import WordMergeService
from ..word.source_processing_service import SourceProcessingService
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

        with WordApp() as word:
            source_service = SourceProcessingService(word, settings)
            pdf_service = PdfExportService(word, settings)
            
            total = len(selected_items)
            for i, item in enumerate(selected_items):
                progress_callback(i, total)
                item.status = DocumentStatus.PROCESSING
                
                try:
                    # 1. Подготовка копии
                    processed_path = source_service.prepare_processed_copy(item.file_path)
                    
                    # 2. Экспорт в PDF (если нужно)
                    if settings.pdf.export_sources:
                        pdf_path = pdf_service.export_source_as_readonly(item.file_path)
                        item.pdf_path = pdf_path
                    
                    item.status = DocumentStatus.OK
                except Exception as e:
                    logging.error(f"Failed to process {item.file_name}: {e}")
                    item.status = DocumentStatus.ERROR
            
            progress_callback(total, total)

    def merge_documents(self, settings: ToolSettings, progress_callback: Callable[[int, int], None]):
        """Слияние выбранных документов."""
        selected_items = [item for item in self.state.documents if item.is_selected]
        if not selected_items:
            return

        with WordApp() as word:
            merge_service = WordMergeService(word, settings)
            pdf_service = PdfExportService(word, settings)
            
            progress_callback(0, 100)
            try:
                # 1. Слияние в Word
                merged_path = merge_service.merge_documents(selected_items)
                
                # 2. Экспорт результата в PDF (если нужно)
                if settings.pdf.export_merged_document and merged_path:
                    pdf_service.export_processed_copy(merged_path)
                
                # 3. Объединение PDF (если нужно)
                if settings.pdf.merge_generated_pdfs:
                    pdf_merge_service = PdfMergeService(settings)
                    pdf_merge_service.merge_pdfs(selected_items)
                    
            except Exception as e:
                logging.error(f"Merge failed: {e}")
            finally:
                progress_callback(100, 100)
