import logging
import os
import shutil

from ..core.models import DocumentItem, DocumentStatus, SourceKind, ToolSettings
from .superscript_scanner import SuperscriptScanner
from .word_app import WordApp


class SourceProcessingService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings
        self.scanner = SuperscriptScanner()

    def process_copies(self, items: list[DocumentItem], log_service=None):
        """
        Создает обработанные копии документов (принятие правок, удаление комментариев).
        """
        if not os.path.exists(self.settings.output_folder):
            os.makedirs(self.settings.output_folder)

        with WordApp() as word:
            if not word.app:
                logging.error("Word is not available for source processing.")
                return

            for item in items:
                if not item.is_selected or item.source_kind != SourceKind.WORD:
                    continue

                try:
                    # 1. Создаем временную копию файла
                    temp_name = f"WMT_{item.file_name}"
                    temp_path = os.path.join(self.settings.output_folder, temp_name)
                    shutil.copy2(item.file_path, temp_path)
                    item.temp_path = temp_path

                    # 2. Открываем копию для редактирования
                    doc = word.open_document(temp_path, read_only=False)
                    if not doc:
                        item.status = DocumentStatus.ERROR
                        item.error_message = "Не удалось открыть временную копию"
                        continue

                    # 3. Применяем настройки обработки
                    if self.settings.source_processing.accept_revisions:
                        doc.Revisions.AcceptAll()

                    if self.settings.source_processing.disable_track_changes:
                        doc.TrackRevisions = False

                    if self.settings.source_processing.remove_comments:
                        # wdDeleteAllComments = 0
                        doc.DeleteAllComments()

                    # 4. Сканируем на надстрочные знаки
                    superscript_pages = self.scanner.scan_document(
                        doc, first_page_number=item.start_page_number or 1
                    )

                    # 5. Сохраняем и закрываем
                    doc.Save()

                    # 6. Экспорт в PDF если нужно
                    if self.settings.pdf.export_processed_copies:
                        pdf_name = os.path.splitext(temp_name)[0] + ".pdf"
                        pdf_path = os.path.join(self.settings.output_folder, pdf_name)
                        word.export_as_pdf(doc, pdf_path)
                        item.pdf_path = pdf_path

                    doc.Close(SaveChanges=-1)  # wdSaveChanges = -1

                    if log_service:
                        log_service.log_item(
                            item,
                            superscript_pages=superscript_pages,
                            message="Обработанная копия создана",
                        )

                except Exception as e:
                    logging.error(f"Error processing {item.file_name}: {e}")
                    item.status = DocumentStatus.ERROR
                    item.error_message = str(e)
                    if log_service:
                        log_service.log_item(item, message=f"Ошибка: {e}")
