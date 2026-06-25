import logging
import os

from ..core.models import DocumentItem, DocumentStatus, ToolSettings
from ..filesystem.paths import build_safe_file_name, ensure_unique_file_path
from ..word.word_app import WordApp


class PdfExportService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def export_sources_as_is(self, items: list[DocumentItem], log_service=None):
        """
        Экспортирует каждый выбранный исходный документ в PDF без изменений.
        """
        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        with WordApp() as word:
            if not word.app:
                logging.error("Word is not available for PDF export.")
                return

            for item in items:
                if not item.is_selected:
                    continue

                try:
                    doc = word.open_document(item.file_path, read_only=True)
                    if not doc:
                        item.status = DocumentStatus.ERROR
                        item.error_message = "Не удалось открыть документ"
                        continue

                    pdf_name = self._build_pdf_name(item)
                    pdf_path = os.path.join(output_folder, pdf_name)
                    pdf_path = ensure_unique_file_path(pdf_path)

                    success = word.export_as_pdf(
                        doc, 
                        pdf_path, 
                        optimize_for_print=self.settings.pdf.optimize_for_print
                    )

                    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges = 0

                    if success:
                        item.pdf_path = pdf_path
                        if log_service:
                            log_service.log_item(item, message="PDF экспортирован (без изменений)")
                    else:
                        item.status = DocumentStatus.ERROR
                        item.error_message = "Ошибка экспорта в PDF"

                except Exception as e:
                    logging.error(f"Error exporting {item.file_name} to PDF: {e}")
                    item.status = DocumentStatus.ERROR
                    item.error_message = str(e)

    def _build_pdf_name(self, item: DocumentItem) -> str:
        """
        Формирует имя PDF файла согласно настройкам.
        """
        mode = self.settings.pdf.naming_mode
        base_name = os.path.splitext(item.file_name)[0]

        if mode == "Добавить порядковый номер":
            name = f"{item.order_index:03d}_{base_name}"
        elif mode == "Добавить диапазон страниц":
            if item.start_page_number > 0 and item.end_page_number > 0:
                name = f"{base_name}_pages_{item.start_page_number}-{item.end_page_number}"
            else:
                name = base_name
        elif mode == "Использовать имя итогового документа":
            name = self.settings.output_file_name or base_name
        else:
            name = base_name

        if item.part_id:
            name = f"{name}_{item.part_id}"

        return f"{build_safe_file_name(name)}.pdf"
