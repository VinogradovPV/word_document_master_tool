import logging
import os

from ..core.models import DocumentItem, DocumentStatus, ToolSettings
from ..word.word_app import WD_DO_NOT_SAVE_CHANGES, WordApp


class PdfExportService:
    """
    Сервис для экспорта документов Word в PDF.
    """

    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def export_source_to_pdf(self, item: DocumentItem, word: WordApp) -> bool:
        """
        Экспортирует исходный документ в PDF без изменений.
        """
        if not item.is_selected:
            return False

        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        pdf_name = os.path.splitext(item.file_name)[0] + "_source.pdf"
        pdf_path = os.path.join(output_folder, pdf_name)

        doc = None
        try:
            # Открываем Read-Only
            doc = word.open_document(item.file_path, read_only=True)
            if doc is None:
                raise RuntimeError("Не удалось открыть документ в Word")

            # Экспорт
            success = word.export_as_pdf(
                doc,
                pdf_path,
                optimize_for_print=self.settings.pdf.optimize_for_print,
                pdf_a=self.settings.pdf.pdf_a,
                include_properties=self.settings.pdf.include_properties,
            )

            if not success:
                raise RuntimeError("ExportAsFixedFormat вернул ошибку")

            item.pdf_path = pdf_path
            item.status = DocumentStatus.OK
            logging.info(f"Успешный экспорт в PDF: {item.file_name}")
            return True

        except Exception as exc:
            item.status = DocumentStatus.ERROR
            item.error_message = str(exc)
            logging.error(f"Ошибка экспорта {item.file_name}: {exc}")
            return False

        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass
