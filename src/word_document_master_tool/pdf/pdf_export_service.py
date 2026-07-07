import contextlib
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

    def _get_output_path(self, source_name: str, suffix: str = "") -> str:
        """
        Формирует путь к PDF-файлу на основе имени и настроек.
        """
        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        base_name = os.path.splitext(source_name)[0]
        pdf_name = f"{base_name}{suffix}.pdf"
        return os.path.join(output_folder, pdf_name)

    def export_source_to_pdf(self, item: DocumentItem, word: WordApp) -> bool:
        """
        Экспортирует исходный документ в PDF без изменений.
        """
        if not item.is_selected:
            return False

        pdf_path = self._get_output_path(item.file_name, suffix="_source")

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
                with contextlib.suppress(Exception):
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)

    def export_processed_copy(self, file_path: str, word: WordApp) -> str | None:
        """
        Экспортирует уже обработанный (или объединенный) документ в PDF.
        """
        pdf_path = self._get_output_path(os.path.basename(file_path))

        doc = None
        try:
            doc = word.open_document(file_path, read_only=True)
            if not doc:
                return None

            success = word.export_as_pdf(
                doc,
                pdf_path,
                optimize_for_print=self.settings.pdf.optimize_for_print,
                pdf_a=self.settings.pdf.pdf_a,
                include_properties=self.settings.pdf.include_properties,
            )
            return pdf_path if success else None
        finally:
            if doc is not None:
                with contextlib.suppress(Exception):
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
