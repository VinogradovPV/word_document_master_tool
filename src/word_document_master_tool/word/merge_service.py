import logging
import os

from ..core.models import DocumentItem, SourceKind, ToolSettings
from ..filesystem.paths import build_safe_file_name, ensure_unique_file_path
from .word_app import WordApp


class WordMergeService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def merge_documents(self, items: list[DocumentItem], log_service=None) -> str:
        """
        Объединяет выбранные документы Word в один новый документ.
        """
        selected_items = [
            item for item in items if item.is_selected and item.source_kind == SourceKind.WORD
        ]
        if not selected_items:
            logging.warning("No documents selected for merge.")
            return ""

        with WordApp() as word:
            if not word.app:
                logging.error("Word is not available for merge.")
                return ""

            try:
                # Создаем новый пустой документ
                merged_doc = word.app.Documents.Add()

                for i, item in enumerate(selected_items):
                    try:
                        # Вставляем содержимое файла
                        # Range(0, 0) - начало документа, но мы вставляем в конец
                        # Поэтому берем Range конца документа
                        end_pos = merged_doc.Content.End - 1
                        rng = merged_doc.Range(end_pos, end_pos)

                        # Вставляем файл
                        rng.InsertFile(FileName=os.path.abspath(item.file_path))

                        # Если это не последний документ, добавляем разрыв согласно настройкам
                        if i < len(selected_items) - 1:
                            self._apply_separator(merged_doc, self.settings.merge.mode)

                        if log_service:
                            log_service.log_info(f"Inserted: {item.file_name}")

                    except Exception as e:
                        logging.error(f"Error inserting {item.file_name}: {e}")
                        if log_service:
                            log_service.log_error(f"Failed to insert {item.file_name}: {e}")

                # Сохраняем итоговый документ
                output_path = self._save_merged_doc(merged_doc)

                if self.settings.merge.open_after_merge:
                    merged_doc.Activate()
                else:
                    merged_doc.Close(SaveChanges=0)

                return output_path

            except Exception as e:
                logging.error(f"Merge operation failed: {e}")
                return ""

    def _apply_separator(self, doc, mode: int):
        """
        Применяет разрыв между документами.
        1: Со следующей строки (Paragraph)
        2: Через один пустой абзац
        3: С новой страницы (Page Break)
        4: С разрывом раздела (Section Break Next Page)
        """
        end_pos = doc.Content.End - 1
        rng = doc.Range(end_pos, end_pos)

        if mode == 1:
            rng.InsertParagraphAfter()
        elif mode == 2:
            rng.InsertParagraphAfter()
            rng.InsertParagraphAfter()
        elif mode == 3:
            # wdPageBreak = 7
            rng.InsertBreak(7)
        elif mode == 4:
            # wdSectionBreakNextPage = 2
            rng.InsertBreak(2)
        else:
            rng.InsertBreak(7)

    def _save_merged_doc(self, doc) -> str:
        """
        Сохраняет объединенный документ согласно настройкам.
        """
        if not os.path.exists(self.settings.output_folder):
            os.makedirs(self.settings.output_folder)

        base_name = self.settings.output_file_name or "merged_document"
        ext = self.settings.merge.output_format.lower()

        full_name = f"{build_safe_file_name(base_name)}.{ext}"
        output_path = os.path.join(self.settings.output_folder, full_name)
        output_path = ensure_unique_file_path(output_path)

        # Сопоставление расширений с форматами Word
        formats = {
            "docx": 16,  # wdFormatXMLDocument
            "docm": 20,  # wdFormatXMLDocumentMacroEnabled
            "doc": 0,  # wdFormatDocument
            "rtf": 6,  # wdFormatRTF
            "pdf": 17,  # wdFormatPDF
        }

        file_format = formats.get(ext, 16)

        if ext == "pdf":
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(output_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
            )
        else:
            doc.SaveAs2(FileName=os.path.abspath(output_path), FileFormat=file_format)

        return output_path
