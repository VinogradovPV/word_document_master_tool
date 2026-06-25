import logging

from ..core.models import ToolSettings


class FootnoteService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def apply_numbering(self, doc, start_number: int = 1):
        """
        Применяет нумерацию сносок к документу Word.
        """
        if not self.settings.footnotes.enabled:
            return

        try:
            # Настройка сносок (Footnotes)
            # wdRestartContinuous = 0
            # wdRestartSection = 1
            # wdRestartPage = 2
            
            restart_mode = 0
            if self.settings.footnotes.restart_each_document:
                restart_mode = 0  # Для отдельных документов это всегда Continuous от начала
            elif self.settings.footnotes.restart_each_section:
                restart_mode = 1

            doc.Footnotes.NumberingRule = restart_mode
            doc.Footnotes.StartingNumber = start_number
            
            # Настройка формата (wdNoteNumberStyleArabic = 0)
            doc.Footnotes.NumberStyle = 0
            
            # Если нужно обрабатывать концевые сноски (Endnotes)
            if self.settings.footnotes.process_endnotes:
                doc.Endnotes.NumberingRule = restart_mode
                doc.Endnotes.StartingNumber = start_number
                doc.Endnotes.NumberStyle = 2  # wdNoteNumberStyleUppercaseRoman = 2

            logging.info(f"Footnote numbering applied starting from {start_number}")

        except Exception as e:
            logging.error(f"Error applying footnote numbering: {e}")

    def get_footnote_count(self, doc) -> int:
        """
        Возвращает количество сносок в документе.
        """
        try:
            return doc.Footnotes.Count
        except Exception:
            return 0
