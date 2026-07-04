import logging
from typing import Any

from ..core.models import ToolSettings


class FootnoteService:
    """
    Сервис для управления нумерацией сносок в документе Word.
    """

    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def apply_numbering(self, doc: Any, start_number: int = 1):
        """
        Применяет настройки нумерации сносок к документу.
        """
        if not doc or not self.settings.footnotes.enabled:
            return

        try:
            # wdRestartContinuous = 0 (сквозная)
            # wdRestartSection = 1 (в каждом разделе)
            # wdRestartPage = 2 (на каждой странице)
            
            restart_mode = 0  # По умолчанию сквозная
            if self.settings.footnotes.restart_each_section:
                restart_mode = 1
                
            # Настройка для всего документа
            doc.Footnotes.NumberingRule = restart_mode
            doc.Footnotes.StartingNumber = start_number
            
            # Настройка формата (wdNoteNumberStyleArabic = 0)
            doc.Footnotes.NumberStyle = 0
            
            # Обработка концевых сносок (Endnotes) если включено
            if self.settings.footnotes.process_endnotes:
                doc.Endnotes.NumberingRule = restart_mode
                doc.Endnotes.StartingNumber = start_number
                # wdNoteNumberStyleLowercaseRoman = 2
                doc.Endnotes.NumberStyle = 2
            
            logging.info(f"Footnote numbering applied: start={start_number}, restart={restart_mode}")
            
        except Exception as e:
            logging.error(f"Failed to apply footnote numbering: {e}")

    def get_footnote_count(self, doc: Any) -> int:
        """
        Возвращает количество сносок в документе.
        """
        try:
            return doc.Footnotes.Count
        except Exception:
            return 0
