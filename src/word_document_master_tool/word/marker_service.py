import logging
from typing import Any

from ..core.models import DocumentItem, ToolSettings


class MarkerService:
    """
    Сервис для управления текстовыми маркерами границ документов.
    """

    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def add_markers(self, doc: Any, item: DocumentItem):
        """
        Добавляет текстовые маркеры начала и конца документа.
        Формат: WMT_START|filename|index
        """
        if not self.settings.markers.use_markers:
            return

        try:
            # Маркер начала (используем \r для Word)
            start_marker = f"WMT_START|{item.file_name}|{item.order_index}"
            doc.Content.InsertBefore(f"{start_marker}\r")

            # Маркер конца
            end_marker = "WMT_END"
            doc.Content.InsertAfter(f"\r{end_marker}")

            logging.info(f"Markers added to {item.file_name}")

        except Exception as e:
            logging.error(f"Error adding markers to {item.file_name}: {e}")

    def remove_markers(self, doc: Any):
        """
        Удаляет все маркеры WMT из документа через поиск и замену.
        """
        try:
            find = doc.Content.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Forward = True
            find.Wrap = 1  # wdFindContinue
            find.Format = False
            find.MatchWildcards = True

            # 1. Удаляем WMT_START|... включая перевод строки (^13)
            find.Text = "WMT_START|*^13"
            find.Replacement.Text = ""
            find.Execute(Replace=2)  # wdReplaceAll

            # 2. Удаляем WMT_END включая перевод строки
            find.Text = "WMT_END^13"
            find.Execute(Replace=2)

            # 3. На случай если маркеры в конце документа без ^13
            find.Text = "WMT_START|*"
            find.Execute(Replace=2)
            find.Text = "WMT_END"
            find.Execute(Replace=2)

            logging.info("WMT markers removed from document.")

        except Exception as e:
            logging.error(f"Error removing markers: {e}")
