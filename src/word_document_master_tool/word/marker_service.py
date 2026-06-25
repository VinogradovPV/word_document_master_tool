import logging

from ..core.models import DocumentItem, ToolSettings


class MarkerService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def add_markers(self, doc, item: DocumentItem):
        """
        Добавляет текстовые маркеры начала и конца документа.
        """
        if not self.settings.markers.use_markers:
            return

        try:
            # Маркер начала
            start_marker = f"WMT_START|{item.file_name}|{item.order_index}"
            doc.Content.InsertBefore(f"{start_marker}\n")
            
            # Маркер конца
            end_marker = "WMT_END"
            doc.Content.InsertAfter(f"\n{end_marker}")
            
            logging.info(f"Markers added to {item.file_name}")

        except Exception as e:
            logging.error(f"Error adding markers: {e}")

    def remove_markers(self, doc):
        """
        Удаляет все маркеры WMT из документа.
        """
        try:
            find = doc.Content.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = "WMT_START|*^13"
            find.Replacement.Text = ""
            find.Forward = True
            find.Wrap = 1  # wdFindContinue = 1
            find.Format = False
            find.MatchWildcards = True
            find.Execute(Replace=2)  # wdReplaceAll = 2
            
            find.Text = "WMT_END^13"
            find.Execute(Replace=2)
            
            # На случай если маркеры без абзаца
            find.Text = "WMT_START|*"
            find.Execute(Replace=2)
            find.Text = "WMT_END"
            find.Execute(Replace=2)

        except Exception as e:
            logging.error(f"Error removing markers: {e}")
