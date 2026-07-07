import logging
from typing import Any

from ..core.models import ToolSettings


class PageNumberingService:
    """
    Сервис для управления нумерацией страниц в документе Word.
    """

    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def apply_numbering(self, doc: Any, start_number: int = 1):
        """
        Применяет настройки нумерации страниц к документу.
        """
        if not doc or not self.settings.page_numbering.enabled:
            return

        try:
            # 1. Настройка полей
            if self.settings.page_numbering.adjust_margins:
                self._apply_margins(doc)

            # 2. Очистка колонтитулов
            if self.settings.page_numbering.remove_headers_footers:
                self._clear_headers_footers(doc)

            # 3. Вставка номеров страниц
            # wdAlignPageNumberLeft = 0, Center = 1, Right = 2
            alignment_map = {"По левому краю": 0, "По центру": 1, "По правому краю": 2}
            alignment = alignment_map.get(self.settings.page_numbering.alignment, 1)

            for section in doc.Sections:
                # Определяем куда вставлять (Верх/Низ)
                # wdHeaderFooterPrimary = 1
                if self.settings.page_numbering.location == "Верхний колонтитул":
                    hf = section.Headers(1)
                else:
                    hf = section.Footers(1)

                # Вставляем номер страницы
                # PageNumbers.Add(PageNumberAlignment, FirstPage)
                page_nums = hf.PageNumbers
                page_nums.Add(PageNumberAlignment=alignment, FirstPage=True)

                # Настройка начального номера для первого раздела
                if section.Index == 1:
                    page_nums.RestartNumberingAtSection = True
                    page_nums.StartingNumber = start_number

                # Настройка шрифта номера страницы
                try:
                    hf_range = hf.Range
                    hf_range.Font.Name = self.settings.page_numbering.font_name
                    hf_range.Font.Size = self.settings.page_numbering.font_size
                except Exception as font_err:
                    logging.warning(f"Failed to set font for page numbers: {font_err}")

            logging.info(
                "Page numbering applied: "
                f"start={start_number}, font={self.settings.page_numbering.font_name}"
            )

        except Exception as e:
            logging.error(f"Failed to apply page numbering: {e}")

    def _apply_margins(self, doc: Any):
        """
        Применяет настройки полей к документу (см в пункты: 1 см = 28.35 пт).
        """
        cm_to_points = 28.35
        try:
            for section in doc.Sections:
                setup = section.PageSetup
                setup.TopMargin = self.settings.page_numbering.top_margin_cm * cm_to_points
                setup.BottomMargin = self.settings.page_numbering.bottom_margin_cm * cm_to_points
                setup.LeftMargin = self.settings.page_numbering.left_margin_cm * cm_to_points
                setup.RightMargin = self.settings.page_numbering.right_margin_cm * cm_to_points
        except Exception as e:
            logging.error(f"Failed to apply margins: {e}")

    def _clear_headers_footers(self, doc: Any):
        """
        Очищает все колонтитулы во всех разделах документа.
        """
        try:
            for section in doc.Sections:
                for i in range(1, 4):  # wdHeaderFooterPrimary, FirstPage, EvenPages
                    section.Headers(i).Range.Delete()
                    section.Footers(i).Range.Delete()
        except Exception as e:
            logging.error(f"Failed to clear headers/footers: {e}")
