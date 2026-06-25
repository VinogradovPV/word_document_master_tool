import logging

from ..core.models import ToolSettings


class PageNumberingService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def apply_numbering(self, doc, start_number: int = 1):
        """
        Применяет нумерацию страниц к документу Word.
        """
        if not self.settings.page_numbering.enabled:
            return

        try:
            # 1. Настройка полей если нужно
            if self.settings.page_numbering.adjust_margins:
                self._apply_margins(doc)

            # 2. Очистка существующих колонтитулов если нужно
            if self.settings.page_numbering.remove_headers_footers:
                self._clear_headers_footers(doc)

            # 3. Вставка номеров страниц
            # wdHeaderFooterPrimary = 1
            # wdHeaderFooterEvenPages = 3
            # wdHeaderFooterFirstPage = 2

            for section in doc.Sections:
                # Определяем куда вставлять (Верх/Низ)
                if self.settings.page_numbering.location == "Верхний колонтитул":
                    hf = section.Headers(1)
                else:
                    hf = section.Footers(1)

                # Настройка выравнивания (wdAlignPageNumberCenter = 1, Left = 0, Right = 2)
                alignment = 1
                if self.settings.page_numbering.alignment == "По левому краю":
                    alignment = 0
                elif self.settings.page_numbering.alignment == "По правому краю":
                    alignment = 2

                # Вставляем номер страницы
                # PageNumbers.Add(PageNumberAlignment, FirstPage)
                hf.PageNumbers.Add(PageNumberAlignment=alignment, FirstPage=True)

                # Настройка начального номера для первого раздела
                if section.Index == 1:
                    hf.PageNumbers.RestartNumberingAtSection = True
                    hf.PageNumbers.StartingNumber = start_number

            logging.info(f"Page numbering applied starting from {start_number}")

        except Exception as e:
            logging.error(f"Error applying page numbering: {e}")

    def _apply_margins(self, doc):
        """
        Применяет настройки полей к документу.
        """
        # Конвертация см в пункты (1 см = 28.35 пунктов)
        cm_to_points = 28.35
        for section in doc.Sections:
            section.PageSetup.TopMargin = self.settings.page_numbering.top_margin_cm * cm_to_points
            section.PageSetup.BottomMargin = (
                self.settings.page_numbering.bottom_margin_cm * cm_to_points
            )
            section.PageSetup.LeftMargin = (
                self.settings.page_numbering.left_margin_cm * cm_to_points
            )
            section.PageSetup.RightMargin = (
                self.settings.page_numbering.right_margin_cm * cm_to_points
            )

    def _clear_headers_footers(self, doc):
        """
        Очищает все колонтитулы в документе.
        """
        for section in doc.Sections:
            for i in range(1, 4):  # Primary, FirstPage, EvenPages
                section.Headers(i).Range.Delete()
                section.Footers(i).Range.Delete()
