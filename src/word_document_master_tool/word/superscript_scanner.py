import logging


class SuperscriptScanner:
    """
    Сканирует документ на наличие надстрочных знаков и определяет номера страниц.
    """

    def scan_document(self, doc, first_page_number: int = 1) -> str:
        """
        Сканирует основной текст документа.
        Возвращает строку с номерами страниц через запятую или '-' если не найдено.
        """
        if not doc:
            return "ERROR"

        try:
            # Пересчитываем страницы
            doc.Repaginate()

            pages = set()
            # wdMainTextStory = 1
            search_range = doc.StoryRanges(1).Duplicate

            find = search_range.Find
            find.ClearFormatting()
            find.Text = ""
            find.Forward = True
            find.Wrap = 0  # wdFindStop = 0
            find.Format = True
            find.Font.Superscript = True

            while find.Execute():
                # Проверяем, что текст не пустой
                text = search_range.Text.strip().replace("\r", "").replace("\n", "")
                if text:
                    # wdActiveEndAdjustedPageNumber = 1
                    # wdActiveEndPageNumber = 3
                    try:
                        page = int(search_range.Information(1))
                    except Exception:
                        page = int(search_range.Information(3))

                    if page > 0:
                        # Корректируем номер страницы согласно настройкам начала нумерации
                        display_page = first_page_number + page - 1
                        pages.add(display_page)

                # Схлопываем диапазон к концу, чтобы искать дальше
                search_range.Collapse(0)  # wdCollapseEnd = 0

            if not pages:
                return "-"

            sorted_pages = sorted(list(pages))
            return ", ".join(map(str, sorted_pages))

        except Exception as e:
            logging.error(f"Error scanning superscript: {e}")
            return "ERROR"
