import logging
import os

from ..core.models import ToolSettings
from .word_app import WordApp


class WordSplitService:
    """
    Сервис для разделения документа Word на части на основе маркеров.
    """

    def __init__(self, word: WordApp, settings: ToolSettings):
        self.word = word
        self.settings = settings

    def split_by_markers(self, source_path: str):
        """
        Разделяет документ на части на основе маркеров WMT_START/WMT_END.
        """
        if not self.word.app:
            return

        abs_source = os.path.abspath(source_path)
        if not os.path.exists(abs_source):
            logging.error(f"Source file for split not found: {abs_source}")
            return

        try:
            main_doc = self.word.open_document(abs_source, read_only=True)
            if not main_doc:
                return

            # Ищем маркеры начала
            content = main_doc.Content
            find = content.Find
            find.ClearFormatting()
            find.Text = "WMT_START|*|*"
            find.MatchWildcards = True
            
            parts_created = 0
            while find.Execute():
                # Текущий найденный диапазон - это маркер начала
                start_rng = content.Duplicate
                marker_text = start_rng.Text.strip()
                
                # Ищем соответствующий маркер конца после начала
                end_search_rng = main_doc.Range(start_rng.End, main_doc.Content.End)
                end_find = end_search_rng.Find
                end_find.ClearFormatting()
                end_find.Text = "WMT_END"
                
                if end_find.Execute():
                    end_rng = end_search_rng.Duplicate
                    
                    # Создаем новый документ для части
                    new_doc = self.word.app.Documents.Add()
                    
                    # Копируем содержимое между маркерами
                    data_rng = main_doc.Range(start_rng.End, end_rng.Start)
                    data_rng.Copy()
                    new_doc.Content.Paste()
                    
                    # Парсим имя файла из маркера
                    # Формат: WMT_START|filename|index
                    try:
                        parts = marker_text.split("|")
                        orig_name = parts[1] if len(parts) > 1 else f"part_{parts_created+1}"
                        file_name = f"split_{orig_name}"
                    except Exception:
                        file_name = f"split_part_{parts_created+1}.docx"
                    
                    if not file_name.lower().endswith(".docx"):
                        file_name += ".docx"

                    output_path = os.path.join(self.settings.output_folder, file_name)
                    new_doc.SaveAs2(FileName=os.path.abspath(output_path))
                    new_doc.Close(SaveChanges=0)
                    
                    parts_created += 1
                    logging.info(f"Created split part: {file_name}")
                
                # Продолжаем поиск после текущего маркера начала
                content.Start = start_rng.End

            main_doc.Close(SaveChanges=0)
            logging.info(f"Splitting completed. Total parts: {parts_created}")

        except Exception as e:
            logging.error(f"Error during document splitting: {e}")
