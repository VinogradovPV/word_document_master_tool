import logging
import os

from ..core.models import ToolSettings
from .word_app import WordApp


class SplitService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def split_by_markers(self, source_path: str):
        """
        Разделяет документ на части на основе маркеров WMT_START/WMT_END.
        """
        if not os.path.exists(source_path):
            logging.error(f"Source file for split not found: {source_path}")
            return

        with WordApp() as word:
            if not word.app:
                return

            try:
                main_doc = word.open_document(source_path, read_only=True)
                
                # Логика разделения:
                # 1. Найти WMT_START
                # 2. Найти соответствующий WMT_END
                # 3. Скопировать диапазон в новый документ
                # 4. Сохранить под именем из маркера
                
                # Упрощенная реализация для структуры:
                content = main_doc.Content
                find = content.Find
                find.ClearFormatting()
                find.Text = "WMT_START|*|*"
                find.MatchWildcards = True
                
                while find.Execute():
                    start_rng = content.Duplicate
                    marker_text = start_rng.Text
                    
                    # Ищем конец
                    end_find = main_doc.Content.Duplicate
                    end_find.Start = start_rng.End
                    find_end = end_find.Find
                    find_end.Text = "WMT_END"
                    
                    if find_end.Execute():
                        end_rng = end_find.Duplicate
                        
                        # Создаем новый документ
                        new_doc = word.app.Documents.Add()
                        
                        # Копируем содержимое между маркерами
                        # (исключая сами маркеры)
                        data_rng = main_doc.Range(start_rng.End, end_rng.Start)
                        data_rng.Copy()
                        new_doc.Content.Paste()
                        
                        # Парсим имя из маркера
                        parts = marker_text.split("|")
                        file_name = parts[1] if len(parts) > 1 else "split_part"
                        
                        output_path = os.path.join(
                            self.settings.output_folder, 
                            f"split_{file_name}"
                        )
                        new_doc.SaveAs2(FileName=os.path.abspath(output_path))
                        new_doc.Close()
                        
                    # Продолжаем поиск после текущего маркера начала
                    content.Start = start_rng.End

                main_doc.Close(SaveChanges=0)

            except Exception as e:
                logging.error(f"Error splitting document: {e}")
