import csv
import os
from datetime import datetime

from .models import DocumentItem


class LoggingService:
    def __init__(self, output_folder: str):
        self.output_folder = output_folder
        self.log_file = self._init_log_file()

    def _init_log_file(self) -> str:
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(self.output_folder, f"processing_log_{timestamp}.tsv")
        
        headers = [
            "Файл",
            "Страницы",
            "Сноски",
            "Надстрочные символы, стр.",
            "Статус",
            "Путь результата",
            "Сообщение"
        ]
        
        with open(log_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter="\t")
            writer.writerow(headers)
            
        return log_path

    def log_item(self, item: DocumentItem, superscript_pages: str = "-", message: str = ""):
        """
        Записывает информацию об обработке одного документа в TSV лог.
        """
        pages = (
            f"{item.start_page_number}-{item.end_page_number}" if item.page_count > 0 else "-"
        )
        footnotes = (
            f"{item.start_footnote_number}-{item.end_footnote_number}"
            if item.footnote_count > 0
            else "-"
        )
        
        row = [
            item.file_name,
            pages,
            footnotes,
            superscript_pages,
            item.status.value,
            item.pdf_path or item.temp_path or "",
            message or item.error_message
        ]
        
        with open(self.log_file, "a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter="\t")
            writer.writerow(row)
            
    def log_info(self, message: str):
        """
        Записывает общую информационную строку в лог.
        """
        with open(self.log_file, "a", encoding="utf-8", newline="") as f:
            f.write(f"# INFO: {message}\n")

    def log_error(self, message: str):
        """
        Записывает информацию об ошибке в лог.
        """
        with open(self.log_file, "a", encoding="utf-8", newline="") as f:
            f.write(f"# ERROR: {message}\n")
