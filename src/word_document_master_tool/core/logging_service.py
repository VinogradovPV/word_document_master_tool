import csv
import logging
import os

from .models import DocumentItem


class ProcessingLogger:
    """
    Сервис для ведения табличного TSV-лога обработки документов.
    """

    def __init__(self, output_folder: str, use_excel_friendly_encoding: bool = False):
        self.output_folder = output_folder
        self.encoding = "utf-8-sig" if use_excel_friendly_encoding else "utf-8"
        self.log_path = os.path.join(output_folder, "processing_log.tsv")
        self._initialize_log()

    def _initialize_log(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        
        headers = [
            "Файл", 
            "Страницы", 
            "Сноски", 
            "Надстрочные символы, стр.", 
            "Статус", 
            "Путь результата", 
            "Сообщение"
        ]
        with open(self.log_path, "w", encoding=self.encoding, newline="") as f:
            writer = csv.writer(f, delimiter="\t")
            writer.writerow(headers)

    def log_item(
        self, 
        item: DocumentItem, 
        superscript_pages: str = "-", 
        message: str = ""
    ):
        """
        Записывает строку в TSV-лог.
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
            item.pdf_path or item.temp_path or "-",
            message or item.error_message or "-"
        ]
        with open(self.log_path, "a", encoding=self.encoding, newline="") as f:
            writer = csv.writer(f, delimiter="\t")
            writer.writerow(row)


def setup_application_logging(output_folder: str):
    """
    Настраивает технический лог приложения.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    log_path = os.path.join(output_folder, "application.log")
    
    # Сброс существующих хендлеров
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
        
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    logging.info("Application logging initialized.")
