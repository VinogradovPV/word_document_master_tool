import csv
import logging
import os
import threading

from .models import DocumentItem


class ProcessingLogger:
    """
    Сервис для ведения табличного TSV-лога обработки документов.
    Потокобезопасен для использования в фоновых задачах.
    """

    def __init__(self, output_folder: str, use_excel_friendly_encoding: bool = True):
        self.output_folder = output_folder
        # UTF-8 with BOM для корректного открытия в Excel (по умолчанию True)
        self.encoding = "utf-8-sig" if use_excel_friendly_encoding else "utf-8"
        self.log_path = os.path.join(output_folder, "processing_log.tsv")
        self._lock = threading.Lock()
        self._initialize_log()

    def _initialize_log(self):
        """Создает файл лога с заголовками."""
        try:
            if not os.path.exists(self.output_folder):
                os.makedirs(self.output_folder, exist_ok=True)
            
            headers = [
                "Файл", 
                "Страницы", 
                "Сноски", 
                "Надстрочные символы, стр.", 
                "Статус", 
                "Путь результата", 
                "Сообщение"
            ]
            
            with self._lock, open(
                self.log_path, "w", encoding=self.encoding, newline=""
            ) as f:
                writer = csv.writer(f, delimiter="\t")
                writer.writerow(headers)
                    
            logging.info(f"TSV report initialized at: {self.log_path}")
        except Exception as e:
            logging.error(f"Failed to initialize TSV log: {e}")

    def log_item(
        self, 
        item: DocumentItem, 
        superscript_pages: str = "-", 
        message: str | None = None
    ):
        """
        Записывает строку в TSV-лог.
        """
        try:
            # Форматирование диапазонов
            pages = (
                f"{item.start_page_number}-{item.end_page_number}"
                if item.page_count > 0
                else "-"
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
            
            with self._lock, open(
                self.log_path, "a", encoding=self.encoding, newline=""
            ) as f:
                writer = csv.writer(f, delimiter="\t")
                writer.writerow(row)
        except Exception as e:
            logging.error(f"Failed to write to TSV log: {e}")


def setup_application_logging(output_folder: str, log_level: int = logging.INFO):
    """
    Настраивает технический лог приложения (application.log).
    """
    try:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)
            
        log_path = os.path.join(output_folder, "application.log")
        
        # Очистка существующих хендлеров для предотвращения дублирования
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
            
        logging.basicConfig(
            level=log_level,
            format="%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s",
            handlers=[
                logging.FileHandler(log_path, encoding="utf-8"),
                logging.StreamHandler()
            ]
        )
        logging.info("Application logging initialized.")
    except Exception as e:
        print(f"CRITICAL: Failed to setup logging: {e}")
