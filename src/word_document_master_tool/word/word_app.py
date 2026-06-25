import logging
import os
import platform

# Пытаемся импортировать win32com только на Windows
if platform.system() == "Windows":
    import win32com.client
    from win32com.client import constants
else:
    win32com = None
    constants = None


class WordApp:
    """
    Обертка для работы с Microsoft Word через COM-автоматизацию.
    """

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.app = None
        self._old_display_alerts = None

    def __enter__(self):
        if platform.system() != "Windows":
            logging.warning("Word COM automation is only available on Windows.")
            return self

        try:
            # Пытаемся подключиться к существующему экземпляру или создать новый
            self.app = win32com.client.Dispatch("Word.Application")
            self.app.Visible = self.visible
            
            # Сохраняем текущее состояние алертов и отключаем их
            self._old_display_alerts = self.app.DisplayAlerts
            self.app.DisplayAlerts = 0  # wdAlertsNone
            
            return self
        except Exception as e:
            logging.error(f"Failed to start Word: {e}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.app:
            try:
                # Восстанавливаем состояние алертов
                if self._old_display_alerts is not None:
                    self.app.DisplayAlerts = self._old_display_alerts
                
                # Если мы создали приложение, закрываем его
                # В реальном приложении может потребоваться более сложная логика
                # чтобы не закрывать Word, если он был открыт пользователем
                self.app.Quit()
            except Exception as e:
                logging.error(f"Error while closing Word: {e}")
            finally:
                self.app = None

    def open_document(self, file_path: str, read_only: bool = True):
        """
        Открывает документ Word.
        """
        if not self.app:
            return None
        
        return self.app.Documents.Open(
            FileName=os.path.abspath(file_path),
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=self.visible
        )

    def export_as_pdf(self, doc, pdf_path: str, optimize_for_print: bool = True):
        """
        Экспортирует документ в PDF.
        """
        if not self.app or not doc:
            return False

        try:
            # wdExportFormatPDF = 17
            # wdExportOptimizeForPrint = 0
            # wdExportOptimizeForOnScreen = 1
            optimize_for = 0 if optimize_for_print else 1
            
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=optimize_for,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
            return True
        except Exception as e:
            logging.error(f"Failed to export PDF: {e}")
            return False
