import contextlib
import logging
import os
import platform
import time

try:
    if platform.system() == "Windows":
        import pythoncom
        import win32com.client
    else:
        win32com = None
        pythoncom = None
except ImportError:
    win32com = None
    pythoncom = None

# Константы Word
WD_DO_NOT_SAVE_CHANGES = 0
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_OPTIMIZE_FOR_SCREEN = 1
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1


class WordApp:
    """
    Безопасная обертка для Microsoft Word COM-автоматизацию.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        self.app = None
        self._owns_app = False
        self._visible = visible
        self._display_alerts = display_alerts
        self._old_display_alerts = None

    def __enter__(self):
        if win32com is None:
            logging.warning(
                "Word COM automation is not available "
                "(not Windows or pywin32 not installed)."
            )
            return self

        try:
            pythoncom.CoInitialize()
            # Используем DispatchEx для создания отдельного экземпляра
            self.app = win32com.client.DispatchEx("Word.Application")
            self._owns_app = True
            
            self.app.Visible = self._visible
            self._old_display_alerts = self.app.DisplayAlerts
            # wdAlertsNone = 0
            self.app.DisplayAlerts = 0 if not self._display_alerts else -1
            
            return self
        except Exception as e:
            logging.error(f"Failed to start Word: {e}")
            self.app = None
            return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.app is not None:
            try:
                # Закрываем все открытые документы без сохранения
                while self.app.Documents.Count > 0:
                    try:
                        self.app.Documents.Item(1).Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                    except Exception:
                        break
                
                if self._old_display_alerts is not None:
                    self.app.DisplayAlerts = self._old_display_alerts
            except Exception as e:
                logging.error(f"Error during Word cleanup: {e}")
            finally:
                if self._owns_app:
                    with contextlib.suppress(Exception):
                        self.app.Quit()
                self.app = None
        
        if pythoncom is not None:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    def open_document(self, file_path: str, read_only: bool = True, retries: int = 3):
        """
        Безопасно открывает документ с поддержкой повторных попыток.
        """
        if not self.app:
            return None
        
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            logging.error(f"File not found: {abs_path}")
            return None

        for attempt in range(retries):
            try:
                doc = self.app.Documents.Open(
                    FileName=abs_path, 
                    ReadOnly=read_only,
                    AddToRecentFiles=False,
                    ConfirmConversions=False,
                    Visible=self._visible
                )
                return doc
            except Exception as e:
                logging.warning(f"Attempt {attempt + 1} to open {file_path} failed: {e}")
                if attempt < retries - 1:
                    time.sleep(1)
                else:
                    logging.error(f"Failed to open document after {retries} attempts: {file_path}")
        
        return None

    def export_as_pdf(
        self, 
        doc, 
        output_path: str, 
        optimize_for_print: bool = True,
        pdf_a: bool = False,
        include_properties: bool = True
    ) -> bool:
        """
        Экспортирует документ в PDF.
        """
        if not doc:
            return False
            
        try:
            abs_output = os.path.abspath(output_path)
            optimize = (
                WD_EXPORT_OPTIMIZE_FOR_PRINT
                if optimize_for_print
                else WD_EXPORT_OPTIMIZE_FOR_SCREEN
            )
            
            doc.ExportAsFixedFormat(
                OutputFileName=abs_output,
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=optimize,
                DocStructureTags=include_properties,
                BitmapMissingFonts=True,
                UseISO19005_1=pdf_a,
                CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS
            )
            return True
        except Exception as e:
            logging.error(f"Failed to export PDF: {e}")
            return False
