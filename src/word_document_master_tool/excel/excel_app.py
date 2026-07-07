import contextlib
import logging
import os
import platform

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

XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


class ExcelApp:
    """
    Безопасная обертка для Microsoft Excel COM-автоматизации.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        self.app = None
        self._owns_app = False
        self._visible = visible
        self._display_alerts = display_alerts

    def __enter__(self):
        if platform.system() != "Windows":
            raise RuntimeError("Excel automation is available only on Windows")
        if win32com is None:
            raise RuntimeError("pywin32 is required for Excel automation")

        try:
            pythoncom.CoInitialize()
            self.app = win32com.client.DispatchEx("Excel.Application")
            self._owns_app = True
            self.app.Visible = self._visible
            self.app.DisplayAlerts = self._display_alerts
            return self
        except Exception as exc:
            self.app = None
            raise RuntimeError(f"Failed to start Excel: {exc}") from exc

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.app is not None:
            try:
                while self.app.Workbooks.Count > 0:
                    try:
                        self.app.Workbooks.Item(1).Close(SaveChanges=False)
                    except Exception:
                        break
            except Exception as exc:
                logging.error(f"Error during Excel cleanup: {exc}")
            finally:
                if self._owns_app:
                    with contextlib.suppress(Exception):
                        self.app.Quit()
                self.app = None

        if pythoncom is not None:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    def open_workbook(self, file_path: str, read_only: bool = True):
        if self.app is None:
            return None

        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            logging.error(f"Excel file not found: {abs_path}")
            return None

        return self.app.Workbooks.Open(
            Filename=abs_path,
            UpdateLinks=0,
            ReadOnly=read_only,
            AddToMru=False,
        )

    @staticmethod
    def export_workbook_as_pdf(
        workbook,
        output_path: str,
        *,
        include_properties: bool = True,
        ignore_print_areas: bool = False,
    ) -> bool:
        if workbook is None:
            return False

        workbook.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=os.path.abspath(output_path),
            Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=include_properties,
            IgnorePrintAreas=ignore_print_areas,
            OpenAfterPublish=False,
        )
        return True

    @staticmethod
    def close_workbook(workbook) -> None:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
