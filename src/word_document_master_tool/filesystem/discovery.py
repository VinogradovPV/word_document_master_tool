import os
from datetime import datetime

from ..core.models import DocumentItem, SourceKind

WORD_EXTENSIONS = {".docx", ".docm", ".doc", ".rtf"}
EXCEL_EXTENSIONS = {".xls", ".xlsx"}


class DocumentDiscovery:
    """Класс-обертка для поиска документов в файлах."""

    def find_documents(self, folder_path: str) -> list[DocumentItem]:
        return find_source_documents(folder_path)


def find_word_documents(folder_path: str) -> list[DocumentItem]:
    """
    Находит поддерживаемые документы Word в указанной папке.
    Игнорирует временные файлы Word (начинающиеся с ~$ ).
    """
    return _find_documents(folder_path, WORD_EXTENSIONS)


def find_source_documents(folder_path: str) -> list[DocumentItem]:
    """
    Находит поддерживаемые исходники Word и Excel.
    Игнорирует временные Office-файлы (начинающиеся с ~$).
    """
    return _find_documents(folder_path, WORD_EXTENSIONS | EXCEL_EXTENSIONS)


def _find_documents(folder_path: str, supported_extensions: set[str]) -> list[DocumentItem]:
    items = []

    if not os.path.isdir(folder_path):
        return items

    files = sorted(os.listdir(folder_path))
    order_index = 1

    for file_name in files:
        if file_name.startswith("~$"):
            continue

        ext = os.path.splitext(file_name)[1].lower()
        if ext in supported_extensions:
            file_path = os.path.join(folder_path, file_name)
            stats = os.stat(file_path)
            source_kind = SourceKind.EXCEL if ext in EXCEL_EXTENSIONS else SourceKind.WORD

            items.append(
                DocumentItem(
                    order_index=order_index,
                    file_path=file_path,
                    file_name=file_name,
                    extension=ext[1:],
                    size_bytes=stats.st_size,
                    modified_at=datetime.fromtimestamp(stats.st_mtime),
                    source_kind=source_kind,
                )
            )
            order_index += 1

    return items
