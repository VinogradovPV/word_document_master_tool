import os
from datetime import datetime

from ..core.models import DocumentItem


def find_word_documents(folder_path: str) -> list[DocumentItem]:
    """
    Находит поддерживаемые документы Word в указанной папке.
    Игнорирует временные файлы Word (начинающиеся с ~$ ).
    """
    supported_extensions = {".docx", ".docm", ".doc", ".rtf"}
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

            items.append(
                DocumentItem(
                    order_index=order_index,
                    file_path=file_path,
                    file_name=file_name,
                    extension=ext[1:],
                    size_bytes=stats.st_size,
                    modified_at=datetime.fromtimestamp(stats.st_mtime),
                )
            )
            order_index += 1

    return items
