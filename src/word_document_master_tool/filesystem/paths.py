import os
import re


def build_safe_file_name(name: str) -> str:
    """
    Очищает имя файла от недопустимых символов.
    """
    # Заменяем недопустимые символы на подчеркивание
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return safe_name.strip()


def ensure_unique_file_path(file_path: str) -> str:
    """
    Если файл существует, добавляет числовой суффикс для обеспечения уникальности.
    """
    if not os.path.exists(file_path):
        return file_path

    directory, full_name = os.path.split(file_path)
    name, ext = os.path.splitext(full_name)

    counter = 1
    while True:
        new_name = f"{name}_{counter:03d}{ext}"
        new_path = os.path.join(directory, new_name)
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def can_write_to_folder(folder_path: str) -> bool:
    """
    Проверяет наличие прав на запись в папку.
    """
    if not os.path.isdir(folder_path):
        return False

    test_file = os.path.join(folder_path, ".write_test")
    try:
        with open(test_file, "w", encoding="utf-8") as f:
            f.write("test")
        os.remove(test_file)
        return True
    except OSError:
        return False
