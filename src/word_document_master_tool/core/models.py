import os
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum


class DocumentStatus(Enum):
    OK = "OK"
    SKIPPED = "SKIPPED"
    ERROR = "ERROR"


class SourceKind(Enum):
    WORD = "word"
    EXCEL = "excel"


@dataclass
class DocumentItem:
    order_index: int
    file_path: str
    file_name: str
    extension: str
    size_bytes: int
    modified_at: datetime
    source_kind: SourceKind = SourceKind.WORD
    is_selected: bool = True
    is_valid: bool = True
    status: DocumentStatus = DocumentStatus.OK
    error_message: str = ""
    part_id: str = ""
    temp_path: str = ""
    pdf_path: str = ""
    start_page_number: int = 0
    end_page_number: int = 0
    page_count: int = 0
    start_footnote_number: int = 0
    end_footnote_number: int = 0
    footnote_count: int = 0


@dataclass
class MergeSettings:
    mode: int = 3  # 1: Next line, 2: Empty paragraph, 3: New page, 4: Section next page
    output_format: str = "docx"
    open_after_merge: bool = True
    create_report: bool = True
    create_backup: bool = True


@dataclass
class SourceProcessingSettings:
    accept_revisions: bool = True
    disable_track_changes: bool = True
    remove_comments: bool = False
    warn_protected_docs: bool = True


@dataclass
class PageNumberingSettings:
    enabled: bool = False
    start_number: int = 1
    scope: str = "Нумеровать обработанные копии документов"
    location: str = "Верхний колонтитул"
    alignment: str = "По центру"
    format: str = "1, 2, 3"
    continuous: bool = True
    restart_each_document: bool = False
    remove_existing: bool = True
    remove_headers_footers: bool = False
    preserve_headers_footers: bool = True
    font_name: str = "Times New Roman"
    font_size: float = 12.0
    adjust_margins: bool = True
    top_margin_cm: float = 2.0
    bottom_margin_cm: float = 2.0
    left_margin_cm: float = 2.0
    right_margin_cm: float = 2.0


@dataclass
class FootnoteSettings:
    enabled: bool = False
    scope: str = "Нумеровать итоговый документ"
    start_number: int = 1
    mode: str = "Сквозная по всему документу"
    format: str = "1, 2, 3"
    continuous: bool = True
    restart_each_document: bool = False
    restart_each_section: bool = False
    preserve_text: bool = True
    process_endnotes: bool = False
    update_fields: bool = True
    do_not_replace_numbers_with_text: bool = True


@dataclass
class PdfSettings:
    export_merged: bool = False
    export_sources: bool = False
    export_excel_sources: bool = False
    export_processed_copies: bool = False
    merge_generated_pdfs: bool = False
    output_folder: str = ""
    naming_mode: str = "Как исходный файл"
    quality: str = "Печать"
    open_after_export: bool = False
    pdf_a: bool = False
    include_properties: bool = True
    optimize_for_print: bool = True
    optimize_for_screen: bool = False


@dataclass
class MarkerSettings:
    use_markers: bool = False
    visibility: int = 1  # 1: Invisible, 2: Subtle, 3: Explicit
    removal_mode: int = 1  # 1: Visual only, 2: All
    backup_before_removal: bool = True


@dataclass
class ToolSettings:
    source_folder: str = ""
    output_folder: str = ""
    output_file_name: str = ""
    merge: MergeSettings = field(default_factory=MergeSettings)
    source_processing: SourceProcessingSettings = field(default_factory=SourceProcessingSettings)
    page_numbering: PageNumberingSettings = field(default_factory=PageNumberingSettings)
    footnotes: FootnoteSettings = field(default_factory=FootnoteSettings)
    pdf: PdfSettings = field(default_factory=PdfSettings)
    markers: MarkerSettings = field(default_factory=MarkerSettings)

    def validate_errors(self) -> list[str]:
        """
        Проверяет настройки на критические ошибки.
        """
        errors = []

        # 1. Проверка папок
        if not self.output_folder:
            errors.append("Папка для сохранения результатов не указана.")

        if self.source_folder and self.output_folder:
            src = os.path.abspath(self.source_folder)
            out = os.path.abspath(self.output_folder)
            if src == out:
                errors.append("Исходная папка и папка результатов не могут совпадать.")

        # 2. Нумерация страниц
        if self.page_numbering.enabled:
            if self.page_numbering.start_number < 1:
                errors.append("Начальный номер страницы должен быть >= 1.")
            if self.page_numbering.font_size <= 0:
                errors.append("Размер шрифта нумерации должен быть больше 0.")
            if any(
                m < 0
                for m in [
                    self.page_numbering.top_margin_cm,
                    self.page_numbering.bottom_margin_cm,
                    self.page_numbering.left_margin_cm,
                    self.page_numbering.right_margin_cm,
                ]
            ):
                errors.append("Поля страницы не могут быть отрицательными.")
            if self.page_numbering.continuous and self.page_numbering.restart_each_document:
                errors.append("Сквозная нумерация страниц и сброс в каждом документе несовместимы.")
            if (
                self.page_numbering.remove_headers_footers
                and self.page_numbering.preserve_headers_footers
            ):
                errors.append("Очистка и сохранение колонтитулов несовместимы.")

        # 3. Сноски
        if self.footnotes.enabled:
            if self.footnotes.start_number < 1:
                errors.append("Начальный номер сносок должен быть >= 1.")
            if self.footnotes.continuous and (
                self.footnotes.restart_each_document or self.footnotes.restart_each_section
            ):
                errors.append(
                    "Сквозная нумерация сносок несовместима со сбросом по документам/секциям."
                )

        # 4. PDF
        if self.pdf.merge_generated_pdfs and not (
            self.pdf.export_sources
            or self.pdf.export_excel_sources
            or self.pdf.export_processed_copies
            or self.pdf.export_merged
        ):
            errors.append("Слияние PDF невозможно без включенного экспорта отдельных PDF.")

        if self.pdf.optimize_for_print and self.pdf.optimize_for_screen:
            errors.append("Выберите только один режим оптимизации PDF (печать или экран).")

        return errors

    def validate_warnings(self) -> list[str]:
        """
        Проверяет настройки на предупреждения.
        """
        warnings = []
        if self.pdf.export_processed_copies and not any(
            [
                self.source_processing.accept_revisions,
                self.source_processing.remove_comments,
                self.page_numbering.enabled,
                self.footnotes.enabled,
            ]
        ):
            warnings.append("Экспорт обработанных копий включен, но изменения в них не вносятся.")

        if not self.output_file_name and self.pdf.export_merged:
            warnings.append("Имя итогового файла не указано, будет использовано имя по умолчанию.")

        return warnings

    def validate(self) -> list[str]:
        """Для обратной совместимости."""
        return self.validate_errors()
