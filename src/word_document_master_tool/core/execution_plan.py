from dataclasses import dataclass, field

from .models import ToolSettings


@dataclass
class ExecutionPlan:
    selected_count: int
    creates_merged_word: bool
    creates_processed_copies: bool
    pdf_sources: bool
    pdf_processed_copies: bool
    pdf_merged_document: bool
    pdf_combined: bool
    operations: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def lines(self) -> list[str]:
        return [
            f"Документов выбрано: {self.selected_count}",
            f"Будет создан итоговый Word: {self._yes_no(self.creates_merged_word)}",
            f"Будут созданы обработанные копии: {self._yes_no(self.creates_processed_copies)}",
            f"PDF исходников без изменений: {self._yes_no(self.pdf_sources)}",
            f"PDF обработанных копий: {self._yes_no(self.pdf_processed_copies)}",
            f"PDF итогового документа: {self._yes_no(self.pdf_merged_document)}",
            f"Общий PDF из созданных PDF: {self._yes_no(self.pdf_combined)}",
            f"Предупреждения: {self._join_or_dash(self.warnings)}",
            f"Ошибки: {self._join_or_dash(self.errors)}",
        ]

    def summary_text(self) -> str:
        return "\n".join(self.lines())

    @staticmethod
    def _yes_no(value: bool) -> str:
        return "да" if value else "нет"

    @staticmethod
    def _join_or_dash(items: list[str]) -> str:
        return "; ".join(items) if items else "-"


def build_execution_plan(
    settings: ToolSettings,
    selected_count: int,
    *,
    word_available: bool | None = None,
) -> ExecutionPlan:
    creates_merged_word = settings.pdf.export_merged or bool(settings.output_file_name)
    creates_processed_copies = any(
        [
            settings.source_processing.accept_revisions,
            settings.source_processing.disable_track_changes,
            settings.source_processing.remove_comments,
            settings.page_numbering.enabled,
            settings.footnotes.enabled,
            settings.pdf.export_processed_copies,
        ]
    )
    plan = ExecutionPlan(
        selected_count=selected_count,
        creates_merged_word=creates_merged_word,
        creates_processed_copies=creates_processed_copies,
        pdf_sources=settings.pdf.export_sources,
        pdf_processed_copies=settings.pdf.export_processed_copies,
        pdf_merged_document=settings.pdf.export_merged,
        pdf_combined=settings.pdf.merge_generated_pdfs,
    )

    _append_operations(plan, settings)
    _append_errors(plan, settings, selected_count, word_available)
    _append_warnings(plan, settings)
    return plan


def _append_operations(plan: ExecutionPlan, settings: ToolSettings) -> None:
    if settings.pdf.export_sources:
        plan.operations.append("PDF исходников")
    if settings.pdf.export_processed_copies:
        plan.operations.append("PDF обработанных копий")
    if settings.pdf.export_merged:
        plan.operations.append("PDF итогового документа")
    if settings.pdf.merge_generated_pdfs:
        plan.operations.append("общий PDF")
    if plan.creates_processed_copies:
        plan.operations.append("обработка копий")
    if plan.creates_merged_word:
        plan.operations.append("итоговый Word")


def _append_errors(
    plan: ExecutionPlan,
    settings: ToolSettings,
    selected_count: int,
    word_available: bool | None,
) -> None:
    if not settings.source_folder:
        plan.errors.append("Не указана исходная папка.")
    if not settings.output_folder:
        plan.errors.append("Не указана папка результата.")
    if selected_count == 0:
        plan.errors.append("Не выбраны документы.")

    pdf_requested = any(
        [
            settings.pdf.export_sources,
            settings.pdf.export_processed_copies,
            settings.pdf.export_merged,
            settings.pdf.merge_generated_pdfs,
        ]
    )
    if pdf_requested and not settings.pdf.output_folder:
        plan.errors.append("Не указана папка PDF.")
    if settings.pdf.merge_generated_pdfs and not (
        settings.pdf.export_sources
        or settings.pdf.export_processed_copies
        or settings.pdf.export_merged
    ):
        plan.errors.append("Включён общий PDF, но не включено создание отдельных PDF.")

    plan.errors.extend(settings.validate_errors())

    word_required = any(
        [
            plan.creates_merged_word,
            plan.creates_processed_copies,
            settings.pdf.export_sources,
            settings.pdf.export_processed_copies,
            settings.pdf.export_merged,
        ]
    )
    if word_available is False and word_required:
        plan.errors.append("Microsoft Word COM недоступен.")


def _append_warnings(plan: ExecutionPlan, settings: ToolSettings) -> None:
    if settings.pdf.export_sources and (
        settings.source_processing.accept_revisions
        or settings.source_processing.disable_track_changes
        or settings.source_processing.remove_comments
        or settings.page_numbering.enabled
        or settings.footnotes.enabled
    ):
        plan.warnings.append(
            "PDF исходников без изменений игнорирует исправления, нумерацию и сноски."
        )
    if settings.pdf.export_processed_copies and not plan.creates_processed_copies:
        plan.warnings.append("PDF обработанных копий требует создания обработанных копий.")
    if settings.markers.removal_mode == 2:
        plan.warnings.append("Полное удаление маркеров может изменить документы.")
    plan.warnings.extend(settings.validate_warnings())
