import logging
import os
from dataclasses import dataclass, field

from pypdf import PdfWriter

from ..core.models import DocumentItem, ToolSettings
from ..filesystem.paths import build_safe_file_name, ensure_unique_file_path


@dataclass
class PdfMergeResult:
    """
    Результат операции слияния PDF.
    """

    success: bool
    output_path: str = ""
    merged_count: int = 0
    skipped_paths: list[str] = field(default_factory=list)
    error_message: str = ""


class PdfMergeService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def merge_pdfs(self, items: list[DocumentItem]) -> PdfMergeResult:
        """
        Объединяет сгенерированные PDF файлы в один.
        """
        selected_items = [item for item in items if item.is_selected]
        pdf_paths = []
        skipped_paths = []
        for item in selected_items:
            if item.pdf_path:
                pdf_paths.append(item.pdf_path)
            else:
                skipped_paths.append(item.file_path)

        if not pdf_paths:
            return PdfMergeResult(
                success=False, skipped_paths=skipped_paths, error_message="Нет PDF для объединения"
            )

        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        base_name = self.settings.output_file_name or "merged"
        output_name = f"{build_safe_file_name(base_name)}_combined.pdf"
        output_path = os.path.join(output_folder, output_name)
        output_path = ensure_unique_file_path(output_path)

        writer = PdfWriter()
        merged_count = 0
        try:
            for path in pdf_paths:
                if os.path.exists(path):
                    writer.append(path)
                    merged_count += 1
                else:
                    skipped_paths.append(path)
                    logging.error(f"PDF file not found for merging: {path}")

            if merged_count == 0:
                return PdfMergeResult(
                    success=False, error_message="Ни один из PDF-файлов не найден на диске."
                )

            with open(output_path, "wb") as f:
                writer.write(f)

            logging.info(f"PDFs merged successfully into: {output_path}")
            return PdfMergeResult(
                success=True,
                output_path=output_path,
                merged_count=merged_count,
                skipped_paths=skipped_paths,
            )

        except Exception as e:
            logging.error(f"Error merging PDFs: {e}")
            return PdfMergeResult(success=False, error_message=str(e))
        finally:
            writer.close()
