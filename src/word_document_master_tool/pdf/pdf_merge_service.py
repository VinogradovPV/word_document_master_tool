import logging
import os

from pypdf import PdfWriter

from ..core.models import DocumentItem, ToolSettings
from ..filesystem.paths import build_safe_file_name, ensure_unique_file_path


class PdfMergeService:
    def __init__(self, settings: ToolSettings):
        self.settings = settings

    def merge_pdfs(self, items: list[DocumentItem]) -> str:
        """
        Объединяет сгенерированные PDF файлы в один.
        Сохраняет порядок согласно списку.
        """
        pdf_paths = [item.pdf_path for item in items if item.is_selected and item.pdf_path]
        
        if not pdf_paths:
            logging.warning("No PDFs to merge.")
            return ""

        output_folder = self.settings.pdf.output_folder or self.settings.output_folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        base_name = self.settings.output_file_name or "merged"
        output_name = f"{build_safe_file_name(base_name)}_combined.pdf"
        output_path = os.path.join(output_folder, output_name)
        output_path = ensure_unique_file_path(output_path)

        writer = PdfWriter()

        try:
            for path in pdf_paths:
                if os.path.exists(path):
                    writer.append(path)
                else:
                    logging.error(f"PDF file not found for merging: {path}")

            with open(output_path, "wb") as f:
                writer.write(f)

            logging.info(f"PDFs merged successfully into: {output_path}")
            return output_path

        except Exception as e:
            logging.error(f"Error merging PDFs: {e}")
            return ""
        finally:
            writer.close()
