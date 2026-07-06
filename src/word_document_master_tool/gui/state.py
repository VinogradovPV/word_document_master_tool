from dataclasses import dataclass, field
from typing import List

from ..core.models import DocumentItem


@dataclass
class GuiState:
    documents: List[DocumentItem] = field(default_factory=list)

    def set_documents(self, documents: List[DocumentItem]) -> None:
        self.documents = documents

    def clear_documents(self) -> None:
        self.documents = []

    def selected_count(self) -> int:
        return sum(1 for doc in self.documents if doc.is_selected)
