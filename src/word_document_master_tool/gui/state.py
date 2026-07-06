from dataclasses import dataclass, field

from ..core.models import DocumentItem


@dataclass
class GuiState:
    documents: list[DocumentItem] = field(default_factory=list)

    def set_documents(self, documents: list[DocumentItem]) -> None:
        self.documents = documents

    def clear_documents(self) -> None:
        self.documents = []

    def selected_count(self) -> int:
        return sum(1 for doc in self.documents if doc.is_selected)
