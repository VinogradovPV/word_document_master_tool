from tkinter import ttk

from ...core.models import SourceKind
from ..state import GuiState


class DocumentTableWidget(ttk.Frame):
    """Виджет таблицы документов (Treeview)."""

    def __init__(self, master, state: GuiState, on_change=None):
        super().__init__(master)
        self.state = state
        self.on_change = on_change
        self.filter_text = ""
        self.filter_mode = "all"
        self._visible_document_indices: list[int] = []
        self._create_widgets()

    def _create_widgets(self):
        self.tree = ttk.Treeview(
            self,
            columns=("enabled", "file", "type", "size", "modified", "status"),
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("enabled", text="Вкл")
        self.tree.heading("file", text="Файл")
        self.tree.heading("type", text="Тип")
        self.tree.heading("size", text="Размер")
        self.tree.heading("modified", text="Изменён")
        self.tree.heading("status", text="Статус")
        self.tree.column("enabled", width=50, anchor="center", stretch=False)
        self.tree.column("file", width=260, anchor="w")
        self.tree.column("type", width=100, anchor="center", stretch=False)
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("modified", width=150, anchor="center", stretch=False)
        self.tree.column("status", width=100, anchor="center")

        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._on_double_click)

    def _on_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or column_id != "#1":
            return

        idx = self.tree.index(item_id)
        if 0 <= idx < len(self._visible_document_indices):
            doc = self.state.documents[self._visible_document_indices[idx]]
            doc.is_selected = not doc.is_selected
            self.sync_with_state()
            self._notify_change()

    def selected_indices(self) -> list[int]:
        """Возвращает индексы выбранных документов в GuiState."""
        indices = []
        for item_id in self.tree.selection():
            visible_index = self.tree.index(item_id)
            if 0 <= visible_index < len(self._visible_document_indices):
                indices.append(self._visible_document_indices[visible_index])
        return indices

    def set_filter(self, text: str = "", mode: str = "all") -> None:
        self.filter_text = text.lower().strip()
        self.filter_mode = mode
        self.sync_with_state()

    def sync_with_state(self):
        """Синхронизирует Treeview с GuiState."""
        # Сохраняем выделение
        selected_state_indices = []
        for item_id in self.tree.selection():
            visible_index = self.tree.index(item_id)
            if 0 <= visible_index < len(self._visible_document_indices):
                selected_state_indices.append(self._visible_document_indices[visible_index])

        self.tree.delete(*self.tree.get_children())
        self._visible_document_indices = []
        for state_index, doc in enumerate(self.state.documents):
            if not self._matches_filter(doc):
                continue
            self._visible_document_indices.append(state_index)
            tag = "selected" if doc.is_selected else "unselected"
            self.tree.insert(
                "",
                "end",
                values=(
                    "Да" if doc.is_selected else "Нет",
                    doc.file_name,
                    self._format_source_type(doc),
                    self._format_size(doc.size_bytes),
                    doc.modified_at.strftime("%Y-%m-%d %H:%M"),
                    doc.status.value,
                ),
                tags=(tag,),
            )

        self.tree.tag_configure("selected", foreground="blue")
        self.tree.tag_configure("unselected", foreground="black")

        # Восстанавливаем выделение
        children = self.tree.get_children()
        for state_index in selected_state_indices:
            if state_index in self._visible_document_indices:
                visible_index = self._visible_document_indices.index(state_index)
                if visible_index < len(children):
                    self.tree.selection_add(children[visible_index])

    def _notify_change(self) -> None:
        if self.on_change is not None:
            self.on_change()

    @staticmethod
    def _format_size(size_bytes: int) -> str:
        if size_bytes < 1024:
            return f"{size_bytes} B"
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.1f} MB"

    def _matches_filter(self, doc) -> bool:
        if self.filter_text and self.filter_text not in doc.file_name.lower():
            return False
        if self.filter_mode == "selected" and not doc.is_selected:
            return False
        if self.filter_mode == "errors" and doc.status.value != "ERROR":
            return False
        if self.filter_mode == "word" and doc.source_kind != SourceKind.WORD:
            return False
        return not (self.filter_mode == "excel" and doc.source_kind != SourceKind.EXCEL)

    @staticmethod
    def _format_source_type(doc) -> str:
        if doc.source_kind == SourceKind.EXCEL:
            return f"Excel / {doc.extension}"
        if doc.source_kind == SourceKind.WORD:
            return f"Word / {doc.extension}"
        return doc.extension
