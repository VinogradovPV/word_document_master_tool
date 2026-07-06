from tkinter import ttk

from ..state import GuiState


class DocumentTableWidget(ttk.Frame):
    """Виджет таблицы документов (Treeview)."""

    def __init__(self, master, state: GuiState, on_change=None):
        super().__init__(master)
        self.state = state
        self.on_change = on_change
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
        self.tree.column("type", width=70, anchor="center", stretch=False)
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
        if not item_id:
            return
        
        idx = self.tree.index(item_id)
        if 0 <= idx < len(self.state.documents):
            doc = self.state.documents[idx]
            doc.is_selected = not doc.is_selected
            self.sync_with_state()
            self._notify_change()

    def selected_indices(self) -> list[int]:
        """Возвращает индексы строк, выделенных в Treeview."""
        return [self.tree.index(item_id) for item_id in self.tree.selection()]

    def sync_with_state(self):
        """Синхронизирует Treeview с GuiState."""
        # Сохраняем выделение
        selected_indices = [self.tree.index(i) for i in self.tree.selection()]
        
        self.tree.delete(*self.tree.get_children())
        for doc in self.state.documents:
            tag = "selected" if doc.is_selected else "unselected"
            self.tree.insert(
                "",
                "end",
                values=(
                    "Да" if doc.is_selected else "Нет",
                    doc.file_name,
                    doc.extension,
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
        for idx in selected_indices:
            if idx < len(children):
                self.tree.selection_add(children[idx])

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
