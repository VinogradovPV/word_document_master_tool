from tkinter import ttk

from ..state import GuiState


class DocumentTableWidget(ttk.Frame):
    """Виджет таблицы документов (Treeview)."""
    def __init__(self, master, state: GuiState):
        super().__init__(master)
        self.state = state
        self._create_widgets()

    def _create_widgets(self):
        self.tree = ttk.Treeview(
            self, 
            columns=("file", "status"), 
            show="headings", 
            selectmode="extended"
        )
        self.tree.heading("file", text="Файл")
        self.tree.heading("status", text="Статус")
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
                values=(doc.file_name, doc.status.value),
                tags=(tag,)
            )
        
        self.tree.tag_configure("selected", foreground="blue")
        self.tree.tag_configure("unselected", foreground="black")
        
        # Восстанавливаем выделение
        children = self.tree.get_children()
        for idx in selected_indices:
            if idx < len(children):
                self.tree.selection_add(children[idx])
