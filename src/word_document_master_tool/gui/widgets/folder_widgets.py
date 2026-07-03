import tkinter as tk
from tkinter import filedialog, ttk


class FolderSelectWidget(ttk.Frame):
    """Виджет выбора папки с текстовым полем и кнопкой обзора."""
    def __init__(self, master, label_text: str, initial_value: str = ""):
        super().__init__(master)
        self.label = ttk.Label(self, text=label_text)
        self.label.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        self.entry = ttk.Entry(self)
        self.entry.insert(0, initial_value)
        self.entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        self.btn_browse = ttk.Button(self, text="Обзор...", command=self._browse)
        self.btn_browse.grid(row=0, column=2, padx=5, pady=2)
        
        self.columnconfigure(1, weight=1)

    def _browse(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, directory)

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: str):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, value)
