import tkinter as tk
from tkinter import messagebox, ttk

from ..core.models import ToolSettings
from ..filesystem.discovery import DocumentDiscovery
from .controller import AppController
from .state import GuiState
from .widgets.document_table import DocumentTableWidget
from .widgets.folder_widgets import FolderSelectWidget


class MainWindow(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.state = GuiState()
        self.controller = AppController(self.state)
        self._create_widgets()

    def _create_widgets(self):
        # Главный контейнер
        self.pack(fill="both", expand=True)
        
        # 1. Папки
        folder_frame = ttk.LabelFrame(self, text="Папки")
        folder_frame.pack(fill="x", padx=10, pady=5)
        
        self.wdg_source_folder = FolderSelectWidget(folder_frame, "Исходная папка:")
        self.wdg_source_folder.pack(fill="x", expand=True)
        
        self.wdg_output_folder = FolderSelectWidget(folder_frame, "Папка результата:")
        self.wdg_output_folder.pack(fill="x", expand=True)
        
        btn_refresh = ttk.Button(folder_frame, text="Обновить список", command=self._refresh_list)
        btn_refresh.pack(pady=5)

        # 2. Список файлов
        list_frame = ttk.LabelFrame(self, text="Документы")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.wdg_table = DocumentTableWidget(list_frame, self.state)
        self.wdg_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Кнопки управления списком
        btn_ctrl_frame = ttk.Frame(list_frame)
        btn_ctrl_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Button(btn_ctrl_frame, text="Вверх", command=lambda: self._move_item(-1)).pack(side="left", padx=2)
        ttk.Button(btn_ctrl_frame, text="Вниз", command=lambda: self._move_item(1)).pack(side="left", padx=2)
        self.lbl_counts = ttk.Label(btn_ctrl_frame, text="Найдено: 0 | Выбрано: 0")
        self.lbl_counts.pack(side="right", padx=5)

        # 3. Настройки PDF
        pdf_frame = ttk.LabelFrame(self, text="Настройки PDF")
        pdf_frame.pack(fill="x", padx=10, pady=5)
        
        self.var_pdf_sources = tk.BooleanVar(value=True)
        ttk.Checkbutton(pdf_frame, text="Экспорт исходников в PDF", variable=self.var_pdf_sources).pack(anchor="w", padx=5)
        
        self.var_pdf_merged = tk.BooleanVar(value=True)
        ttk.Checkbutton(pdf_frame, text="Экспорт результата в PDF", variable=self.var_pdf_merged).pack(anchor="w", padx=5)
        
        self.wdg_pdf_folder = FolderSelectWidget(pdf_frame, "Папка для PDF:")
        self.wdg_pdf_folder.pack(fill="x", expand=True)

        # 4. Прогресс и действия
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=5)
        
        self.btn_process = ttk.Button(action_frame, text="Обработать файлы", command=self._on_process_files)
        self.btn_process.pack(side="right", padx=5)
        
        self.btn_merge = ttk.Button(action_frame, text="Слияние документов", command=self._on_merge_documents)
        self.btn_merge.pack(side="right", padx=5)
        
        self.btn_split = ttk.Button(action_frame, text="Разделить по маркерам", command=self._on_split_documents)
        self.btn_split.pack(side="right", padx=5)

    def _refresh_list(self):
        source_dir = self.wdg_source_folder.get()
        if not source_dir:
            messagebox.showwarning("Внимание", "Выберите исходную папку")
            return
            
        discovery = DocumentDiscovery()
        files = discovery.find_documents(source_dir)
        self.state.set_documents(files)
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _move_item(self, direction: int):
        selected = self.wdg_table.tree.selection()
        if not selected:
            return
        
        idx = self.wdg_table.tree.index(selected[0])
        new_idx = idx + direction
        
        if 0 <= new_idx < len(self.state.documents):
            docs = self.state.documents
            docs[idx], docs[new_idx] = docs[new_idx], docs[idx]
            self.wdg_table.sync_with_state()
            # Восстанавливаем выделение на новом месте
            children = self.wdg_table.tree.get_children()
            self.wdg_table.tree.selection_set(children[new_idx])

    def _update_counters(self):
        total = len(self.state.documents)
        selected = sum(1 for d in self.state.documents if d.is_selected)
        self.lbl_counts.config(text=f"Найдено: {total} | Выбрано: {selected}")

    def _get_current_settings(self) -> ToolSettings:
        settings = ToolSettings(
            output_folder=self.wdg_output_folder.get(),
        )
        settings.pdf.export_sources = self.var_pdf_sources.get()
        settings.pdf.export_merged_document = self.var_pdf_merged.get()
        settings.pdf.output_folder = self.wdg_pdf_folder.get()
        return settings

    def _on_process_files(self):
        settings = self._get_current_settings()
        errors = settings.validate_errors()
        if errors:
            messagebox.showerror("Ошибка настроек", "\n".join(errors))
            return

        self.btn_process.config(state="disabled")
        self.controller.run_in_thread(
            lambda: self.controller.process_files(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self.btn_process.config(state="normal"))
        )

    def _on_merge_documents(self):
        settings = self._get_current_settings()
        self.btn_merge.config(state="disabled")
        self.controller.run_in_thread(
            lambda: self.controller.merge_documents(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self.btn_merge.config(state="normal"))
        )

    def _on_split_documents(self):
        settings = self._get_current_settings()
        self.btn_split.config(state="disabled")
        self.controller.run_in_thread(
            lambda: self.controller.split_documents(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self.btn_split.config(state="normal"))
        )

    def _update_progress(self, current: int, total: int):
        self.master.after(0, lambda: self._perform_progress_update(current, total))

    def _perform_progress_update(self, current: int, total: int):
        if total > 0:
            val = (current / total) * 100
            self.progress["value"] = val
        self.wdg_table.sync_with_state()
        self._update_counters()
