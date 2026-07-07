import os
import tkinter as tk
from tkinter import messagebox, ttk

from ..core.models import DocumentStatus, ToolSettings
from ..filesystem.discovery import DocumentDiscovery
from .controller import AppController
from .state import GuiState
from .widgets.document_table import DocumentTableWidget
from .widgets.folder_widgets import FolderSelectWidget

MERGE_MODE_LABELS = {
    "С новой строки": 1,
    "Пустой абзац": 2,
    "С новой страницы": 3,
    "Раздел со следующей страницы": 4,
}
PAGE_NUMBERING_MODES = ("Сквозная", "Заново в каждом документе")
HEADER_FOOTER_MODES = ("Не трогать", "Сохранять", "Очистить")
MARKER_VISIBILITY_LABELS = {
    "Невидимые": 1,
    "Ненавязчивые": 2,
    "Явные": 3,
}
MARKER_REMOVAL_LABELS = {
    "Только визуальные": 1,
    "Полное удаление": 2,
}
FOOTNOTE_MODES = (
    "Сквозная",
    "Заново в документе",
    "Заново в секции",
)


class MainWindow(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.state = GuiState()
        self.controller = AppController(self.state)
        self.last_operation = "Готово"
        self.last_log_path = ""
        self._create_widgets()

    def _create_widgets(self):
        self.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        self.tab_main = ttk.Frame(self.notebook)
        self.tab_documents = ttk.Frame(self.notebook)
        self.tab_processing = ttk.Frame(self.notebook)
        self.tab_page_numbering = ttk.Frame(self.notebook)
        self.tab_pdf = ttk.Frame(self.notebook)
        self.tab_markers = ttk.Frame(self.notebook)
        self.tab_footnotes = ttk.Frame(self.notebook)
        self.tab_journal = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_main, text="Главное")
        self.notebook.add(self.tab_documents, text="Документы")
        self.notebook.add(self.tab_processing, text="Обработка")
        self.notebook.add(self.tab_page_numbering, text="Нумерация страниц")
        self.notebook.add(self.tab_pdf, text="PDF")
        self.notebook.add(self.tab_markers, text="Маркеры")
        self.notebook.add(self.tab_footnotes, text="Сноски")
        self.notebook.add(self.tab_journal, text="Журнал")

        self._create_main_tab()
        self._create_documents_tab()
        self._create_processing_tab()
        self._create_page_numbering_tab()
        self._create_pdf_tab()
        self._create_markers_tab()
        self._create_footnotes_tab()
        self._create_journal_tab()
        self._create_action_panel()
        self._refresh_summary()

    def _create_main_tab(self):
        folder_frame = ttk.LabelFrame(self.tab_main, text="Папки и результат")
        folder_frame.pack(fill="x", padx=10, pady=10)

        self.wdg_source_folder = FolderSelectWidget(folder_frame, "Исходная папка:")
        self.wdg_source_folder.pack(fill="x", expand=True)

        self.wdg_output_folder = FolderSelectWidget(folder_frame, "Папка результата:")
        self.wdg_output_folder.pack(fill="x", expand=True)

        output_name_frame = ttk.Frame(folder_frame)
        output_name_frame.pack(fill="x", expand=True, padx=5, pady=2)
        ttk.Label(output_name_frame, text="Имя файла:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.ent_output_file_name = ttk.Entry(output_name_frame)
        self.ent_output_file_name.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        output_name_frame.columnconfigure(1, weight=1)

        ttk.Button(
            folder_frame, text="Обновить список", command=self._refresh_list
        ).pack(anchor="w", padx=5, pady=5)

        summary_frame = ttk.LabelFrame(self.tab_main, text="Сводка")
        summary_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.lbl_summary_documents = ttk.Label(summary_frame, text="")
        self.lbl_summary_documents.pack(anchor="w", padx=8, pady=3)
        self.lbl_summary_pdf = ttk.Label(summary_frame, text="")
        self.lbl_summary_pdf.pack(anchor="w", padx=8, pady=3)
        self.lbl_summary_operation = ttk.Label(summary_frame, text="")
        self.lbl_summary_operation.pack(anchor="w", padx=8, pady=3)
        self.lbl_summary_log = ttk.Label(summary_frame, text="")
        self.lbl_summary_log.pack(anchor="w", padx=8, pady=3)

    def _create_documents_tab(self):
        table_frame = ttk.LabelFrame(self.tab_documents, text="Документы")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.wdg_table = DocumentTableWidget(table_frame, self.state, self._update_counters)
        self.wdg_table.pack(fill="both", expand=True, padx=5, pady=5)

        buttons = ttk.Frame(table_frame)
        buttons.pack(fill="x", padx=5, pady=5)
        ttk.Button(buttons, text="Проверить", command=self._check_documents).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Вверх", command=lambda: self._move_item(-1)).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Вниз", command=lambda: self._move_item(1)).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Вкл", command=self._enable_selected).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Выкл", command=self._disable_selected).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Выбрать все", command=self._select_all).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Снять все", command=self._clear_all).pack(
            side="left", padx=2
        )

        self.lbl_counts = ttk.Label(buttons, text="Найдено: 0 | Выбрано: 0")
        self.lbl_counts.pack(side="right", padx=5)

    def _create_processing_tab(self):
        merge_frame = ttk.LabelFrame(self.tab_processing, text="Настройки объединения")
        merge_frame.pack(fill="x", padx=10, pady=10)
        merge_frame.columnconfigure(1, weight=1)
        merge_frame.columnconfigure(3, weight=1)

        ttk.Label(merge_frame, text="Формат:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_merge_format = ttk.Combobox(
            merge_frame, values=("docx", "docm", "rtf"), state="readonly", width=12
        )
        self.cmb_merge_format.set("docx")
        self.cmb_merge_format.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(merge_frame, text="Режим:").grid(
            row=0, column=2, sticky="w", padx=5, pady=2
        )
        self.cmb_merge_mode = ttk.Combobox(
            merge_frame, values=tuple(MERGE_MODE_LABELS), state="readonly"
        )
        self.cmb_merge_mode.set("С новой страницы")
        self.cmb_merge_mode.grid(row=0, column=3, sticky="ew", padx=5, pady=2)

        self.var_open_after_merge = tk.BooleanVar(value=True)
        self.var_create_report = tk.BooleanVar(value=True)
        self.var_create_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            merge_frame,
            text="Открыть после объединения",
            variable=self.var_open_after_merge,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            merge_frame, text="Создать отчёт", variable=self.var_create_report
        ).grid(row=1, column=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            merge_frame, text="Создать backup", variable=self.var_create_backup
        ).grid(row=1, column=3, sticky="w", padx=5, pady=2)

        source_frame = ttk.LabelFrame(self.tab_processing, text="Исправления и комментарии")
        source_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.var_accept_revisions = tk.BooleanVar(value=True)
        self.var_disable_track_changes = tk.BooleanVar(value=True)
        self.var_remove_comments = tk.BooleanVar(value=False)
        self.var_warn_protected_docs = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            source_frame,
            text="Принять все исправления",
            variable=self.var_accept_revisions,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            source_frame,
            text="Отключить Track Changes",
            variable=self.var_disable_track_changes,
        ).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            source_frame, text="Удалить комментарии", variable=self.var_remove_comments
        ).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            source_frame,
            text="Предупреждать о защите",
            variable=self.var_warn_protected_docs,
        ).grid(row=1, column=1, sticky="w", padx=5, pady=2)

    def _create_page_numbering_tab(self):
        page_frame = ttk.LabelFrame(self.tab_page_numbering, text="Нумерация страниц")
        page_frame.pack(fill="x", padx=10, pady=10)
        for column in (1, 3, 5):
            page_frame.columnconfigure(column, weight=1)

        self.var_page_numbering_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            page_frame,
            text="Включить нумерацию страниц",
            variable=self.var_page_numbering_enabled,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        self.spn_page_start = self._add_labeled_spinbox(page_frame, "Начать с:", 0, 2)
        self.spn_page_start.set("1")

        self.cmb_page_scope = self._add_labeled_combobox(
            page_frame,
            "Область:",
            1,
            0,
            ("Нумеровать обработанные копии документов", "Нумеровать итоговый документ"),
            "Нумеровать обработанные копии документов",
        )
        self.cmb_page_location = self._add_labeled_combobox(
            page_frame,
            "Место:",
            1,
            2,
            ("Верхний колонтитул", "Нижний колонтитул"),
            "Верхний колонтитул",
        )
        self.cmb_page_alignment = self._add_labeled_combobox(
            page_frame,
            "Выравнивание:",
            1,
            4,
            ("По левому краю", "По центру", "По правому краю"),
            "По центру",
        )
        self.cmb_page_format = self._add_labeled_combobox(
            page_frame,
            "Формат:",
            2,
            0,
            ("1, 2, 3", "i, ii, iii", "I, II, III"),
            "1, 2, 3",
        )

        ttk.Label(page_frame, text="Шрифт:").grid(
            row=2, column=2, sticky="w", padx=5, pady=2
        )
        self.ent_page_font = ttk.Entry(page_frame)
        self.ent_page_font.insert(0, "Times New Roman")
        self.ent_page_font.grid(row=2, column=3, sticky="ew", padx=5, pady=2)

        self.spn_page_font_size = self._add_labeled_spinbox(page_frame, "Размер:", 2, 4)
        self.spn_page_font_size.set("12")
        self.cmb_page_numbering_mode = self._add_labeled_combobox(
            page_frame, "Режим нумерации:", 3, 0, PAGE_NUMBERING_MODES, "Сквозная"
        )
        self.cmb_header_footer_mode = self._add_labeled_combobox(
            page_frame, "Колонтитулы:", 3, 2, HEADER_FOOTER_MODES, "Сохранять"
        )

        self.var_remove_existing_page = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            page_frame,
            text="Удалить старые PAGE",
            variable=self.var_remove_existing_page,
        ).grid(row=3, column=4, columnspan=2, sticky="w", padx=5, pady=2)

        self.var_adjust_page_margins = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            page_frame, text="Настроить поля", variable=self.var_adjust_page_margins
        ).grid(row=4, column=0, sticky="w", padx=5, pady=2)

        self.spn_margin_top = self._add_margin_spinbox(page_frame, "Верх:", 4, 1)
        self.spn_margin_bottom = self._add_margin_spinbox(page_frame, "Низ:", 4, 2)
        self.spn_margin_left = self._add_margin_spinbox(page_frame, "Лево:", 4, 3)
        self.spn_margin_right = self._add_margin_spinbox(page_frame, "Право:", 4, 4)

    def _create_pdf_tab(self):
        pdf_frame = ttk.LabelFrame(self.tab_pdf, text="PDF-экспорт")
        pdf_frame.pack(fill="x", padx=10, pady=10)
        pdf_frame.columnconfigure(1, weight=1)
        pdf_frame.columnconfigure(3, weight=1)

        self.var_pdf_merged = tk.BooleanVar(value=True)
        self.var_pdf_sources = tk.BooleanVar(value=True)
        self.var_pdf_processed = tk.BooleanVar(value=False)
        self.var_pdf_merge_generated = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            pdf_frame, text="PDF итогового документа", variable=self.var_pdf_merged
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            pdf_frame,
            text="PDF исходников без изменений",
            variable=self.var_pdf_sources,
            command=self._refresh_summary,
        ).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            pdf_frame, text="PDF обработанных копий", variable=self.var_pdf_processed
        ).grid(row=0, column=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            pdf_frame,
            text="Создать общий PDF из созданных PDF",
            variable=self.var_pdf_merge_generated,
        ).grid(row=0, column=3, sticky="w", padx=5, pady=2)

        self.wdg_pdf_folder = FolderSelectWidget(pdf_frame, "Папка PDF:")
        self.wdg_pdf_folder.grid(row=1, column=0, columnspan=4, sticky="ew")

        self.cmb_pdf_naming = self._add_labeled_combobox(
            pdf_frame,
            "Режим наименования:",
            2,
            0,
            ("Как исходный файл", "С префиксом порядка", "С суффиксом типа"),
            "Как исходный файл",
        )
        self.cmb_pdf_quality = self._add_labeled_combobox(
            pdf_frame, "Качество:", 2, 2, ("Печать", "Экран"), "Печать"
        )

        self.var_pdf_open = tk.BooleanVar(value=False)
        self.var_pdf_a = tk.BooleanVar(value=False)
        self.var_pdf_properties = tk.BooleanVar(value=True)
        ttk.Checkbutton(pdf_frame, text="Открыть PDF", variable=self.var_pdf_open).grid(
            row=3, column=0, sticky="w", padx=5, pady=2
        )
        ttk.Checkbutton(pdf_frame, text="PDF/A", variable=self.var_pdf_a).grid(
            row=3, column=1, sticky="w", padx=5, pady=2
        )
        ttk.Checkbutton(
            pdf_frame, text="Свойства", variable=self.var_pdf_properties
        ).grid(row=3, column=2, sticky="w", padx=5, pady=2)

    def _create_markers_tab(self):
        markers_frame = ttk.LabelFrame(self.tab_markers, text="Маркеры и разделение")
        markers_frame.pack(fill="x", padx=10, pady=10)
        markers_frame.columnconfigure(1, weight=1)
        markers_frame.columnconfigure(3, weight=1)

        self.var_use_markers = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            markers_frame,
            text="Добавлять маркеры частей",
            variable=self.var_use_markers,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.cmb_marker_visibility = self._add_labeled_combobox(
            markers_frame,
            "Вид маркеров:",
            0,
            1,
            tuple(MARKER_VISIBILITY_LABELS),
            "Невидимые",
        )
        self.cmb_marker_removal = self._add_labeled_combobox(
            markers_frame,
            "Режим удаления:",
            1,
            0,
            tuple(MARKER_REMOVAL_LABELS),
            "Только визуальные",
        )

        self.var_backup_before_marker_removal = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            markers_frame,
            text="Backup перед удалением",
            variable=self.var_backup_before_marker_removal,
        ).grid(row=1, column=2, sticky="w", padx=5, pady=2)
        ttk.Button(
            markers_frame, text="Удалить маркеры", command=self._show_backend_stub
        ).grid(row=1, column=3, sticky="w", padx=5, pady=2)
        ttk.Button(
            markers_frame, text="Разделить", command=self._on_split_documents
        ).grid(row=2, column=3, sticky="w", padx=5, pady=2)
        ttk.Label(
            markers_frame,
            text="Полное удаление опасно",
            foreground="firebrick",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=2)

    def _create_footnotes_tab(self):
        footnote_frame = ttk.LabelFrame(self.tab_footnotes, text="Сноски и концевые сноски")
        footnote_frame.pack(fill="x", padx=10, pady=10)
        footnote_frame.columnconfigure(1, weight=1)
        footnote_frame.columnconfigure(3, weight=1)

        self.var_footnotes_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            footnote_frame,
            text="Нумерация сносок",
            variable=self.var_footnotes_enabled,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.spn_footnote_start = self._add_labeled_spinbox(
            footnote_frame, "Начать с:", 0, 1
        )
        self.spn_footnote_start.set("1")
        self.cmb_footnote_scope = self._add_labeled_combobox(
            footnote_frame,
            "Область:",
            1,
            0,
            ("Нумеровать итоговый документ", "Нумеровать обработанные копии"),
            "Нумеровать итоговый документ",
        )
        self.cmb_footnote_mode = self._add_labeled_combobox(
            footnote_frame, "Режим:", 1, 2, FOOTNOTE_MODES, "Сквозная"
        )
        self.cmb_footnote_format = self._add_labeled_combobox(
            footnote_frame,
            "Формат:",
            2,
            0,
            ("1, 2, 3", "i, ii, iii", "I, II, III", "a, b, c"),
            "1, 2, 3",
        )

        self.var_preserve_footnote_text = tk.BooleanVar(value=True)
        self.var_process_endnotes = tk.BooleanVar(value=False)
        self.var_update_footnote_fields = tk.BooleanVar(value=True)
        self.var_keep_footnote_numbers = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            footnote_frame,
            text="Сохранять текст",
            variable=self.var_preserve_footnote_text,
        ).grid(row=2, column=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            footnote_frame, text="Концевые", variable=self.var_process_endnotes
        ).grid(row=2, column=3, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            footnote_frame, text="Поля", variable=self.var_update_footnote_fields
        ).grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(
            footnote_frame,
            text="Не заменять номера текстом",
            variable=self.var_keep_footnote_numbers,
        ).grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=2)

    def _create_journal_tab(self):
        journal_frame = ttk.LabelFrame(self.tab_journal, text="Журнал")
        journal_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.lbl_log_path = ttk.Label(journal_frame, text="Последний log-файл: -")
        self.lbl_log_path.pack(anchor="w", padx=5, pady=3)

        self.txt_log = tk.Text(journal_frame, height=12, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=5, pady=5)

        buttons = ttk.Frame(journal_frame)
        buttons.pack(fill="x", padx=5, pady=5)
        ttk.Button(buttons, text="Открыть log-файл", command=self._open_log_file).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Открыть папку logs", command=self._open_logs_folder).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Очистить журнал", command=self._clear_journal).pack(
            side="left", padx=2
        )
        ttk.Button(
            buttons, text="Копировать диагностику", command=self._copy_diagnostics
        ).pack(side="left", padx=2)

    def _create_action_panel(self):
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=8, pady=(4, 8))

        self.lbl_status = ttk.Label(action_frame, text="Статус: готово")
        self.lbl_status.pack(side="top", anchor="w", fill="x", padx=5, pady=2)

        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(side="top", fill="x", expand=True, padx=5, pady=2)

        buttons = ttk.Frame(action_frame)
        buttons.pack(side="top", fill="x", padx=5, pady=2)
        ttk.Button(buttons, text="Сбросить", command=self._reset).pack(side="left", padx=5)
        ttk.Button(buttons, text="Закрыть", command=self.master.destroy).pack(
            side="left", padx=5
        )
        self.btn_process = ttk.Button(
            buttons, text="Обработать файлы", command=self._on_process_files
        )
        self.btn_process.pack(side="right", padx=5)
        self.btn_merge = ttk.Button(buttons, text="Объединить", command=self._on_merge_documents)
        self.btn_merge.pack(side="right", padx=5)

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
        self._set_status(f"Найдено документов: {len(files)}")
        self._append_log(f"Список документов обновлен: {len(files)}")

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
            children = self.wdg_table.tree.get_children()
            self.wdg_table.tree.selection_set(children[new_idx])

    def _check_documents(self):
        for doc in self.state.documents:
            if doc.is_valid:
                doc.status = DocumentStatus.OK
                doc.error_message = ""
            else:
                doc.status = DocumentStatus.ERROR
                doc.error_message = "Документ не прошел проверку."
        self.wdg_table.sync_with_state()
        self._update_counters()
        self._set_status("Проверка документов завершена")
        self._append_log("Проверка документов завершена")

    def _enable_selected(self):
        self._set_selected_documents_enabled(True)

    def _disable_selected(self):
        self._set_selected_documents_enabled(False)

    def _set_selected_documents_enabled(self, enabled: bool):
        for idx in self.wdg_table.selected_indices():
            if 0 <= idx < len(self.state.documents):
                self.state.documents[idx].is_selected = enabled
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _select_all(self):
        for doc in self.state.documents:
            doc.is_selected = True
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _clear_all(self):
        for doc in self.state.documents:
            doc.is_selected = False
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _update_counters(self):
        total = len(self.state.documents)
        selected = self.state.selected_count()
        self.lbl_counts.config(text=f"Найдено: {total} | Выбрано: {selected}")
        self._refresh_summary()

    def _get_current_settings(self) -> ToolSettings:
        settings = ToolSettings(
            source_folder=self.wdg_source_folder.get(),
            output_folder=self.wdg_output_folder.get(),
            output_file_name=self.ent_output_file_name.get().strip(),
        )
        settings.merge.output_format = self.cmb_merge_format.get()
        settings.merge.mode = MERGE_MODE_LABELS.get(self.cmb_merge_mode.get(), 3)
        settings.merge.open_after_merge = self.var_open_after_merge.get()
        settings.merge.create_report = self.var_create_report.get()
        settings.merge.create_backup = self.var_create_backup.get()
        settings.source_processing.accept_revisions = self.var_accept_revisions.get()
        settings.source_processing.disable_track_changes = (
            self.var_disable_track_changes.get()
        )
        settings.source_processing.remove_comments = self.var_remove_comments.get()
        settings.source_processing.warn_protected_docs = self.var_warn_protected_docs.get()
        self._apply_page_settings(settings)
        self._apply_marker_settings(settings)
        self._apply_pdf_settings(settings)
        self._apply_footnote_settings(settings)
        return settings

    def _apply_page_settings(self, settings: ToolSettings):
        settings.page_numbering.enabled = self.var_page_numbering_enabled.get()
        settings.page_numbering.start_number = self._get_int(self.spn_page_start, 1)
        settings.page_numbering.scope = self.cmb_page_scope.get()
        settings.page_numbering.location = self.cmb_page_location.get()
        settings.page_numbering.alignment = self.cmb_page_alignment.get()
        settings.page_numbering.format = self.cmb_page_format.get()
        settings.page_numbering.font_name = self.ent_page_font.get().strip()
        settings.page_numbering.font_size = self._get_float(self.spn_page_font_size, 12)
        page_mode = self.cmb_page_numbering_mode.get()
        settings.page_numbering.continuous = page_mode == "Сквозная"
        settings.page_numbering.restart_each_document = (
            page_mode == "Заново в каждом документе"
        )
        header_mode = self.cmb_header_footer_mode.get()
        settings.page_numbering.preserve_headers_footers = header_mode == "Сохранять"
        settings.page_numbering.remove_headers_footers = header_mode == "Очистить"
        settings.page_numbering.remove_existing = self.var_remove_existing_page.get()
        settings.page_numbering.adjust_margins = self.var_adjust_page_margins.get()
        settings.page_numbering.top_margin_cm = self._get_float(self.spn_margin_top, 2)
        settings.page_numbering.bottom_margin_cm = self._get_float(
            self.spn_margin_bottom, 2
        )
        settings.page_numbering.left_margin_cm = self._get_float(self.spn_margin_left, 2)
        settings.page_numbering.right_margin_cm = self._get_float(
            self.spn_margin_right, 2
        )

    def _apply_marker_settings(self, settings: ToolSettings):
        settings.markers.use_markers = self.var_use_markers.get()
        settings.markers.visibility = MARKER_VISIBILITY_LABELS.get(
            self.cmb_marker_visibility.get(), 1
        )
        settings.markers.removal_mode = MARKER_REMOVAL_LABELS.get(
            self.cmb_marker_removal.get(), 1
        )
        settings.markers.backup_before_removal = (
            self.var_backup_before_marker_removal.get()
        )

    def _apply_pdf_settings(self, settings: ToolSettings):
        settings.pdf.export_sources = self.var_pdf_sources.get()
        settings.pdf.export_merged = self.var_pdf_merged.get()
        settings.pdf.export_processed_copies = self.var_pdf_processed.get()
        settings.pdf.merge_generated_pdfs = self.var_pdf_merge_generated.get()
        settings.pdf.output_folder = self.wdg_pdf_folder.get()
        settings.pdf.naming_mode = self.cmb_pdf_naming.get()
        settings.pdf.quality = self.cmb_pdf_quality.get()
        settings.pdf.open_after_export = self.var_pdf_open.get()
        settings.pdf.pdf_a = self.var_pdf_a.get()
        settings.pdf.include_properties = self.var_pdf_properties.get()
        settings.pdf.optimize_for_print = self.cmb_pdf_quality.get() == "Печать"
        settings.pdf.optimize_for_screen = self.cmb_pdf_quality.get() == "Экран"

    def _apply_footnote_settings(self, settings: ToolSettings):
        settings.footnotes.enabled = self.var_footnotes_enabled.get()
        settings.footnotes.start_number = self._get_int(self.spn_footnote_start, 1)
        settings.footnotes.scope = self.cmb_footnote_scope.get()
        footnote_mode = self.cmb_footnote_mode.get()
        settings.footnotes.mode = footnote_mode
        settings.footnotes.continuous = footnote_mode == "Сквозная"
        settings.footnotes.restart_each_document = footnote_mode == "Заново в документе"
        settings.footnotes.restart_each_section = footnote_mode == "Заново в секции"
        settings.footnotes.format = self.cmb_footnote_format.get()
        settings.footnotes.preserve_text = self.var_preserve_footnote_text.get()
        settings.footnotes.process_endnotes = self.var_process_endnotes.get()
        settings.footnotes.update_fields = self.var_update_footnote_fields.get()
        settings.footnotes.do_not_replace_numbers_with_text = (
            self.var_keep_footnote_numbers.get()
        )

    def _on_process_files(self):
        settings = self._get_current_settings()
        errors = settings.validate_errors()
        if errors:
            messagebox.showerror("Ошибка настроек", "\n".join(errors))
            return

        self.btn_process.config(state="disabled")
        self._set_status("Обработка файлов...")
        self._append_log("Запущена обработка файлов")
        self.controller.run_in_thread(
            lambda: self.controller.process_files(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self._finish_action(self.btn_process)),
        )

    def _on_merge_documents(self):
        settings = self._get_current_settings()
        self.btn_merge.config(state="disabled")
        self._set_status("Объединение документов...")
        self._append_log("Запущено объединение документов")
        self.controller.run_in_thread(
            lambda: self.controller.merge_documents(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self._finish_action(self.btn_merge)),
        )

    def _on_split_documents(self):
        settings = self._get_current_settings()
        self._set_status("Разделение по маркерам...")
        self._append_log("Запущено разделение по маркерам")
        self.controller.run_in_thread(
            lambda: self.controller.split_documents(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self._finish_action(None)),
        )

    def _show_backend_stub(self):
        messagebox.showinfo(
            "Функция не подключена",
            "Функция пока не подключена к backend.",
        )
        self._set_status("Функция пока не подключена к backend.")
        self._append_log("Функция пока не подключена к backend.")

    def _update_progress(self, current: int, total: int):
        self.master.after(0, lambda: self._perform_progress_update(current, total))

    def _perform_progress_update(self, current: int, total: int):
        if total > 0:
            self.progress["value"] = (current / total) * 100
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _finish_action(self, button):
        if button is not None:
            button.config(state="normal")
        self._set_status("Готово")
        self._append_log("Операция завершена")

    def _reset(self):
        self.state.clear_documents()
        self.wdg_table.sync_with_state()
        self.progress["value"] = 0
        self._update_counters()
        self._set_status("Сброшено")
        self._append_log("Состояние GUI сброшено")

    def _set_status(self, message: str):
        self.last_operation = message
        self.lbl_status.config(text=f"Статус: {message}")
        self._refresh_summary()

    def _append_log(self, message: str):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"{message}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self._refresh_summary(last_log=message)

    def _refresh_summary(self, last_log: str | None = None):
        if not hasattr(self, "lbl_summary_documents"):
            return
        total = len(self.state.documents)
        selected = self.state.selected_count()
        pdf_modes = []
        if hasattr(self, "var_pdf_sources") and self.var_pdf_sources.get():
            pdf_modes.append("исходники")
        if hasattr(self, "var_pdf_merged") and self.var_pdf_merged.get():
            pdf_modes.append("итоговый")
        if hasattr(self, "var_pdf_processed") and self.var_pdf_processed.get():
            pdf_modes.append("обработанные копии")
        self.lbl_summary_documents.config(text=f"Документы: найдено {total}, выбрано {selected}")
        self.lbl_summary_pdf.config(text=f"PDF-режимы: {', '.join(pdf_modes) or '-'}")
        self.lbl_summary_operation.config(text=f"Последняя операция: {self.last_operation}")
        self.lbl_summary_log.config(text=f"Последний лог: {last_log or '-'}")

    def _open_log_file(self):
        if self.last_log_path and os.path.exists(self.last_log_path):
            os.startfile(self.last_log_path)
            return
        messagebox.showinfo("Лог", "Log-файл пока не создан.")

    def _open_logs_folder(self):
        settings = self._get_current_settings()
        folder = settings.output_folder or settings.pdf.output_folder
        if folder and os.path.isdir(folder):
            os.startfile(folder)
            return
        messagebox.showinfo("Логи", "Папка logs пока не задана или не создана.")

    def _clear_journal(self):
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self._set_status("Журнал очищен")

    def _copy_diagnostics(self):
        diagnostics = [
            self.lbl_status.cget("text"),
            self.lbl_counts.cget("text"),
            f"Документов в state: {len(self.state.documents)}",
            f"Последний log-файл: {self.last_log_path or '-'}",
        ]
        self.clipboard_clear()
        self.clipboard_append("\n".join(diagnostics))
        self._set_status("Диагностика скопирована")

    def _add_labeled_combobox(
        self, master, label: str, row: int, column: int, values, default: str
    ):
        ttk.Label(master, text=label).grid(row=row, column=column, sticky="w", padx=5, pady=2)
        combobox = ttk.Combobox(master, values=values, state="readonly")
        combobox.set(default)
        combobox.grid(row=row, column=column + 1, sticky="ew", padx=5, pady=2)
        return combobox

    def _add_labeled_spinbox(self, master, label: str, row: int, column: int):
        ttk.Label(master, text=label).grid(row=row, column=column, sticky="w", padx=5, pady=2)
        spinbox = ttk.Spinbox(master, from_=1, to=9999, width=8)
        spinbox.grid(row=row, column=column + 1, sticky="w", padx=5, pady=2)
        return spinbox

    def _add_margin_spinbox(self, master, label: str, row: int, column: int):
        frame = ttk.Frame(master)
        frame.grid(row=row, column=column, sticky="ew", padx=5, pady=2)
        ttk.Label(frame, text=label).pack(side="left")
        spinbox = ttk.Spinbox(frame, from_=0, to=20, increment=0.1, width=6)
        spinbox.set("2.0")
        spinbox.pack(side="left", padx=2)
        return spinbox

    @staticmethod
    def _get_int(widget, default: int) -> int:
        try:
            return int(widget.get())
        except ValueError:
            return default

    @staticmethod
    def _get_float(widget, default: float) -> float:
        try:
            return float(widget.get().replace(",", "."))
        except ValueError:
            return default
