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
HEADER_FOOTER_MODES = ("Сохранять колонтитулы", "Очистить колонтитулы")
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

        output_name_frame = ttk.Frame(folder_frame)
        output_name_frame.pack(fill="x", expand=True, padx=5, pady=2)
        ttk.Label(output_name_frame, text="Имя файла:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.ent_output_file_name = ttk.Entry(output_name_frame)
        self.ent_output_file_name.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        output_name_frame.columnconfigure(1, weight=1)
        
        btn_refresh = ttk.Button(folder_frame, text="Обновить список", command=self._refresh_list)
        btn_refresh.pack(pady=5)

        # 2. Список файлов
        list_frame = ttk.LabelFrame(self, text="Документы")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.wdg_table = DocumentTableWidget(list_frame, self.state, self._update_counters)
        self.wdg_table.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Кнопки управления списком
        btn_ctrl_frame = ttk.Frame(list_frame)
        btn_ctrl_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Button(btn_ctrl_frame, text="Проверить", command=self._check_documents).pack(
            side="left", padx=2
        )
        ttk.Button(
            btn_ctrl_frame, text="Вверх", command=lambda: self._move_item(-1)
        ).pack(side="left", padx=2)
        ttk.Button(
            btn_ctrl_frame, text="Вниз", command=lambda: self._move_item(1)
        ).pack(side="left", padx=2)
        ttk.Button(btn_ctrl_frame, text="Вкл", command=self._enable_selected).pack(
            side="left", padx=2
        )
        ttk.Button(btn_ctrl_frame, text="Выкл", command=self._disable_selected).pack(
            side="left", padx=2
        )
        ttk.Button(btn_ctrl_frame, text="Выбрать все", command=self._select_all).pack(
            side="left", padx=2
        )
        ttk.Button(btn_ctrl_frame, text="Снять все", command=self._clear_all).pack(
            side="left", padx=2
        )
        self.lbl_counts = ttk.Label(btn_ctrl_frame, text="Найдено: 0 | Выбрано: 0")
        self.lbl_counts.pack(side="right", padx=5)

        # 3. Настройки объединения
        merge_frame = ttk.LabelFrame(self, text="Настройки объединения")
        merge_frame.pack(fill="x", padx=10, pady=5)
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
        ttk.Checkbutton(
            merge_frame,
            text="Открыть после объединения",
            variable=self.var_open_after_merge,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        self.var_create_report = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            merge_frame, text="Создать отчёт", variable=self.var_create_report
        ).grid(row=1, column=2, sticky="w", padx=5, pady=2)

        self.var_create_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            merge_frame, text="Создать backup", variable=self.var_create_backup
        ).grid(row=1, column=3, sticky="w", padx=5, pady=2)

        # 4. Исправления и комментарии
        source_frame = ttk.LabelFrame(self, text="Исправления и комментарии")
        source_frame.pack(fill="x", padx=10, pady=5)

        self.var_accept_revisions = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            source_frame,
            text="Принять все исправления",
            variable=self.var_accept_revisions,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.var_disable_track_changes = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            source_frame,
            text="Отключить Track Changes",
            variable=self.var_disable_track_changes,
        ).grid(row=0, column=1, sticky="w", padx=5, pady=2)

        self.var_remove_comments = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            source_frame, text="Удалить комментарии", variable=self.var_remove_comments
        ).grid(row=1, column=0, sticky="w", padx=5, pady=2)

        self.var_warn_protected_docs = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            source_frame,
            text="Предупреждать о защите",
            variable=self.var_warn_protected_docs,
        ).grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # 5. Нумерация страниц
        page_frame = ttk.LabelFrame(self, text="Нумерация страниц")
        page_frame.pack(fill="x", padx=10, pady=5)
        page_frame.columnconfigure(1, weight=1)
        page_frame.columnconfigure(3, weight=1)
        page_frame.columnconfigure(5, weight=1)

        self.var_page_numbering_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            page_frame,
            text="Включить нумерацию страниц",
            variable=self.var_page_numbering_enabled,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        ttk.Label(page_frame, text="Начать с:").grid(
            row=0, column=2, sticky="w", padx=5, pady=2
        )
        self.spn_page_start = ttk.Spinbox(page_frame, from_=1, to=9999, width=8)
        self.spn_page_start.set("1")
        self.spn_page_start.grid(row=0, column=3, sticky="w", padx=5, pady=2)

        ttk.Label(page_frame, text="Область:").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_page_scope = ttk.Combobox(
            page_frame,
            values=(
                "Нумеровать обработанные копии документов",
                "Нумеровать итоговый документ",
            ),
            state="readonly",
        )
        self.cmb_page_scope.set("Нумеровать обработанные копии документов")
        self.cmb_page_scope.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Место:").grid(
            row=1, column=2, sticky="w", padx=5, pady=2
        )
        self.cmb_page_location = ttk.Combobox(
            page_frame,
            values=("Верхний колонтитул", "Нижний колонтитул"),
            state="readonly",
        )
        self.cmb_page_location.set("Верхний колонтитул")
        self.cmb_page_location.grid(row=1, column=3, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Выравнивание:").grid(
            row=1, column=4, sticky="w", padx=5, pady=2
        )
        self.cmb_page_alignment = ttk.Combobox(
            page_frame,
            values=("По левому краю", "По центру", "По правому краю"),
            state="readonly",
        )
        self.cmb_page_alignment.set("По центру")
        self.cmb_page_alignment.grid(row=1, column=5, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Формат:").grid(
            row=2, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_page_format = ttk.Combobox(
            page_frame, values=("1, 2, 3", "i, ii, iii", "I, II, III"), state="readonly"
        )
        self.cmb_page_format.set("1, 2, 3")
        self.cmb_page_format.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Шрифт:").grid(
            row=2, column=2, sticky="w", padx=5, pady=2
        )
        self.ent_page_font = ttk.Entry(page_frame)
        self.ent_page_font.insert(0, "Times New Roman")
        self.ent_page_font.grid(row=2, column=3, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Размер:").grid(
            row=2, column=4, sticky="w", padx=5, pady=2
        )
        self.spn_page_font_size = ttk.Spinbox(page_frame, from_=1, to=72, width=8)
        self.spn_page_font_size.set("12")
        self.spn_page_font_size.grid(row=2, column=5, sticky="w", padx=5, pady=2)

        ttk.Label(page_frame, text="Режим:").grid(
            row=3, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_page_numbering_mode = ttk.Combobox(
            page_frame, values=PAGE_NUMBERING_MODES, state="readonly"
        )
        self.cmb_page_numbering_mode.set("Сквозная")
        self.cmb_page_numbering_mode.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(page_frame, text="Колонтитулы:").grid(
            row=3, column=2, sticky="w", padx=5, pady=2
        )
        self.cmb_header_footer_mode = ttk.Combobox(
            page_frame, values=HEADER_FOOTER_MODES, state="readonly"
        )
        self.cmb_header_footer_mode.set("Сохранять колонтитулы")
        self.cmb_header_footer_mode.grid(row=3, column=3, sticky="ew", padx=5, pady=2)

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

        # 6. Маркеры и разделение
        markers_frame = ttk.LabelFrame(self, text="Маркеры и разделение")
        markers_frame.pack(fill="x", padx=10, pady=5)
        markers_frame.columnconfigure(1, weight=1)
        markers_frame.columnconfigure(3, weight=1)

        self.var_use_markers = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            markers_frame,
            text="Добавлять маркеры частей",
            variable=self.var_use_markers,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)

        ttk.Label(markers_frame, text="Вид маркеров:").grid(
            row=0, column=1, sticky="w", padx=5, pady=2
        )
        self.cmb_marker_visibility = ttk.Combobox(
            markers_frame, values=tuple(MARKER_VISIBILITY_LABELS), state="readonly"
        )
        self.cmb_marker_visibility.set("Невидимые")
        self.cmb_marker_visibility.grid(row=0, column=2, sticky="ew", padx=5, pady=2)

        ttk.Label(markers_frame, text="Режим удаления:").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_marker_removal = ttk.Combobox(
            markers_frame, values=tuple(MARKER_REMOVAL_LABELS), state="readonly"
        )
        self.cmb_marker_removal.set("Только визуальные")
        self.cmb_marker_removal.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        self.var_backup_before_marker_removal = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            markers_frame,
            text="Backup перед удалением",
            variable=self.var_backup_before_marker_removal,
        ).grid(row=1, column=2, sticky="w", padx=5, pady=2)

        ttk.Button(
            markers_frame, text="Удалить маркеры", command=self._show_backend_stub
        ).grid(row=1, column=3, sticky="w", padx=5, pady=2)

        ttk.Label(
            markers_frame,
            text="Полное удаление опасно",
            foreground="firebrick",
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=2)

        # 7. Настройки PDF
        pdf_frame = ttk.LabelFrame(self, text="Настройки PDF")
        pdf_frame.pack(fill="x", padx=10, pady=5)
        pdf_frame.columnconfigure(1, weight=1)
        pdf_frame.columnconfigure(3, weight=1)
        
        self.var_pdf_sources = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            pdf_frame, text="Экспорт исходников в PDF", variable=self.var_pdf_sources
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        self.var_pdf_merged = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            pdf_frame, text="Экспорт результата в PDF", variable=self.var_pdf_merged
        ).grid(row=0, column=1, sticky="w", padx=5, pady=2)

        self.var_pdf_processed = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            pdf_frame,
            text="PDF обработанных копий",
            variable=self.var_pdf_processed,
        ).grid(row=0, column=2, sticky="w", padx=5, pady=2)

        self.var_pdf_merge_generated = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            pdf_frame,
            text="Создать общий PDF из созданных PDF",
            variable=self.var_pdf_merge_generated,
        ).grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        self.wdg_pdf_folder = FolderSelectWidget(pdf_frame, "Папка для PDF:")
        self.wdg_pdf_folder.grid(row=1, column=0, columnspan=4, sticky="ew")

        ttk.Label(pdf_frame, text="Режим наименования:").grid(
            row=2, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_pdf_naming = ttk.Combobox(
            pdf_frame,
            values=("Как исходный файл", "С префиксом порядка", "С суффиксом типа"),
            state="readonly",
        )
        self.cmb_pdf_naming.set("Как исходный файл")
        self.cmb_pdf_naming.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(pdf_frame, text="Качество:").grid(
            row=2, column=2, sticky="w", padx=5, pady=2
        )
        self.cmb_pdf_quality = ttk.Combobox(
            pdf_frame, values=("Печать", "Экран"), state="readonly", width=12
        )
        self.cmb_pdf_quality.set("Печать")
        self.cmb_pdf_quality.grid(row=2, column=3, sticky="w", padx=5, pady=2)

        self.var_pdf_open = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            pdf_frame, text="Открыть PDF", variable=self.var_pdf_open
        ).grid(row=3, column=0, sticky="w", padx=5, pady=2)

        self.var_pdf_a = tk.BooleanVar(value=False)
        ttk.Checkbutton(pdf_frame, text="PDF/A", variable=self.var_pdf_a).grid(
            row=3, column=1, sticky="w", padx=5, pady=2
        )

        self.var_pdf_properties = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            pdf_frame, text="Свойства", variable=self.var_pdf_properties
        ).grid(row=3, column=2, sticky="w", padx=5, pady=2)

        # 8. Сноски и концевые сноски
        footnote_frame = ttk.LabelFrame(self, text="Сноски и концевые сноски")
        footnote_frame.pack(fill="x", padx=10, pady=5)
        footnote_frame.columnconfigure(1, weight=1)
        footnote_frame.columnconfigure(3, weight=1)

        self.var_footnotes_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            footnote_frame,
            text="Нумерация сносок",
            variable=self.var_footnotes_enabled,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)

        ttk.Label(footnote_frame, text="Начать с:").grid(
            row=0, column=1, sticky="w", padx=5, pady=2
        )
        self.spn_footnote_start = ttk.Spinbox(footnote_frame, from_=1, to=9999, width=8)
        self.spn_footnote_start.set("1")
        self.spn_footnote_start.grid(row=0, column=2, sticky="w", padx=5, pady=2)

        ttk.Label(footnote_frame, text="Область:").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_footnote_scope = ttk.Combobox(
            footnote_frame,
            values=("Нумеровать итоговый документ", "Нумеровать обработанные копии"),
            state="readonly",
        )
        self.cmb_footnote_scope.set("Нумеровать итоговый документ")
        self.cmb_footnote_scope.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(footnote_frame, text="Режим:").grid(
            row=1, column=2, sticky="w", padx=5, pady=2
        )
        self.cmb_footnote_mode = ttk.Combobox(
            footnote_frame, values=FOOTNOTE_MODES, state="readonly"
        )
        self.cmb_footnote_mode.set("Сквозная")
        self.cmb_footnote_mode.grid(row=1, column=3, sticky="ew", padx=5, pady=2)

        ttk.Label(footnote_frame, text="Формат:").grid(
            row=2, column=0, sticky="w", padx=5, pady=2
        )
        self.cmb_footnote_format = ttk.Combobox(
            footnote_frame,
            values=("1, 2, 3", "i, ii, iii", "I, II, III", "a, b, c"),
            state="readonly",
        )
        self.cmb_footnote_format.set("1, 2, 3")
        self.cmb_footnote_format.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        self.var_preserve_footnote_text = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            footnote_frame,
            text="Сохранять текст",
            variable=self.var_preserve_footnote_text,
        ).grid(row=2, column=2, sticky="w", padx=5, pady=2)

        self.var_process_endnotes = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            footnote_frame, text="Концевые", variable=self.var_process_endnotes
        ).grid(row=2, column=3, sticky="w", padx=5, pady=2)

        self.var_update_footnote_fields = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            footnote_frame, text="Поля", variable=self.var_update_footnote_fields
        ).grid(row=3, column=0, sticky="w", padx=5, pady=2)

        self.var_keep_footnote_numbers = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            footnote_frame,
            text="Не заменять номера текстом",
            variable=self.var_keep_footnote_numbers,
        ).grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=2)

        # 9. Прогресс и действия
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=10, pady=10)

        self.lbl_status = ttk.Label(action_frame, text="Статус: готово")
        self.lbl_status.pack(side="top", anchor="w", fill="x", padx=5, pady=2)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(side="top", fill="x", expand=True, padx=5, pady=2)

        log_frame = ttk.Frame(action_frame)
        log_frame.pack(side="top", fill="both", expand=True, padx=5, pady=2)
        self.txt_log = tk.Text(log_frame, height=4, wrap="word", state="disabled")
        self.txt_log.pack(side="left", fill="both", expand=True)
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        log_scrollbar.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=log_scrollbar.set)

        buttons_frame = ttk.Frame(action_frame)
        buttons_frame.pack(side="top", fill="x", padx=5, pady=2)
        
        self.btn_process = ttk.Button(
            buttons_frame, text="Обработать файлы", command=self._on_process_files
        )
        self.btn_process.pack(side="right", padx=5)
        
        self.btn_merge = ttk.Button(
            buttons_frame, text="Объединить", command=self._on_merge_documents
        )
        self.btn_merge.pack(side="right", padx=5)
        
        self.btn_split = ttk.Button(
            buttons_frame, text="Разделить по маркерам", command=self._on_split_documents
        )
        self.btn_split.pack(side="right", padx=5)

        ttk.Button(buttons_frame, text="Сбросить", command=self._reset).pack(
            side="left", padx=5
        )
        ttk.Button(buttons_frame, text="Закрыть", command=self.master.destroy).pack(
            side="left", padx=5
        )

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
            # Восстанавливаем выделение на новом месте
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
        selected = sum(1 for d in self.state.documents if d.is_selected)
        self.lbl_counts.config(text=f"Найдено: {total} | Выбрано: {selected}")

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
        settings.source_processing.warn_protected_docs = (
            self.var_warn_protected_docs.get()
        )
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
        settings.page_numbering.preserve_headers_footers = (
            header_mode == "Сохранять колонтитулы"
        )
        settings.page_numbering.remove_headers_footers = (
            header_mode == "Очистить колонтитулы"
        )
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
        return settings

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
        self.btn_split.config(state="disabled")
        self._set_status("Разделение по маркерам...")
        self._append_log("Запущено разделение по маркерам")
        self.controller.run_in_thread(
            lambda: self.controller.split_documents(settings, self._update_progress),
            lambda: self.master.after(0, lambda: self._finish_action(self.btn_split)),
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
            val = (current / total) * 100
            self.progress["value"] = val
        self.wdg_table.sync_with_state()
        self._update_counters()

    def _finish_action(self, button):
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
        self.lbl_status.config(text=f"Статус: {message}")

    def _append_log(self, message: str):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"{message}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

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
