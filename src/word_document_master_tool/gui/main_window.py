import tkinter as tk
from tkinter import ttk


class MainWindow(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.document_items = []
        self._create_widgets()

    def _create_widgets(self):
        # Главный контейнер со скроллбаром, если нужно
        main_canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        self.scrollable_frame = ttk.Frame(main_canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # 1. Папки и результат
        self._create_folder_frame(self.scrollable_frame)

        # 2. Таблица документов
        self._create_document_table_frame(self.scrollable_frame)

        # 3. Настройки слияния
        self._create_merge_settings_frame(self.scrollable_frame)

        # 4. Рецензии и комментарии
        self._create_source_processing_frame(self.scrollable_frame)

        # 5. Нумерация страниц
        self._create_page_numbering_frame(self.scrollable_frame)

        # 6. Экспорт в PDF
        self._create_pdf_export_frame(self.scrollable_frame)

        # 7. Маркеры и разделение
        self._create_markers_frame(self.scrollable_frame)

        # 8. Сноски
        self._create_footnotes_frame(self.scrollable_frame)

        # 9. Статус и действия
        self._create_actions_frame(self.scrollable_frame)

        # Размещение главного контейнера
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _create_folder_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Папки и результат")
        frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame, text="Исходная папка:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.ent_source_folder = ttk.Entry(frame)
        self.ent_source_folder.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(frame, text="Обзор...").grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(frame, text="Папка результата:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.ent_output_folder = ttk.Entry(frame)
        self.ent_output_folder.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(frame, text="Обзор...").grid(row=1, column=2, padx=5, pady=2)

        ttk.Label(frame, text="Имя файла:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.ent_output_filename = ttk.Entry(frame)
        self.ent_output_filename.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        frame.columnconfigure(1, weight=1)

    def _create_document_table_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Список документов")
        frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("include", "file", "type", "size", "modified", "status")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=10)
        
        self.tree.heading("include", text="Вкл")
        self.tree.heading("file", text="Файл")
        self.tree.heading("type", text="Тип")
        self.tree.heading("size", text="Размер")
        self.tree.heading("modified", text="Изменён")
        self.tree.heading("status", text="Статус")

        self.tree.column("include", width=40, anchor="center")
        self.tree.column("file", width=300)
        self.tree.column("type", width=50)
        self.tree.column("size", width=80)
        self.tree.column("modified", width=120)
        self.tree.column("status", width=100)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side="right", fill="y", padx=5)
        ttk.Button(btn_frame, text="Вверх", command=self._move_up).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="Вниз", command=self._move_down).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="Включить все", command=self._select_all).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="Выключить все", command=self._clear_all).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="Обновить", command=self._refresh_list).pack(fill="x", pady=5)

    def _create_merge_settings_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Настройки слияния")
        frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame, text="Режим слияния:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.cmb_merge_mode = ttk.Combobox(frame, values=[
            "Со следующей строки",
            "Через один пустой абзац",
            "С новой страницы",
            "С разрывом раздела"
        ])
        self.cmb_merge_mode.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.cmb_merge_mode.current(2)

        self.chk_open_after = ttk.Checkbutton(frame, text="Открыть после слияния")
        self.chk_open_after.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)

    def _create_source_processing_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Рецензии и комментарии")
        frame.pack(fill="x", padx=10, pady=5)

        self.chk_accept_revisions = ttk.Checkbutton(frame, text="Принять исправления")
        self.chk_accept_revisions.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.chk_disable_track = ttk.Checkbutton(frame, text="Отключить Track Changes")
        self.chk_disable_track.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        self.chk_remove_comments = ttk.Checkbutton(frame, text="Удалить комментарии")
        self.chk_remove_comments.grid(row=1, column=0, sticky="w", padx=5, pady=2)

    def _create_page_numbering_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Нумерация страниц")
        frame.pack(fill="x", padx=10, pady=5)

        self.chk_page_num_enabled = ttk.Checkbutton(frame, text="Включить нумерацию страниц")
        self.chk_page_num_enabled.grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        ttk.Label(frame, text="Начать с:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.ent_page_start = ttk.Entry(frame, width=10)
        self.ent_page_start.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(frame, text="Расположение:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.cmb_page_loc = ttk.Combobox(frame, values=["Верхний колонтитул", "Нижний колонтитул"])
        self.cmb_page_loc.grid(row=1, column=3, sticky="w", padx=5, pady=2)

    def _create_pdf_export_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Экспорт в PDF")
        frame.pack(fill="x", padx=10, pady=5)

        self.chk_pdf_merged = ttk.Checkbutton(frame, text="PDF итогового документа")
        self.chk_pdf_merged.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.chk_pdf_sources = ttk.Checkbutton(frame, text="PDF исходников без изменений")
        self.chk_pdf_sources.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(frame, text="Папка PDF:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.ent_pdf_folder = ttk.Entry(frame)
        self.ent_pdf_folder.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(frame, text="Обзор...").grid(row=1, column=2, padx=5, pady=2)

        frame.columnconfigure(1, weight=1)

    def _create_markers_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Маркеры и разделение")
        frame.pack(fill="x", padx=10, pady=5)

        self.chk_use_markers = ttk.Checkbutton(frame, text="Использовать маркеры")
        self.chk_use_markers.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        ttk.Button(frame, text="Разделить по маркерам").grid(row=0, column=1, padx=5, pady=2)

    def _create_footnotes_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Сноски")
        frame.pack(fill="x", padx=10, pady=5)

        self.chk_footnote_enabled = ttk.Checkbutton(frame, text="Включить нумерацию сносок")
        self.chk_footnote_enabled.grid(row=0, column=0, sticky="w", padx=5, pady=2)

    def _create_actions_frame(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", padx=10, pady=10)

        self.lbl_status = ttk.Label(frame, text="Готово")
        self.lbl_status.pack(side="left", padx=5)

        self.lbl_counts = ttk.Label(frame, text="Найдено: 0 | Выбрано: 0")
        self.lbl_counts.pack(side="left", padx=15)

        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=5)

        ttk.Button(frame, text="Обработать файлы").pack(side="right", padx=5)
        ttk.Button(frame, text="Слияние документов").pack(side="right", padx=5)

    def _refresh_list(self) -> None:
        from tkinter import messagebox

        from ..filesystem.discovery import find_word_documents

        source_folder = self.ent_source_folder.get().strip()

        if not source_folder:
            messagebox.showwarning("Исходная папка", "Укажите исходную папку.")
            return

        try:
            items = find_word_documents(source_folder)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось прочитать папку:\n{exc}")
            return

        self.document_items = items

        for row_id in self.tree.get_children():
            self.tree.delete(row_id)

        for item in items:
            self.tree.insert(
                "",
                "end",
                values=(
                    "Да" if item.is_selected else "Нет",
                    item.file_name,
                    item.extension,
                    item.size_bytes,
                    item.modified_at.strftime("%Y-%m-%d %H:%M"),
                    item.status.value if hasattr(item.status, "value") else str(item.status),
                ),
            )

        self._update_counters()
        self.lbl_status.config(text=f"Найдено документов: {len(items)}")

    def _move_up(self):
        selected = self.tree.selection()
        for item in selected:
            index = self.tree.index(item)
            if index > 0:
                self.tree.move(item, self.tree.parent(item), index - 1)

    def _move_down(self):
        selected = self.tree.selection()
        for item in reversed(selected):
            index = self.tree.index(item)
            self.tree.move(item, self.tree.parent(item), index + 1)

    def _select_all(self):
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            values[0] = "Да"
            self.tree.item(item, values=values)
        self._update_counters()

    def _clear_all(self):
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            values[0] = "Нет"
            self.tree.item(item, values=values)
        self._update_counters()

    def _update_counters(self):
        total = len(self.tree.get_children())
        selected = sum(
            1
            for item in self.tree.get_children()
            if self.tree.item(item, "values")[0] == "Да"
        )
        self.lbl_counts.config(text=f"Найдено: {total} | Выбрано: {selected}")
