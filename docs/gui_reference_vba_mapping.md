# Карта соответствия VBA v1.1 и Python GUI

## Назначение

Документ фиксирует соответствие между референсной VBA-формой `Word Document Master Tool v1.1` и текущим Python/Tkinter GUI.

## Статусы

- implemented
- partial
- missing
- deferred

## Таблица соответствия

| VBA-блок | VBA-контрол / функция | Python GUI | Backend setting/service | Статус | Комментарий |
|---|---|---|---|---|---|
| Папки и результат | Исходная папка | `FolderSelectWidget` в `MainWindow` | `ToolSettings.source_folder`, `DocumentDiscovery` | partial | Папка используется для обновления списка, теперь передается в настройки. |
| Папки и результат | Папка результата | `FolderSelectWidget` в `MainWindow` | `ToolSettings.output_folder` | implemented | Используется при обработке, объединении, логировании и PDF. |
| Папки и результат | Имя файла | Поле `Имя файла` | `ToolSettings.output_file_name` | implemented | Значение передается из GUI в настройки. |
| Папки и результат | Обзор исходной папки | Кнопка `Обзор...` | `FolderSelectWidget._browse` | implemented | Реализовано через `filedialog.askdirectory`. |
| Папки и результат | Обзор папки результата | Кнопка `Обзор...` | `FolderSelectWidget._browse` | implemented | Реализовано через общий виджет выбора папки. |
| Папки и результат | Обновить список | Кнопка `Обновить список` | `DocumentDiscovery.find_documents` | implemented | Ищет `.docx`, `.docm`, `.doc`, `.rtf`. |
| Документы | Колонка `Вкл` | Колонка `Вкл` | `DocumentItem.is_selected` | implemented | Показывает `Да`/`Нет`; переключается кнопками и двойным кликом. |
| Документы | Колонка `Файл` | Колонка `Файл` | `DocumentItem.file_name` | implemented | Отображается в `DocumentTableWidget`. |
| Документы | Колонка `Тип` | Колонка `Тип` | `DocumentItem.extension` | implemented | Отображается расширение документа. |
| Документы | Колонка `Размер` | Колонка `Размер` | `DocumentItem.size_bytes` | implemented | Отображается человекочитаемый размер. |
| Документы | Колонка `Изменён` | Колонка `Изменён` | `DocumentItem.modified_at` | implemented | Отображается дата и время изменения. |
| Документы | Колонка `Статус` | Колонка `Статус` | `DocumentItem.status` | implemented | Отображается значение статуса. |
| Документы | Найдено | Счетчик `Найдено` | `GuiState.documents` | implemented | Обновляется после поиска и прогресса. |
| Документы | Выбрано | Счетчик `Выбрано` | `GuiState.selected_count` / подсчет в GUI | implemented | Отображает выбранные документы. |
| Документы | Проверить | Кнопка `Проверить` | `DocumentItem.is_valid`, `DocumentItem.status` | partial | Выполняет легкую GUI-проверку статусов без Word COM. |
| Документы | Вверх | Кнопка `Вверх` | Перестановка `GuiState.documents` | implemented | Перемещает выбранную строку вверх. |
| Документы | Вниз | Кнопка `Вниз` | Перестановка `GuiState.documents` | implemented | Перемещает выбранную строку вниз. |
| Документы | Вкл | Кнопка `Вкл` | `DocumentItem.is_selected` | implemented | Включает выделенные строки Treeview. |
| Документы | Выкл | Кнопка `Выкл` | `DocumentItem.is_selected` | implemented | Отключает выделенные строки Treeview. |
| Документы | Выбрать все | Кнопка `Выбрать все` | `DocumentItem.is_selected` | implemented | Массово включает все найденные документы. |
| Документы | Снять все | Кнопка `Снять все` | `DocumentItem.is_selected` | implemented | Массово отключает все найденные документы. |
| Настройки объединения | Формат | Combobox `Формат` | `MergeSettings.output_format` | implemented | Поддерживает `docx`, `docm`, `rtf`. |
| Настройки объединения | Режим | Combobox `Режим` | `MergeSettings.mode` | implemented | Значения GUI мапятся на числовые режимы сервиса. |
| Настройки объединения | Открыть после объединения | Чекбокс | `MergeSettings.open_after_merge` | implemented | Значение передается в настройки. |
| Настройки объединения | Создать отчет | Чекбокс `Создать отчёт` | `MergeSettings.create_report` | implemented | Значение передается в настройки. |
| Настройки объединения | Создать backup | Чекбокс `Создать backup` | `MergeSettings.create_backup` | implemented | Значение передается в настройки. |
| Исправления и комментарии | Принять все исправления | Чекбокс | `SourceProcessingSettings.accept_revisions` | implemented | Значение передается в настройки обработки. |
| Исправления и комментарии | Отключить Track Changes | Чекбокс | `SourceProcessingSettings.disable_track_changes` | implemented | Значение передается в настройки обработки. |
| Исправления и комментарии | Удалить комментарии | Чекбокс | `SourceProcessingSettings.remove_comments` | implemented | Значение передается в настройки обработки. |
| Исправления и комментарии | Предупреждать о защите | Чекбокс | `SourceProcessingSettings.warn_protected_docs` | implemented | Значение передается в настройки обработки. |
| Нумерация страниц | Включить нумерацию страниц | Чекбокс | `PageNumberingSettings.enabled` | implemented | Значение передается в настройки. |
| Нумерация страниц | Начать с | Spinbox `Начать с` | `PageNumberingSettings.start_number` | implemented | Значение нормализуется в `int`. |
| Нумерация страниц | Область | Combobox `Область` | `PageNumberingSettings.scope` | implemented | Значение передается в настройки. |
| Нумерация страниц | Место | Combobox `Место` | `PageNumberingSettings.location` | implemented | Значение передается в настройки. |
| Нумерация страниц | Выравнивание | Combobox `Выравнивание` | `PageNumberingSettings.alignment` | implemented | Значение передается в настройки. |
| Нумерация страниц | Формат | Combobox `Формат` | `PageNumberingSettings.format` | partial | GUI есть; применение формата в Word-сервисе требует проверки полноты. |
| Нумерация страниц | Шрифт | Поле `Шрифт` | `PageNumberingSettings.font_name` | implemented | Значение передается в настройки. |
| Нумерация страниц | Размер | Spinbox `Размер` | `PageNumberingSettings.font_size` | implemented | Значение нормализуется в `float`. |
| Нумерация страниц | Поля | Spinbox-поля верх/низ/лево/право | `PageNumberingSettings.*_margin_cm` | implemented | Значения нормализуются в `float`. |
| Нумерация страниц | Сквозная | Combobox `Режим` | `PageNumberingSettings.continuous` | implemented | Используется взаимоисключающий режим. |
| Нумерация страниц | Заново в каждом документе | Combobox `Режим` | `PageNumberingSettings.restart_each_document` | implemented | Используется взаимоисключающий режим. |
| Нумерация страниц | Удалить старые PAGE | Чекбокс | `PageNumberingSettings.remove_existing` | partial | GUI есть; применение в Word-сервисе требует проверки полноты. |
| Нумерация страниц | Сохранять колонтитулы | Combobox `Колонтитулы` | `PageNumberingSettings.preserve_headers_footers` | implemented | Используется взаимоисключающий режим. |
| Нумерация страниц | Очистить колонтитулы | Combobox `Колонтитулы` | `PageNumberingSettings.remove_headers_footers` | implemented | Используется взаимоисключающий режим. |
| Нумерация страниц | Настроить поля | Чекбокс | `PageNumberingSettings.adjust_margins` | implemented | Значение передается в настройки. |
| PDF-экспорт | PDF итогового документа | Чекбокс `Экспорт результата в PDF` | `PdfSettings.export_merged` | implemented | GUI пишет в модельное поле. |
| PDF-экспорт | PDF исходников без изменений | Чекбокс `Экспорт исходников в PDF` | `PdfSettings.export_sources`, `PdfExportService.export_source_to_pdf` | partial | Контрол есть, требуется проверить полный поток действия. |
| PDF-экспорт | PDF обработанных копий | Чекбокс | `PdfSettings.export_processed_copies` | implemented | Значение передается в настройки. |
| PDF-экспорт | Папка PDF | `FolderSelectWidget` | `PdfSettings.output_folder` | implemented | Используется как папка PDF при наличии значения. |
| PDF-экспорт | Обзор | Кнопка `Обзор...` | `FolderSelectWidget._browse` | implemented | Реализовано общим виджетом. |
| PDF-экспорт | Режим наименования | Combobox | `PdfSettings.naming_mode` | partial | GUI есть; backend-использование режимов требует проверки. |
| PDF-экспорт | Качество | Combobox `Качество` | `PdfSettings.quality`, optimize flags | implemented | Взаимоисключающий выбор `Печать`/`Экран`. |
| PDF-экспорт | Открыть PDF | Чекбокс | `PdfSettings.open_after_export` | partial | GUI есть; открытие после экспорта требует backend-проверки. |
| PDF-экспорт | PDF/A | Чекбокс | `PdfSettings.pdf_a` | implemented | Значение передается в Word PDF export. |
| PDF-экспорт | Свойства | Чекбокс | `PdfSettings.include_properties` | implemented | Значение передается в Word PDF export. |
| PDF-экспорт | Печать | Combobox `Качество` | `PdfSettings.optimize_for_print` | implemented | Взаимоисключающий выбор качества. |
| PDF-экспорт | Экран | Combobox `Качество` | `PdfSettings.optimize_for_screen` | implemented | Взаимоисключающий выбор качества. |
| PDF-экспорт | Создать общий PDF из созданных PDF | Чекбокс | `PdfSettings.merge_generated_pdfs`, `PdfMergeService` | implemented | Backend использует `pypdf`. |
| Маркеры и разделение | Добавлять маркеры частей | Чекбокс | `MarkerSettings.use_markers` | implemented | Значение передается в настройки. |
| Маркеры и разделение | Вид маркеров | Combobox | `MarkerSettings.visibility` | implemented | Значение мапится на числовой режим модели. |
| Маркеры и разделение | Режим удаления | Combobox | `MarkerSettings.removal_mode` | implemented | Значение мапится на числовой режим модели. |
| Маркеры и разделение | Удалить маркеры | Кнопка `Удалить маркеры` | Требуется подключение к `WordMarkerService` | partial | Кнопка показывает честный backend-stub. |
| Маркеры и разделение | Разделить | Кнопка `Разделить по маркерам` | `WordSplitService` | partial | Кнопка есть, но входной файл и статусы не оформлены полноценно. |
| Маркеры и разделение | Backup перед удалением | Чекбокс | `MarkerSettings.backup_before_removal` | implemented | Значение передается в настройки. |
| Маркеры и разделение | Полное удаление опасно | Warning label | GUI warning | implemented | Предупреждение показано в блоке маркеров. |
| Сноски и концевые сноски | Нумерация сносок | Чекбокс | `FootnoteSettings.enabled` | implemented | Значение передается в настройки. |
| Сноски и концевые сноски | Начать с | Spinbox | `FootnoteSettings.start_number` | implemented | Значение нормализуется в `int`. |
| Сноски и концевые сноски | Область | Combobox | `FootnoteSettings.scope` | implemented | Значение передается в настройки. |
| Сноски и концевые сноски | Режим | Combobox | `FootnoteSettings.mode` | implemented | Используется взаимоисключающий режим. |
| Сноски и концевые сноски | Формат | Combobox | `FootnoteSettings.format` | partial | GUI есть; применение формата в Word-сервисе требует проверки полноты. |
| Сноски и концевые сноски | Сквозная | Combobox `Режим` | `FootnoteSettings.continuous` | implemented | Используется взаимоисключающий режим. |
| Сноски и концевые сноски | Заново в документе | Combobox `Режим` | `FootnoteSettings.restart_each_document` | implemented | Используется взаимоисключающий режим. |
| Сноски и концевые сноски | Заново в секции | Combobox `Режим` | `FootnoteSettings.restart_each_section` | implemented | Используется взаимоисключающий режим. |
| Сноски и концевые сноски | Сохранять текст | Чекбокс | `FootnoteSettings.preserve_text` | implemented | Значение передается в настройки. |
| Сноски и концевые сноски | Концевые | Чекбокс | `FootnoteSettings.process_endnotes` | implemented | Значение передается в настройки. |
| Сноски и концевые сноски | Поля | Чекбокс | `FootnoteSettings.update_fields` | implemented | Значение передается в настройки. |
| Сноски и концевые сноски | Не заменять номера текстом | Чекбокс | `FootnoteSettings.do_not_replace_numbers_with_text` | implemented | Настройка добавлена в модель состояния. |
| Статус и действия | Статус | Нет отдельной строки статуса | Логи и статусы документов | partial | Есть прогресс и статус в таблице, нет общего status label. |
| Статус и действия | Лог / прогресс | Progressbar | `ProcessingLogger`, `setup_application_logging` | partial | Progressbar есть, просмотр лога в GUI отсутствует. |
| Статус и действия | Обработать файлы | Кнопка `Обработать файлы` | `SourceProcessingService` | partial | Кнопка есть, требуется проверка полного сценария. |
| Статус и действия | Объединить | Кнопка `Слияние документов` | `WordMergeService` | partial | Кнопка есть, зависит от Word COM и настроек. |
| Статус и действия | Сбросить | Нет кнопки | `GuiState.clear_documents` | missing | Метод есть, GUI-кнопка отсутствует. |
| Статус и действия | Закрыть | Нет кнопки | Tk root destroy | missing | Требуется отдельная кнопка закрытия. |
