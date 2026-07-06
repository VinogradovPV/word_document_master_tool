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
| Настройки объединения | Формат | Нет контрола | `MergeSettings.output_format` | missing | Модель есть, GUI отсутствует. |
| Настройки объединения | Режим | Нет контрола | `MergeSettings.mode` | missing | Модель есть, GUI отсутствует. |
| Настройки объединения | Открыть после объединения | Нет контрола | `MergeSettings.open_after_merge` | missing | Модель есть, GUI отсутствует. |
| Настройки объединения | Создать отчет | Нет контрола | `MergeSettings.create_report` | missing | Модель есть, GUI отсутствует. |
| Настройки объединения | Создать backup | Нет контрола | `MergeSettings.create_backup` | missing | Модель есть, GUI отсутствует. |
| Исправления и комментарии | Принять все исправления | Нет контрола | `SourceProcessingSettings.accept_revisions` | missing | Модель и сервис есть, GUI отсутствует. |
| Исправления и комментарии | Отключить Track Changes | Нет контрола | `SourceProcessingSettings.disable_track_changes` | missing | Модель и сервис есть, GUI отсутствует. |
| Исправления и комментарии | Удалить комментарии | Нет контрола | `SourceProcessingSettings.remove_comments` | missing | Модель и сервис есть, GUI отсутствует. |
| Исправления и комментарии | Предупреждать о защите | Нет контрола | `SourceProcessingSettings.warn_protected_docs` | missing | Модель есть, GUI отсутствует. |
| Нумерация страниц | Включить нумерацию страниц | Нет контрола | `PageNumberingSettings.enabled` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Начать с | Нет контрола | `PageNumberingSettings.start_number` | missing | Модель есть, GUI отсутствует. |
| Нумерация страниц | Область | Нет контрола | `PageNumberingSettings.scope` | missing | Модель есть, GUI отсутствует. |
| Нумерация страниц | Место | Нет контрола | `PageNumberingSettings.location` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Выравнивание | Нет контрола | `PageNumberingSettings.alignment` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Формат | Нет контрола | `PageNumberingSettings.format` | missing | Модель есть, сервис применяет не полностью. |
| Нумерация страниц | Шрифт | Нет контрола | `PageNumberingSettings.font_name` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Размер | Нет контрола | `PageNumberingSettings.font_size` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Поля | Нет контрола | `PageNumberingSettings.*_margin_cm` | missing | Модель и сервис есть, GUI отсутствует. |
| Нумерация страниц | Сквозная | Нет контрола | `PageNumberingSettings.continuous` | missing | Модель и валидация есть, GUI отсутствует. |
| Нумерация страниц | Заново в каждом документе | Нет контрола | `PageNumberingSettings.restart_each_document` | missing | Модель и валидация есть, GUI отсутствует. |
| Нумерация страниц | Удалить старые PAGE | Нет контрола | `PageNumberingSettings.remove_existing` | missing | Модель есть, сервис требует проверки полноты. |
| Нумерация страниц | Сохранять колонтитулы | Нет контрола | `PageNumberingSettings.preserve_headers_footers` | missing | Модель и валидация есть, GUI отсутствует. |
| Нумерация страниц | Очистить колонтитулы | Нет контрола | `PageNumberingSettings.remove_headers_footers` | missing | Модель, сервис и валидация есть, GUI отсутствует. |
| Нумерация страниц | Настроить поля | Нет контрола | `PageNumberingSettings.adjust_margins` | missing | Модель и сервис есть, GUI отсутствует. |
| PDF-экспорт | PDF итогового документа | Чекбокс `Экспорт результата в PDF` | `PdfSettings.export_merged` | implemented | GUI пишет в модельное поле. |
| PDF-экспорт | PDF исходников без изменений | Чекбокс `Экспорт исходников в PDF` | `PdfSettings.export_sources`, `PdfExportService.export_source_to_pdf` | partial | Контрол есть, требуется проверить полный поток действия. |
| PDF-экспорт | PDF обработанных копий | Нет контрола | `PdfSettings.export_processed_copies` | missing | Модель есть, GUI отсутствует. |
| PDF-экспорт | Папка PDF | `FolderSelectWidget` | `PdfSettings.output_folder` | implemented | Используется как папка PDF при наличии значения. |
| PDF-экспорт | Обзор | Кнопка `Обзор...` | `FolderSelectWidget._browse` | implemented | Реализовано общим виджетом. |
| PDF-экспорт | Режим наименования | Нет контрола | `PdfSettings.naming_mode` | missing | Модель есть, GUI отсутствует. |
| PDF-экспорт | Качество | Нет контрола | `PdfSettings.quality`, optimize flags | missing | Модель и валидация есть, GUI отсутствует. |
| PDF-экспорт | Открыть PDF | Нет контрола | `PdfSettings.open_after_export` | missing | Модель есть, GUI отсутствует. |
| PDF-экспорт | PDF/A | Нет контрола | `PdfSettings.pdf_a` | missing | Модель и сервис есть, GUI отсутствует. |
| PDF-экспорт | Свойства | Нет контрола | `PdfSettings.include_properties` | missing | Модель и сервис есть, GUI отсутствует. |
| PDF-экспорт | Печать | Нет контрола | `PdfSettings.optimize_for_print` | missing | Модель и валидация есть, GUI отсутствует. |
| PDF-экспорт | Экран | Нет контрола | `PdfSettings.optimize_for_screen` | missing | Модель и валидация есть, GUI отсутствует. |
| PDF-экспорт | Создать общий PDF из созданных PDF | Нет контрола | `PdfSettings.merge_generated_pdfs`, `PdfMergeService` | missing | Backend через `pypdf` есть, GUI отсутствует. |
| Маркеры и разделение | Добавлять маркеры частей | Нет контрола | `MarkerSettings.use_markers` | missing | Модель есть, GUI отсутствует. |
| Маркеры и разделение | Вид маркеров | Нет контрола | `MarkerSettings.visibility` | missing | Модель есть, GUI отсутствует. |
| Маркеры и разделение | Режим удаления | Нет контрола | `MarkerSettings.removal_mode` | missing | Модель есть, GUI отсутствует. |
| Маркеры и разделение | Удалить маркеры | Нет кнопки | `WordMarkerService` | missing | Сервис есть, GUI отсутствует. |
| Маркеры и разделение | Разделить | Кнопка `Разделить по маркерам` | `WordSplitService` | partial | Кнопка есть, но входной файл и статусы не оформлены полноценно. |
| Маркеры и разделение | Backup перед удалением | Нет контрола | `MarkerSettings.backup_before_removal` | missing | Модель есть, GUI отсутствует. |
| Маркеры и разделение | Полное удаление опасно | Нет предупреждения | Требуется GUI validation/warning | missing | Требуется отдельное предупреждение. |
| Сноски и концевые сноски | Нумерация сносок | Нет контрола | `FootnoteSettings.enabled` | missing | Модель и сервис есть, GUI отсутствует. |
| Сноски и концевые сноски | Начать с | Нет контрола | `FootnoteSettings.start_number` | missing | Модель и сервис есть, GUI отсутствует. |
| Сноски и концевые сноски | Область | Нет контрола | `FootnoteSettings.scope` | missing | Модель есть, GUI отсутствует. |
| Сноски и концевые сноски | Режим | Нет контрола | `FootnoteSettings.mode` | missing | Модель есть, GUI отсутствует. |
| Сноски и концевые сноски | Формат | Нет контрола | `FootnoteSettings.format` | missing | Модель есть, сервис применяет не полностью. |
| Сноски и концевые сноски | Сквозная | Нет контрола | `FootnoteSettings.continuous` | missing | Модель и валидация есть, GUI отсутствует. |
| Сноски и концевые сноски | Заново в документе | Нет контрола | `FootnoteSettings.restart_each_document` | missing | Модель и валидация есть, GUI отсутствует. |
| Сноски и концевые сноски | Заново в секции | Нет контрола | `FootnoteSettings.restart_each_section` | missing | Модель, сервис и валидация есть, GUI отсутствует. |
| Сноски и концевые сноски | Сохранять текст | Нет контрола | `FootnoteSettings.preserve_text` | missing | Модель есть, GUI отсутствует. |
| Сноски и концевые сноски | Концевые | Нет контрола | `FootnoteSettings.process_endnotes` | missing | Модель и сервис есть, GUI отсутствует. |
| Сноски и концевые сноски | Поля | Нет контрола | `FootnoteSettings.update_fields` | missing | Модель есть, GUI отсутствует. |
| Сноски и концевые сноски | Не заменять номера текстом | Нет контрола | Требуется настройка | missing | В модели нет отдельного поля. |
| Статус и действия | Статус | Нет отдельной строки статуса | Логи и статусы документов | partial | Есть прогресс и статус в таблице, нет общего status label. |
| Статус и действия | Лог / прогресс | Progressbar | `ProcessingLogger`, `setup_application_logging` | partial | Progressbar есть, просмотр лога в GUI отсутствует. |
| Статус и действия | Обработать файлы | Кнопка `Обработать файлы` | `SourceProcessingService` | partial | Кнопка есть, требуется проверка полного сценария. |
| Статус и действия | Объединить | Кнопка `Слияние документов` | `WordMergeService` | partial | Кнопка есть, зависит от Word COM и настроек. |
| Статус и действия | Сбросить | Нет кнопки | `GuiState.clear_documents` | missing | Метод есть, GUI-кнопка отсутствует. |
| Статус и действия | Закрыть | Нет кнопки | Tk root destroy | missing | Требуется отдельная кнопка закрытия. |
