# TSV-Converter — Code Navigation Map
*Для LLM-агентов: полная карта проекта для быстрой навигации и понимания архитектуры*

> **Путь проекта:** `D:\Синхронизация\Git\TSV-Converter\`  
> **Сгенерировано:** 2026-04-27  
> **Для кого:** Другие LLM-агенты, которым нужно ориентироваться в коде

---

## 1. PROJECT OVERVIEW

TSV-Converter — десктопное приложение (PySide6/Qt6 + xlsxwriter) для конвертации TSV/CSV файлов в Excel (XLSX/CSV) с расширенной обработкой данных.

**Ключевой функционал:**
- Пакетная конвертация с поддержкой drag-and-drop
- Выбор и переупорядочивание столбцов вывода ("Столбцы вывода")
- Фильтрация и независимое разделение по значениям столбцов: отдельно на файлы и отдельно на листы внутри файлов
- Генерация сводных таблиц с настраиваемой агрегацией
- Асинхронная обработка (фоновые потоки) для отзывчивости UI
- Светлая/тёмная темы с нативным рендерингом виджетов
- Сохранение сессии (filter/split/column settings)
- Постоянные настройки через QSettings

---

## 2. FILE MAP

| Файл | Строк | Ключевые классы | Назначение |
|-------|-------|-----------------|------------|
| `main.py` | 896 | `TSVConverterApp` | Точка входа, маршрутизация сигналов GUI ↔ бизнес-логика, QSettings, оркестрация конвертации |
| `gui.py` | 3693 | `StylePalette`, `StyledCheckBox`, `StyledButton`, `PrimaryButton`, `SecondaryButton`, `DangerButton`, `FileListWidget`, `LogTextEdit`, `LoadingOverlay`, `LoadingWorker`, `UniqueValuesWorker*`, `SettingsDialog`, `CheckBoxListWidget`, `CheckBoxListDelegate`, `ColumnCheckBoxListWidget`, `ColumnValuesDialog`, `ColumnSelectionDialog`, `PivotSettingsDialog`, `PivotPreviewDialog`, `TSVPreviewDialog`, `MainWindow` | Весь UI, кастомные виджеты, диалоги, темы, обёртки для фоновых задач |
| `converter.py` | 2001 | `ConversionConfig`, `ProgressData`, `ProgressTracker`, `FileUtilities`, `TSVToExcelConverter`, `PivotTableProcessor` | Движок конвертации, бизнес-логика, утилиты работы с файлами, сводные таблицы |

---

## 3. CLASS HIERARCHY

```
QMainWindow
└── MainWindow                        (gui.py:2710)

QDialog
├── SettingsDialog                   (gui.py:854)
├── ColumnValuesDialog               (gui.py:1194)
├── ColumnSelectionDialog            (gui.py:1423)
├── PivotSettingsDialog              (gui.py:1617)
├── PivotPreviewDialog               (gui.py:1904)
└── TSVPreviewDialog                 (gui.py:2329)

QWidget
├── StylePalette                     (gui.py:75)        [static utility]
├── StyledCheckBox                  (gui.py:200)
├── LoadingOverlay                  (gui.py:626)
├── FileListWidget                  (gui.py:347)
└── LogTextEdit                     (gui.py:398)

QPushButton
└── StyledButton                    (gui.py:303)
    ├── PrimaryButton               (gui.py:325)
    ├── SecondaryButton             (gui.py:333)
    └── DangerButton                (gui.py:340)

QListWidget
├── CheckBoxListWidget              (gui.py:974)
└── ColumnCheckBoxListWidget        (gui.py:1157)

QThread
├── TSVToExcelConverter             (converter.py:404)
├── LoadingWorker                   (gui.py:748)
└── UniqueValuesWorker              (gui.py:779)

QStyledItemDelegate
└── CheckBoxListDelegate            (gui.py:1016)

QObject
└── UniqueValuesWorkerSignals       (gui.py:772)

dataclass
├── ConversionConfig                (converter.py:26)
└── ProgressData                    (converter.py:48)
```

### Зависимости между классами
- `TSVConverterApp` (main.py) → `MainWindow`, `TSVToExcelConverter`, `FileUtilities`, `ConversionConfig`, `ColumnValuesDialog`, `PivotSettingsDialog`, `TSVPreviewDialog`
- `MainWindow` → все кастомные виджеты в gui.py, `ConversionConfig`, `FileUtilities`, `LoadingWorker`
- `TSVToExcelConverter` → `ConversionConfig`, `ProgressData`, `ProgressTracker`, `FileUtilities`, `PivotTableProcessor`, `xlsxwriter`
- `PivotTableProcessor` → `FileUtilities`

---

## 4. SIGNAL ARCHITECTURE

### Пользовательские сигналы

| Класс | Сигнал | Сигнатура | Назначение |
|-------|--------|-----------|------------|
| `MainWindow` (gui.py:2716-2726) | `conversion_started` | `Signal(object)` (ConversionConfig) | Запуск конвертации |
| | `conversion_stopped` | `Signal()` | Остановка конвертации |
| | `files_added` | `Signal(object)` (list of paths) | Файлы добавлены в список |
| | `preview_requested` | `Signal()` | Предпросмотр файла |
| | `pivot_settings_requested` | `Signal()` | Диалог сводной таблицы |
| | `open_converted_file_requested` | `Signal()` | Открыть конвертированный файл |
| | `delete_converted_file_requested` | `Signal()` | Удалить конвертированный файл |
| | `export_report_requested` | `Signal()` | Экспорт отчёта |
| | `settings_saved` | `Signal(dict)` | Настройки сохранены |
| | `columns_changed` | `Signal()` | Изменён выбор столбцов |
| `FileListWidget` (gui.py:352) | `files_dropped` | `Signal(object)` (list of paths) | Drag-and-drop файлов |
| `TSVToExcelConverter` (converter.py:416-421) | `update_progress` | `Signal(int)` | Прогресс (0-100%) |
| | `progress_data` | `Signal(object)` (ProgressData) | Детальные метрики прогресса |
| | `log_message` | `Signal(str, QColor)` | Сообщение лога из конвертера |
| | `finished_signal` | `Signal()` | Конвертация завершена |
| | `stopped_signal` | `Signal()` | Конвертация остановлена |
| | `error` | `Signal(str)` | Ошибка конвертации |
| `LoadingWorker` (gui.py:754-755) | `finished` | `Signal(object)` | Фоновая задача завершена |
| | `error` | `Signal(str)` | Ошибка фоновой задачи |
| `UniqueValuesWorkerSignals` (gui.py:775-776) | `finished` | `Signal(object)` (set) | Уникальные значения загружены |
| | `error` | `Signal(str)` | Ошибка загрузки значений |

### Ключевые соединения сигналов (main.py)

```
MainWindow.conversion_started        → TSVConverterApp._start_conversion
MainWindow.conversion_stopped       → TSVConverterApp._stop_conversion
MainWindow.files_added              → TSVConverterApp._on_files_added
MainWindow.preview_requested        → TSVConverterApp._preview_file
MainWindow.pivot_settings_requested  → TSVConverterApp._show_pivot_settings
MainWindow.settings_saved           → TSVConverterApp._on_settings_saved
MainWindow.columns_changed          → TSVConverterApp._on_columns_changed

TSVToExcelConverter.progress_data   → MainWindow.update_progress_details
TSVToExcelConverter.log_message     → TSVConverterApp._log_message
TSVToExcelConverter.finished_signal → TSVConverterApp._on_conversion_finished
TSVToExcelConverter.stopped_signal → TSVConverterApp._on_conversion_stopped
TSVToExcelConverter.error           → TSVConverterApp._on_conversion_error
```

---

## 5. UI WIDGET MAP

| Виджет | Файл:Строка | Базовый класс | Роль |
|--------|-------------|---------------|------|
| `StylePalette` | gui.py:75 | object | Централизованное хранение цветов темы |
| `StyledCheckBox` | gui.py:200 | QCheckBox | Кастомный чекбокс с ручной отрисовкой для совместимости с темами |
| `StyledButton` | gui.py:303 | QPushButton | Базовая стилизованная кнопка |
| `PrimaryButton` | gui.py:325 | StyledButton | Основная кнопка действия (напр., "Начать конвертацию") |
| `SecondaryButton` | gui.py:333 | StyledButton | Вторичная кнопка действия |
| `DangerButton` | gui.py:340 | StyledButton | Опасная кнопка (удалить/закрыть) |
| `FileListWidget` | gui.py:347 | QListWidget | Список файлов с поддержкой drag-and-drop |
| `LogTextEdit` | gui.py:398 | QTextEdit | Отображение лога с автоскроллом и цветами темы |
| `LoadingOverlay` | gui.py:626 | QWidget | Анимированный оверлей загрузки со спиннером |
| `LoadingWorker` | gui.py:748 | QThread | Универсальный исполнитель фоновых задач |
| `UniqueValuesWorkerSignals` | gui.py:772 | QObject | Контейнер сигналов для UniqueValuesWorker |
| `UniqueValuesWorker` | gui.py:779 | QThread | Асинхронная загрузка уникальных значений столбца |
| `SettingsDialog` | gui.py:854 | QDialog | Настройки приложения (тема, пути, поведение) |
| `CheckBoxListWidget` | gui.py:974 | QListWidget | Список с чекбоксами, хранит состояние в `_check_states` |
| `CheckBoxListDelegate` | gui.py:1016 | QStyledItemDelegate | Кастомная отрисовка чекбоксов в списках |
| `ColumnCheckBoxListWidget` | gui.py:1157 | QListWidget | Список чекбоксов для выбора столбцов с drag-drop; хранит состояние в `item.data(_CHECK_STATE_ROLE)` |
| `ColumnValuesDialog` | gui.py:1194 | QDialog | Выбор значений столбца (split/filter) |
| `ColumnSelectionDialog` | gui.py:1423 | QDialog | Настройка столбцов вывода и их порядка |
| `PivotSettingsDialog` | gui.py:1617 | QDialog | Настройка сводной таблицы |
| `PivotPreviewDialog` | gui.py:1904 | QDialog | Предпросмотр сводной таблицы |
| `TSVPreviewDialog` | gui.py:2329 | QDialog | Предпросмотр содержимого TSV/CSV с пагинацией |
| `MainWindow` | gui.py:2710 | QMainWindow | Главное окно приложения |

---

## 6. DATA FLOW (Вход → Выход)

### Шаг за шагом

1. **Ввод файлов**
   - Пользователь добавляет TSV/CSV через `FileListWidget` (drag-drop или кнопка "Добавить файлы")
   - Сигнал `FileListWidget.files_dropped` → `MainWindow._add_files()`
   - `LoadingWorker` запускает `_process_files_operation` в фоне:
     - Определяет кодировку (`FileUtilities.get_encoding`)
     - Определяет разделитель (`FileUtilities.get_delimiter`)
     - Считает строки (`FileUtilities.count_rows`)
     - Извлекает заголовки из первого файла
   - Заголовки заполняют `split_column_combo`, `filter_column_combo`, активируют `ColumnSelectionDialog`

2. **Конфигурация**
   - Пользователь задаёт формат вывода (XLSX/CSV), папку, стили
   - Опционально настраивает:
     - Столбец разделения на файлы — вывод по уникальным значениям в отдельные файлы
     - Столбец разделения на листы — вывод по уникальным значениям в листы внутри каждого файла
     - Столбец фильтра (filter) — фильтрация строк по выбранным значениям
     - Столбцы вывода через диалог "Столбцы вывода"
     - Сводную таблицу через диалог "Сводная таблица"

3. **Запуск конвертации**
   - Кнопка "Начать конвертацию" → `MainWindow.conversion_started` передаёт `ConversionConfig`
   - `TSVConverterApp` создаёт `TSVToExcelConverter` (QThread) с конфигом
   - `TSVToExcelConverter.start()` начинает фоновую обработку

4. **Фоновая обработка (TSVToExcelConverter.run())**
   - Подсчёт строк во всех файлах (с учётом filter/split/выбора столбцов)
   - Инициализация `ProgressTracker` для метрик
   - Для каждого файла:
     - Чтение с определённой кодировкой/разделителем
     - Извлечение заголовков, проверка индексов split/filter
     - Получение выходных столбцов через `_get_output_columns()` (применяет выбор столбцов)
     - Обработка строк:
      - Применение фильтра (пропуск строк, не совпадающих с выбранными значениями)
      - Логика разделения на файлы (если выбран `file_split_column`)
      - Логика разделения на листы внутри файла/файлов (если выбран `split_column`)
       - Проекция строк на выбранные столбцы через `_project_row()`
       - Дедупликация строк, если выбор столбцов исключает колонки
       - Отправка обновлений прогресса каждые N строк
   - Генерация сводной таблицы если настроена (`PivotTableProcessor.create_pivot_data()`)
   - Запись в XLSX (через `xlsxwriter`) или CSV

5. **Вывод**
   - Прогресс отображается в progress bar и панели деталей `MainWindow`
   - Сообщения лога выводятся в `LogTextEdit`
   - По завершении: кнопки "Открыть файл" и "Удалить файл" активируются
   - Авто-открытие через `QDesktopServices.openUrl()` если настроено
   - Авто-удаление удаляет исходные файлы если настроено

---

## 7. COLUMN SELECTION DEEP DIVE ("Столбцы вывода")

### Обзор фичи
Позволяет пользователю выбрать, какие столбцы попадут в выходной файл, и изменить порядок их отображения. Доступно только когда загружен **ровно 1 файл** (отключено для нескольких файлов).

### Реализация

1. **Триггер**: `MainWindow.columns_btn` (gui.py:2977) → открывает `ColumnSelectionDialog`
2. **Диалог**: `ColumnSelectionDialog` (gui.py:1423)
   - Использует `ColumnCheckBoxListWidget` (gui.py:1157) для чекбоксов с drag-drop переупорядочиванием
   - Столбцы отображаются в порядке текущего выбора, ненужные добавляются в конец
   - Функционал: поиск, "Выбрать все", "Снять все", "Сбросить порядок"
3. **Хранение состояния**: Выбранные столбцы в `MainWindow._selected_output_columns` (список имён столбцов)
4. **Интеграция с конвертацией**:
   - Выбранные столбцы передаются в `ConversionConfig.selected_columns`
   - `TSVToExcelConverter._get_output_columns()` (converter.py:585):
     - Сопоставляет имена столбцов с индексами заголовков
     - Возвращает отфильтрованные заголовки и индексы вывода
     - Пишет предупреждение в лог, если выбранные столбцы не найдены в файле
   - `TSVToExcelConverter._project_row()` (converter.py:625): возвращает только выбранные столбцы для каждой строки
5. **Дедупликация**: Автоматически включается, когда выбранные столбцы составляют меньше всех заголовков:
   - Ключи дедупликации — кортежи значений только в выбранных столбцах
   - Удаляет строки, идентичные по выбранным столбцам
   - Отслеживается счётчиком `TSVToExcelConverter.duplicates_removed`

### Важные ограничения
- Активируется только при `file_list.count() == 1` (gui.py:3420-3427)
- Выбор столбцов сбрасывается при добавлении новых файлов (gui.py:3236-3244)
- File split / sheet split / filter комбо-боксы показывают только выбранные столбцы вывода при загрузке 1 файла

### Технические детали хранения состояния чекбоксов
`ColumnCheckBoxListWidget` хранит состояние чекбокса в кастомной роли item data:
- Роль: `Qt.ItemDataRole.UserRole.value + 1` (константа `_CHECK_STATE_ROLE`)
- Чтение: `item.data(_CHECK_STATE_ROLE)` → `Qt.CheckState`
- Запись: `item.setData(_CHECK_STATE_ROLE, state)`
- **Почему не `item.checkState()`**: флаг `ItemIsUserCheckable` удалён для предотвращения конфликтов с кастомным `CheckBoxListDelegate`, поэтому нативный `checkState()` всегда возвращает `Unchecked`

---

## 8. WORKER PATTERN (Асинхронные операции)

### Реализации воркеров

1. **LoadingWorker** (gui.py:748)
   - Универсальная обёртка QThread для любого вызываемого объекта
   - Использование: подсчёт строк, загрузка значений столбцов, сборка превью сводной таблицы, обработка добавленных файлов
   - Сигналы: `finished` (возвращает результат вызова), `error` (возвращает строку ошибки)

2. **UniqueValuesWorker** (gui.py:779)
   - Специализированный QThread для загрузки уникальных значений столбца
   - Использует отдельный класс `UniqueValuesWorkerSignals` (gui.py:772) для избежания проблем владения сигналами
   - Обрабатывает несколько файлов, применяет опциональный фильтр, возвращает множество уникальных значений
   - Используется в `ColumnValuesDialog` для асинхронной загрузки значений

3. **TSVToExcelConverter** (converter.py:404)
   - Основной поток конвертации (наследует QThread)
   - Длительная задача: обрабатывает все входные файлы, пишет вывод
   - Отправляет детальный прогресс через сигнал `progress_data` (ProgressData)
   - Флаг остановки проверяется каждые `STOP_CHECK_INTERVAL` (2000) строк для возможности отмены

4. **ProgressTracker** (converter.py:112)
   - Не поток; вычисляет метрики прогресса:
     - Процент завершения, затраченное время, ETA, строк в секунду
     - Дросселирование до обновления не чаще 0.1 секунды
   - Используется `TSVToExcelConverter` для генерации ProgressData

### Жизненный цикл воркеров
- `TSVConverterApp` хранит ссылки на активные воркеры в списке `_active_workers` (main.py:48), чтобы предотвратить сборку мусора работающих потоков (избегает ошибок "QThread: Destroyed")
- Воркеры удаляются из списка по завершении или ошибке
- Оверлеи загрузки (`LoadingOverlay`) показываются во время выполнения воркера

---

## 9. СИСТЕМА НАСТРОЕК И СОХРАНЕНИЯ

### Постоянное хранение (QSettings)
- Расположение: Реестр (Windows) или plist (macOS) через `QSettings("TSVConverter", "App")`
- Управляется в `TSVConverterApp` (main.py:35, 87-123)

| Ключ | Тип | По умолчанию | Описание |
|------|-----|---------------|----------|
| `theme` | str | `"Светлая"` | `"Светлая"` или `"Тёмная"` |
| `default_path` | str | `""` | Папка вывода по умолчанию |
| `auto_open` | str (`"true"`/`"false"`) | `"false"` | Авто-открытие файла после конвертации |
| `auto_delete` | str (`"true"`/`"false"`) | `"false"` | Авто-удаление исходных файлов после конвертации |
| `ram_threshold` | str | `"500000"` | Порог строк для режима `constant_memory` |

### Состояние сессии (не сохраняется)
Хранится в переменных экземпляра `MainWindow` (с префиксом `_`):
- `_settings`: Dict с текущими настройками приложения (синхронизировано с QSettings)
- `_filter_values`: `Dict[str, List[str]]` → имя столбца → выбранные значения фильтра
- `_file_split_values`: `Dict[str, List[str]]` → столбец разделения на файлы → выбранные значения
- `_sheet_split_values`: `Dict[str, List[str]]` → столбец разделения на листы → выбранные значения
- `_pivot_settings`: `Optional[Dict]` → конфигурация сводной таблицы
- `_available_headers`: `List[str]` → заголовки из загруженного файла(ов)
- `_selected_output_columns`: `List[str]` → выбранные столбцы вывода

### Применение темы
- Используется нативная модификация `QPalette` вместо QSS для кросс-платформенной согласованности
- `StylePalette` (gui.py:75) предоставляет карты цветов для светлой/тёмной тем
- `apply_native_style()` (gui.py:116) рекурсивно применяет палитру ко всем дочерним виджетам
- Кастомная отрисовка чекбоксов (`StyledCheckBox`, `CheckBoxListDelegate`) необходима, так как Windows 11 переопределяет стандартные стили QSS чекбоксов

---

## 10. ГЕНЕРАЦИЯ СВОДНОЙ ТАБЛИЦЫ

### Конфигурация
1. Кнопка "Сводная таблица" → `PivotSettingsDialog` (gui.py:1617)
2. Настройка:
   - Фильтры: Столбцы для фильтрации (использует текущие настройки фильтра)
   - Строки: Столбцы для группировки в строках сводной таблицы
   - Столбцы: Столбцы для группировки в столбцах сводной таблицы
   - Значения: Столбцы для агрегации с методом (`Сумма`/`Среднее`/`Количество`/`Максимум`/`Минимум`)
   - Опция удаления дубликатов
3. Настройки хранятся в `MainWindow._pivot_settings`

### Генерация
1. Во время конвертации: `TSVToExcelConverter` использует `PivotTableProcessor` (converter.py:1709)
2. `PivotTableProcessor.create_pivot_data()`:
   - Читает входной файл, извлекает индексы строк/столбцов/значений
   - Применяет фильтр и дедупликацию если настроено
   - Агрегирует значения онлайн (сумма, количество, макс, мин) по группе строка+столбец
   - Возвращает вложенный dict: `{row_key: {col_key: {metric_key: aggregated_value}}}`
3. `PivotTableProcessor.write_pivot_table()`:
   - Пишет на отдельный лист (`"Сводная таблица"`) или в отдельный файл
   - Форматирует заголовки с акцентным цветом, выравнивает значения по правому краю
   - Авто-размер столбцов до ширины 15

### Предпросмотр
- `PivotPreviewDialog` (gui.py:1904) генерирует данные сводной таблицы асинхронно через `LoadingWorker`
- Отображает в `QTableWidget` с возможностью копирования в буфер и экспорта в Excel

---

## 11. KNOWN GOTCHAS (Точки путаницы для LLM)

1. **Паттерны хранения состояния**:
   - Многие переменные состояния используют защищённый префикс `_` в `MainWindow`, но доступаются напрямую из `main.py` (напр., `self.window._settings`)
   - Состояние сессии не сохраняется в QSettings: выбор столбцов, значения фильтров, настройки split теряются при перезапуске

2. **Сборка мусора воркеров**:
   - Необходимо хранить ссылки на активные воркеры (напр., `self._active_workers` в main.py), чтобы предотвратить ошибки уничтожения QThread
   - Воркеры не привязываются к главному окну, чтобы избежать проблем с потоками

3. **Ограничения выбора столбцов**:
   - Доступно только для ровно 1 загруженного файла (комбо-боксы отключены для нескольких файлов)
   - Столбцы file split / sheet split / filter ограничены выбранными столбцами вывода при загрузке 1 файла

4. **Логика дедупликации**:
   - Автоматически включается, когда выбор столбцов исключает некоторые колонки (дедупликация по выбранным столбцам только)
   - Дедупликация не применяется, если выбраны все столбцы

5. **Режимы xlsxwriter**:
   - Режим `constant_memory` используется, когда общее число строк >= `ram_threshold` (по умолчанию 500k) для снижения потребления RAM
   - Обычный режим (временные файлы в памяти) быстрее для малых наборов данных
   - `in_memory: True` явно избегается из-за медленного сжатия при закрытии

6. **Рендеринг темы**:
   - Кастомная отрисовка чекбоксов (`StyledCheckBox`, `CheckBoxListDelegate`) необходима, потому что Windows 11 переопределяет стандартные стили QSS чекбоксов
   - Нативная палитра используется вместо QSS для согласованной темизации

7. **Хранение в FileListWidget**:
   - Количество строк на файл хранится в user data `QListWidgetItem` (`Qt.UserRole`) для быстрого подсчёта общего числа строк
   - Drag-and-drop принимает только файлы `.tsv`, `.csv`, `.txt`

8. **Именование сигналов**:
   - `TSVToExcelConverter` использует `finished_signal` (не `finished`), чтобы избежать конфликта со встроенным сигналом QThread
   - `MainWindow` использует `conversion_started` (не `conversion_start`) для ясности

9. **Отслеживание прогресса**:
   - Dataclass `ProgressData` используется для детального прогресса, в то время как сигнал `update_progress` сохранён для обратной совместимости
   - Обновления прогресса дросселируются до 0.1 секунды во избежание перегрузки UI

10. **Особенности многих файлов**:
    - Функции file split / sheet split / фильтрации / выбора столбцов отключены при загрузке нескольких файлов
    - Общее количество строк для нескольких файлов суммирует строки всех файлов (без кросс-файловой дедупликации)

---

## 12. БЫСТРАЯ НАВИГАЦИЯ ДЛЯ ЗАДАЧ

| Задача | Где искать |
|--------|------------|
| Изменить внешний вид кнопок | `StyledButton`, `PrimaryButton`, `SecondaryButton`, `DangerButton` (gui.py:303-342) |
| Изменить тему / цвета | `StylePalette` (gui.py:75), `apply_native_style()` (gui.py:116) |
| Исправить чекбоксы / списки с галочками | `CheckBoxListWidget` (gui.py:974), `CheckBoxListDelegate` (gui.py:1016), `ColumnCheckBoxListWidget` (gui.py:1157) |
| Изменить логику выбора столбцов | `ColumnSelectionDialog` (gui.py:1423), `MainWindow._show_column_selection` (gui.py:3398) |
| Исправить split/filter | `ColumnValuesDialog` (gui.py:1194), `MainWindow._update_column_combos` (gui.py:3222), `TSVConverterApp._on_split_dimension_selected` (main.py) |
| Изменить логику конвертации | `TSVToExcelConverter.run()` (converter.py:435), `_process_file()` (converter.py:500) |
| Изменить сводные таблицы | `PivotTableProcessor` (converter.py:1709), `PivotSettingsDialog` (gui.py:1617) |
| Добавить новый диалог | Следовать паттерну существующих QDialog в gui.py |
| Изменить формат вывода Excel | `TSVToExcelConverter._write_excel()` (converter.py:700) |
| Исправить drag-and-drop | `FileListWidget` (gui.py:347), `dragEnterEvent`/`dropEvent` |
| Добавить новый воркер | Следовать паттерну `LoadingWorker` (gui.py:748) или `UniqueValuesWorker` (gui.py:779) |
| Изменить постоянные настройки | `TSVConverterApp` (main.py:35), `_load_settings` (main.py:87), `_save_settings` (main.py:108) |
