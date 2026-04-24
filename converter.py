"""
Модуль бизнес-логики для конвертера TSV/CSV в Excel.
Содержит классы для конвертации, утилиты и вспомогательные функции.
"""

import os
import csv
import time
from datetime import datetime
from typing import List, Dict, Any, Optional, Set
from collections import defaultdict
from dataclasses import dataclass, field

from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QColor

import xlsxwriter


# ============================================================================
# ТИПЫ ДАННЫХ ДЛЯ ПРОГРЕССА
# ============================================================================


@dataclass
class ConversionConfig:
    """Конфигурация конвертации, собираемая из GUI."""

    input_files: List[str]
    output_directory: str
    output_format: str
    auto_open: bool
    auto_delete: bool
    styles: Dict[str, Any]
    header_color: str
    split_column: str
    split_mode: str  # "sheets" или "files"
    selected_values: List[str]
    filter_column: str
    filter_values: List[str]
    pivot_settings: Dict[str, Any]
    selected_columns: List[str] = field(default_factory=list)
    deduplicate_rows: bool = True
    ram_threshold: int = 500000


@dataclass
class ProgressData:
    """
    Структура данных для расширенной информации о прогрессе.

    Attributes:
        percent: Процент выполнения (0-100)
        processed_rows: Обработано строк
        total_rows: Всего строк
        elapsed_seconds: Прошло времени (секунды)
        eta_seconds: Осталось времени (секунды, ETA)
        rows_per_second: Скорость обработки (строк/сек)
        current_file: Текущий обрабатываемый файл
        current_operation: Текущая операция (текст)
    """

    percent: int = 0
    processed_rows: int = 0
    total_rows: int = 0
    elapsed_seconds: float = 0.0
    eta_seconds: float = 0.0
    rows_per_second: float = 0.0
    current_file: str = ""
    current_operation: str = ""

    def format_eta(self) -> str:
        """Форматирует ETA в читаемый вид (MM:SS или HH:MM:SS)."""
        if self.eta_seconds < 0:
            return "--:--"

        total_seconds = int(self.eta_seconds)
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60

        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def format_elapsed(self) -> str:
        """Форматирует прошедшее время в читаемый вид."""
        total_seconds = int(self.elapsed_seconds)
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60

        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def format_speed(self) -> str:
        """Форматирует скорость обработки."""
        if self.rows_per_second < 1:
            return f"{self.rows_per_second:.1f} строк/сек"
        elif self.rows_per_second < 1000:
            return f"{int(self.rows_per_second)} строк/сек"
        else:
            return f"{self.rows_per_second / 1000:.1f}K строк/сек"


# ============================================================================
# УТИЛИТЫ
# ============================================================================


class ProgressTracker:
    """
    Трекер прогресса с расчётом метрик производительности.

    Вычисляет:
    - Процент выполнения
    - Оставшееся время (ETA)
    - Скорость обработки (строк/сек)
    - Прошедшее время
    """

    def __init__(self, total_rows: int = 0):
        """
        Инициализация трекера.

        Args:
            total_rows: Общее количество строк для обработки
        """
        self.total_rows = total_rows
        self.processed_rows = 0
        self.start_time: Optional[float] = None
        self.last_update_time: float = 0
        self.min_update_interval: float = 0.1  # Минимальный интервал обновления (сек)

    def start(self):
        """Запускает отсчёт времени."""
        self.start_time = time.time()
        self.last_update_time = self.start_time

    def update(self, processed_rows: int, force: bool = False) -> ProgressData:
        """
        Обновляет прогресс и возвращает метрики.

        Args:
            processed_rows: Количество обработанных строк

        Returns:
            ProgressData с актуальными метриками
        """
        current_time = time.time()

        # Защита от слишком частых обновлений
        if (
            not force
            and current_time - self.last_update_time < self.min_update_interval
        ):
            return None

        self.processed_rows = processed_rows
        self.last_update_time = current_time

        # Расчёт метрик
        elapsed = current_time - self.start_time if self.start_time else 0

        # Скорость обработки
        speed = processed_rows / elapsed if elapsed > 0 else 0

        # Процент выполнения
        percent = (
            int((processed_rows / max(self.total_rows, 1)) * 100)
            if self.total_rows > 0
            else 0
        )

        # ETA (оставшееся время)
        if speed > 0 and self.total_rows > 0:
            remaining_rows = self.total_rows - processed_rows
            eta = remaining_rows / speed
        else:
            eta = 0

        return ProgressData(
            percent=min(percent, 100),
            processed_rows=processed_rows,
            total_rows=self.total_rows,
            elapsed_seconds=elapsed,
            eta_seconds=eta,
            rows_per_second=speed,
        )

    def reset(self, new_total: int = 0):
        """
        Сбрасывает трекер для новой задачи.

        Args:
            new_total: Новое общее количество строк (0 = не менять)
        """
        self.processed_rows = 0
        self.start_time = None
        self.last_update_time = 0
        if new_total > 0:
            self.total_rows = new_total


class FileUtilities:
    """
    Утилиты для работы с файлами.
    Все методы статические - класс не требует инициализации.
    """

    @staticmethod
    def get_encoding(file_path: str) -> str:
        """
        Определяет кодировку файла по BOM или содержимому.

        Args:
            file_path: Путь к файлу

        Returns:
            Название кодировки ('utf-8', 'utf-8-sig', 'utf-16', 'windows-1251')
        """
        try:
            with open(file_path, "rb") as f:
                raw_data = f.read(4)

                # Проверка BOM
                if raw_data.startswith(b"\xff\xfe\x00\x00"):
                    return "utf-32-le"
                elif raw_data.startswith(b"\x00\x00\xfe\xff"):
                    return "utf-32-be"
                elif raw_data.startswith(b"\xff\xfe"):
                    return "utf-16-le"
                elif raw_data.startswith(b"\xfe\xff"):
                    return "utf-16-be"
                elif raw_data.startswith(b"\xef\xbb\xbf"):
                    return "utf-8-sig"
                else:
                    # Пробуем UTF-8 first, затем windows-1251
                    try:
                        with open(file_path, "r", encoding="utf-8") as f:
                            f.read(8192)
                        return "utf-8"
                    except UnicodeDecodeError:
                        return "windows-1251"
        except Exception:
            return "utf-8"

    @staticmethod
    def get_delimiter(file_path: str) -> str:
        """
        Определяет разделитель файла.

        Args:
            file_path: Путь к файлу

        Returns:
            Разделитель ('\t', ';', ',')
        """
        try:
            encoding = FileUtilities.get_encoding(file_path)
            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                first_line = f.readline(4096)  # Читаем первые 4KB

                # Удаляем BOM
                if first_line.startswith("\ufeff"):
                    first_line = first_line[1:]

                # Считаем разделители (упрощённая эвристика)
                delimiters = {
                    "\t": first_line.count("\t"),
                    ";": first_line.count(";"),
                    ",": first_line.count(","),
                }

                # Выбираем максимальный
                best_delim = max(delimiters.items(), key=lambda x: x[1])

                if best_delim[1] > 0:
                    return best_delim[0]
                else:
                    # По умолчанию
                    return "," if file_path.lower().endswith(".csv") else "\t"

        except Exception:
            return "\t"

    @staticmethod
    def sanitize_sheet_name(name: str, used_names: Optional[Set[str]] = None) -> str:
        """
        Очищает имя листа для Excel.

        Args:
            name: Исходное имя
            used_names: Множество уже использованных имён

        Returns:
            Безопасное имя листа
        """
        if used_names is None:
            used_names = set()

        name = name.strip()
        name = name.strip("'").strip()

        # Удаляем запрещённые символы
        invalid_chars = [":", "\\", "/", "?", "*", "[", "]"]
        for char in invalid_chars:
            name = name.replace(char, "_")

        # Обрезаем
        name = name.strip()[:29]

        if not name:
            name = "Sheet"

        # Уникальность
        base_name = name
        counter = 1
        while name.lower() in {n.lower() for n in used_names}:
            suffix = f"_{counter}"
            name = f"{base_name[: 29 - len(suffix)]}{suffix}"
            counter += 1

        used_names.add(name)
        return name

    @staticmethod
    def sanitize_file_stem(name: str, used_names: Optional[Set[str]] = None) -> str:
        """Очищает часть имени файла и гарантирует уникальность."""
        if used_names is None:
            used_names = set()

        invalid_chars = '<>:"/\\|?*'
        cleaned_name = "".join(
            "_" if char in invalid_chars or ord(char) < 32 else char for char in name
        )
        cleaned_name = " ".join(cleaned_name.split())
        cleaned_name = cleaned_name.strip().rstrip(". ")
        cleaned_name = cleaned_name[:80].rstrip(". ")

        if not cleaned_name:
            cleaned_name = "Unknown"

        base_name = cleaned_name
        counter = 1
        used_names_lower = {item.lower() for item in used_names}

        while cleaned_name.lower() in used_names_lower:
            suffix = f"_{counter}"
            trimmed_base = base_name[: 80 - len(suffix)].rstrip(". ")
            cleaned_name = (
                f"{trimmed_base}{suffix}" if trimmed_base else f"File{suffix}"
            )
            counter += 1

        used_names.add(cleaned_name)
        return cleaned_name

    @staticmethod
    def count_rows(
        file_path: str,
        delimiter: str,
        encoding: str,
        filter_column_idx: Optional[int] = None,
        filter_values: Optional[Set[str]] = None,
    ) -> int:
        """
        Подсчитывает количество строк в файле с учётом фильтра.

        Args:
            file_path: Путь к файлу
            delimiter: Разделитель
            encoding: Кодировка
            filter_column_idx: Индекс столбца для фильтра (-1 если нет)
            filter_values: Значения для фильтра

        Returns:
            Количество строк или -1 при ошибке
        """
        count = 0
        try:
            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                reader = csv.reader(f, delimiter=delimiter)
                next(reader, None)  # Пропускаем заголовок

                for row in reader:
                    if filter_column_idx is not None and filter_values:
                        if filter_column_idx < len(row):
                            if row[filter_column_idx] not in filter_values:
                                continue
                    count += 1
        except (UnicodeDecodeError, OSError, IOError, csv.Error):
            return -1

        return count


# ============================================================================
# КОНВЕРТЕР
# ============================================================================


class TSVToExcelConverter(QThread):
    """
    Поток конвертации TSV/CSV в Excel.

    Сигналы:
        update_progress: int - процент выполнения (для совместимости)
        progress_data: ProgressData - расширенные данные прогресса
        log_message: (str, QColor) - сообщение лога
        finished: () - завершение работы
        error: str - ошибка
    """

    update_progress = Signal(int)
    progress_data = Signal(object)  # ProgressData
    log_message = Signal(str, QColor)
    finished_signal = Signal()
    stopped_signal = Signal()
    error = Signal(str)

    @staticmethod
    def _build_output_base_name(input_file: str, file_index: int = 0) -> str:
        """Формирует базовое имя выходного файла от исходного имени и текущего времени."""
        source_name = os.path.splitext(os.path.basename(input_file))[0]
        if file_index > 0:
            source_name = f"{source_name}_{file_index + 1}"

        timestamp = datetime.now().strftime("%H.%M.%S")
        return f"{source_name} (conver_{timestamp})"

    # Архитектурные константы
    MAX_EXCEL_ROWS = 1000000  # Ограничение в 1млн строк на один лист Excel
    STOP_CHECK_INTERVAL = (
        2000  # Как часто (в строках) проверять флаг остановки конвертации
    )

    def __init__(
        self,
        input_files: List[str],
        output_directory: str,
        output_format: str = "xlsx",
        auto_open: bool = False,
        auto_delete: bool = False,
        styles: Optional[Dict[str, Any]] = None,
        header_color: str = "#C8DCF0",
        split_column: str = "",
        split_mode: str = "sheets",
        selected_values: Optional[List[str]] = None,
        filter_column: str = "",
        filter_values: Optional[List[str]] = None,
        pivot_settings: Optional[Dict[str, Any]] = None,
        selected_columns: Optional[List[str]] = None,
        deduplicate_rows: bool = True,
        ram_threshold: int = 500000,
    ):
        super().__init__()

        self.input_files = input_files
        self.output_directory = output_directory
        self.output_format = output_format
        self.auto_open = auto_open
        self.auto_delete = auto_delete
        self.styles = styles or {}
        self.header_color = header_color
        self.split_column = split_column
        self.split_mode = split_mode
        self.selected_values = set(selected_values) if selected_values else set()
        self.filter_column = filter_column
        self.filter_values = set(filter_values) if filter_values else set()
        self.pivot_settings = pivot_settings
        self.selected_columns = selected_columns or []
        self.deduplicate_rows = deduplicate_rows
        self._timing = {}
        self._timing_total = 0.0
        self.ram_threshold = ram_threshold

        self.stop_flag = False
        self.output_file_path: Optional[str] = None
        self.generated_files: List[str] = []
        self.total_rows = 0
        self.processed_rows = 0
        self.duplicates_removed = 0
        self._deduplication_was_used = False

        # Трекер прогресса
        self.progress_tracker = ProgressTracker()

        # Кэш форматов
        self._cached_formats: Dict[str, Any] = {}

    def run(self):
        """Основной метод потока."""
        try:
            self.log_message.emit("Начало конвертации...", QColor("blue"))
            self.processed_rows = 0
            self._timing = {}
            self._timing_total = 0.0
            self.generated_files = []
            self.duplicates_removed = 0
            self._deduplication_was_used = False

            # Валидация
            if not self.input_files:
                raise ValueError("Нет файлов для конвертации")

            if not os.path.isdir(self.output_directory):
                raise ValueError(f"Директория не существует: {self.output_directory}")

            # Подсчёт общего количества строк
            t0 = time.time()
            self._count_total_rows()
            self._timing['count_rows'] = time.time() - t0

            if self.stop_flag:
                self._finish_as_stopped()
                return

            # Инициализация трекера прогресса
            self.progress_tracker.reset(self.total_rows)
            self.progress_tracker.start()

            # Конвертация каждого файла
            processed = 0
            for tsv_file in self.input_files:
                if self.stop_flag:
                    self._finish_as_stopped()
                    return

                # Обновляем информацию о текущем файле без сброса метрик прогресса
                self._emit_progress_update(
                    os.path.basename(tsv_file), "Чтение файла...", force=True
                )

                t_file_start = time.time()
                conversion_result = self._convert_file(tsv_file, processed)
                t_file_elapsed = time.time() - t_file_start
                self._timing[f'convert_file_{processed}'] = t_file_elapsed

                if conversion_result == "success":
                    processed += 1
                elif conversion_result == "stopped":
                    self._finish_as_stopped()
                    return
                else:
                    self.log_message.emit(
                        f"Ошибка обработки файла: {os.path.basename(tsv_file)}",
                        QColor("red"),
                    )

            # Финальный отчёт
            total_time = time.time() - t0
            self._timing_total = total_time
            self._log_timing_summary(total_time)
            if self._deduplication_was_used:
                self.log_message.emit(
                    f"Удалено дубликатов после исключения столбцов: {self.duplicates_removed}",
                    QColor("blue"),
                )

            self.log_message.emit(
                f"Конвертация завершена. Обработано файлов: {processed}",
                QColor("green"),
            )
            self.finished_signal.emit()

        except Exception as e:
            self.log_message.emit(f"Критическая ошибка: {str(e)}", QColor("red"))
            self.error.emit(str(e))

    def _get_split_value(
        self, row: List[str], split_idx: int, selected_vals: Set[str]
    ) -> str:
        """Извлекает значение для разделения и применяет логику 'Остальные'."""
        value = row[split_idx] if split_idx < len(row) else ""
        if not value or not value.strip():
            return ""

        if selected_vals and value not in selected_vals:
            return "Остальные"

        return value

    def _get_output_columns(self, headers: List[str]) -> tuple[List[str], List[int]]:
        """Возвращает заголовки и индексы столбцов для итогового файла."""
        if not self.selected_columns:
            return headers, list(range(len(headers)))

        header_to_index = {}
        for index, header in enumerate(headers):
            if header not in header_to_index:
                header_to_index[header] = index

        output_indices = []
        missing_columns = []
        for column in self.selected_columns:
            if column in header_to_index:
                output_indices.append(header_to_index[column])
            else:
                missing_columns.append(column)

        if missing_columns:
            self.log_message.emit(
                "Не найдены выбранные столбцы: " + ", ".join(missing_columns),
                QColor("orange"),
            )

        if not output_indices:
            self.log_message.emit(
                "Выбранные столбцы не найдены. Будут сохранены все столбцы.",
                QColor("orange"),
            )
            return headers, list(range(len(headers)))

        if len(output_indices) != len(headers) or output_indices != list(range(len(headers))):
            self.log_message.emit(
                f"Столбцы вывода: {len(output_indices)} из {len(headers)}",
                QColor("blue"),
            )

        return [headers[index] for index in output_indices], output_indices

    @staticmethod
    def _project_row(row: List[str], output_indices: List[int]) -> List[str]:
        """Оставляет в строке только выбранные столбцы в заданном порядке."""
        return [row[index] if index < len(row) else "" for index in output_indices]

    def _should_deduplicate_rows(
        self, headers: List[str], output_indices: List[int]
    ) -> bool:
        """Дедупликация включается только если пользователь исключил столбцы."""
        enabled = (
            self.deduplicate_rows
            and bool(self.selected_columns)
            and len(output_indices) < len(headers)
        )
        if enabled:
            self._deduplication_was_used = True
        return enabled

    def _finish_as_stopped(self):
        """Завершает поток в статусе пользовательской остановки."""
        self.output_file_path = None
        self.log_message.emit(
            "Конвертация остановлена пользователем. Результат может быть неполным.",
            QColor("orange"),
        )
        self.stopped_signal.emit()

    def _emit_progress_update(
        self, current_file: str, current_operation: str, force: bool = False
    ):
        """Отправляет обновление детального прогресса конвертации."""
        progress_data = self.progress_tracker.update(self.processed_rows, force=force)
        if not progress_data:
            return

        progress_data.current_file = current_file
        progress_data.current_operation = current_operation
        self.progress_data.emit(progress_data)
        self.update_progress.emit(progress_data.percent)

    def _log_timing_summary(self, total_time: float):
        """Логирует сводку времени конвертации."""
        self.log_message.emit("=== Профилирование ===", QColor("cyan"))

        # Основные этапы
        if 'count_rows' in self._timing:
            self.log_message.emit(f"Подсчёт строк: {self._timing['count_rows']:.2f}s", QColor("gray"))

        # Детали файловых конверсий
        file_times = []
        for k, v in sorted(self._timing.items()):
            if k.startswith('convert_file_'):
                file_idx = int(k.split('_')[-1]) if '_' in k else 0
                file_times.append((file_idx, k, v))

        for file_idx, k, v in sorted(file_times):
            pct = (v / total_time * 100) if total_time > 0 else 0
            self.log_message.emit(f"  Файл #{file_idx + 1}: {v:.2f}s ({pct:.1f}%)", QColor("gray"))

        # Детальные метрики по всем файлам
        detail_keys = ['create_workbook', 'close_workbook', 'Запись данных...']
        for k in detail_keys:
            if k in self._timing:
                v = self._timing[k]
                pct = (v / total_time * 100) if total_time > 0 else 0
                self.log_message.emit(f"  {k}: {v:.2f}s ({pct:.1f}%)", QColor("gray"))

        # Суммарное время закрытия всех файлов
        if 'close_workbook_total' in self._timing:
            v = self._timing['close_workbook_total']
            pct = (v / total_time * 100) if total_time > 0 else 0
            self.log_message.emit(f"  Суммарное закрытие файлов: {v:.2f}s ({pct:.1f}%)", QColor("red"))

        if total_time > 0:
            self.log_message.emit(f"Общее время: {total_time:.2f}s", QColor("blue"))

    def _count_total_rows(self):
        """Подсчитывает общее количество строк для прогресс-бара."""
        self.total_rows = 0

        for tsv_file in self.input_files:
            if self.stop_flag:
                break

            try:
                encoding = FileUtilities.get_encoding(tsv_file)
                delimiter = FileUtilities.get_delimiter(tsv_file)

                with open(tsv_file, "r", encoding=encoding, errors="replace") as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    headers = next(reader)

                    # Индекс фильтра
                    filter_idx = None
                    if self.filter_column and self.filter_column != "Не фильтровать":
                        try:
                            filter_idx = headers.index(self.filter_column)
                        except ValueError:
                            filter_idx = None

                    # Индекс разделения
                    split_idx = None
                    if self.split_column and self.split_column != "Не разделять":
                        try:
                            split_idx = headers.index(self.split_column)
                        except ValueError:
                            split_idx = None

                    output_indices = None
                    if self.selected_columns and len(self.selected_columns) < len(headers):
                        header_to_index = {}
                        for index, header in enumerate(headers):
                            if header not in header_to_index:
                                header_to_index[header] = index
                        output_indices = [
                            header_to_index[column]
                            for column in self.selected_columns
                            if column in header_to_index
                        ]

                    deduplicate_rows = (
                        self.deduplicate_rows
                        and bool(self.selected_columns)
                        and bool(output_indices)
                        and len(output_indices) < len(headers)
                    )
                    seen_rows = set()
                    seen_rows_by_split = defaultdict(set)

                    def should_count_row(row: List[str]) -> bool:
                        if split_idx is not None:
                            split_value = self._get_split_value(
                                row, split_idx, self.selected_values
                            )
                            if not split_value:
                                return False
                        else:
                            split_value = None

                        if deduplicate_rows:
                            row_key = tuple(
                                row[index] if index < len(row) else ""
                                for index in output_indices
                            )
                            if split_value is not None:
                                if row_key in seen_rows_by_split[split_value]:
                                    return False
                                seen_rows_by_split[split_value].add(row_key)
                            else:
                                if row_key in seen_rows:
                                    return False
                                seen_rows.add(row_key)

                        return True

                    # Считаем строки (оптимизированные циклы)
                    if filter_idx is not None and self.filter_values:
                        filter_vals = self.filter_values  # Локальная ссылка быстрее
                        for i, row in enumerate(reader):
                            if i % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                                return
                            if (
                                filter_idx < len(row)
                                and row[filter_idx] in filter_vals
                                and should_count_row(row)
                            ):
                                self.total_rows += 1
                    else:
                        for i, row in enumerate(reader):
                            if i % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                                return
                            if should_count_row(row):
                                self.total_rows += 1

            except PermissionError:
                self.log_message.emit(
                    f"Ошибка доступа к файлу {os.path.basename(tsv_file)}. Возможно, он открыт в другой программе (Excel?). Закройте файл и попробуйте снова.",
                    QColor("red"),
                )
            except (OSError, IOError, UnicodeDecodeError, csv.Error) as e:
                self.log_message.emit(
                    f"Ошибка подсчёта строк {os.path.basename(tsv_file)}: {e}",
                    QColor("red"),
                )

        self.log_message.emit(
            f"Всего строк для обработки: {self.total_rows}", QColor("blue")
        )

    def _convert_file(self, input_file: str, processed_files: int) -> str:
        """
        Конвертирует один файл.

        Args:
            input_file: Путь к входному файлу
            processed_files: Количество уже обработанных файлов (для уникальности имени)

        Returns:
            `success`, `stopped` или `error`
        """
        start_time = time.time()

        try:
            # Определяем параметры файла
            encoding = FileUtilities.get_encoding(input_file)
            delimiter = FileUtilities.get_delimiter(input_file)

            self.log_message.emit(
                f"Обработка {os.path.basename(input_file)} (кодировка: {encoding})",
                QColor("blue"),
            )

            # Выбор стратегии в зависимости от формата
            if self.output_format.lower() == "csv":
                return self._convert_to_csv(
                    input_file, encoding, delimiter, start_time, processed_files
                )
            else:
                return self._convert_to_xlsx(
                    input_file, encoding, delimiter, start_time, processed_files
                )

        except PermissionError:
            self.log_message.emit(
                f"Ошибка доступа к файлу {os.path.basename(input_file)}. Возможно, он открыт в другой программе (Excel?). Закройте файл и попробуйте снова.",
                QColor("red"),
            )
            return "error"
        except (OSError, IOError, UnicodeDecodeError) as e:
            self.log_message.emit(
                f"Ошибка конвертации {os.path.basename(input_file)}: {str(e)}",
                QColor("red"),
            )
            return "error"

    def _convert_to_csv(
        self,
        input_file: str,
        encoding: str,
        delimiter: str,
        start_time: float,
        file_index: int = 0,
    ) -> str:
        """Конвертация в CSV."""
        base_name = self._build_output_base_name(input_file, file_index)

        with open(input_file, "r", encoding=encoding, errors="replace") as f:
            reader = csv.reader(f, delimiter=delimiter)

            try:
                headers = next(reader)
            except StopIteration:
                self.log_message.emit("Пустой файл", QColor("red"))
                return "error"

            # Индексы
            split_idx = None
            filter_idx = None

            if self.split_column and self.split_column != "Не разделять":
                try:
                    split_idx = headers.index(self.split_column)
                except ValueError:
                    split_idx = None

            if self.filter_column and self.filter_column != "Не фильтровать":
                try:
                    filter_idx = headers.index(self.filter_column)
                except ValueError:
                    filter_idx = None

            # Если есть разделение, пишем в разные файлы
            if split_idx is not None:
                self._write_split_csv(reader, headers, base_name, split_idx, filter_idx)
            else:
                self._write_single_csv(reader, headers, base_name, filter_idx)

        if self.stop_flag:
            self.output_file_path = None
            return "stopped"

        # Сводная таблица для CSV (отдельный файл)
        if self.pivot_settings:
            self._create_csv_pivot(input_file, base_name)

        elapsed = time.time() - start_time
        minutes = int(elapsed // 60)
        seconds = int(elapsed % 60)
        self.log_message.emit(
            f"Конвертация в CSV завершена ({minutes:02d}:{seconds:02d})",
            QColor("green"),
        )
        return "success"

    def _write_single_csv(self, reader, headers, base_name, filter_idx):
        output_path = os.path.join(self.output_directory, f"{base_name}.csv")
        current_file = os.path.basename(output_path)
        output_headers, output_indices = self._get_output_columns(headers)
        deduplicate_rows = self._should_deduplicate_rows(headers, output_indices)
        seen_rows = set()

        with open(output_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(output_headers)

            def process_row(row):
                output_row = self._project_row(row, output_indices)
                if deduplicate_rows:
                    row_key = tuple(output_row)
                    if row_key in seen_rows:
                        self.duplicates_removed += 1
                        return False
                    seen_rows.add(row_key)

                writer.writerow(output_row)
                return True

            self._process_rows_with_progress(
                reader, filter_idx, current_file, process_row, "Запись CSV..."
            )

        self.output_file_path = output_path
        self.generated_files.append(output_path)

    def _write_split_csv(self, reader, headers, base_name, split_idx, filter_idx):
        MAX_OPEN_FILES = 200
        open_files = {}
        writers = {}
        used_file_names: Set[str] = set()
        file_paths: Dict[str, str] = {}
        row_counts: Dict[str, int] = {}
        selected_vals = self.selected_values
        current_file = f"{base_name}_*.csv"
        output_headers, output_indices = self._get_output_columns(headers)
        deduplicate_rows = self._should_deduplicate_rows(headers, output_indices)
        seen_rows_by_key = defaultdict(set)

        def _create_csv_file(key: str):
            """Создаёт CSV файл и записывает заголовок."""
            safe_key = FileUtilities.sanitize_file_stem(key, used_file_names)
            file_path = os.path.join(
                self.output_directory, f"{base_name}_{safe_key}.csv"
            )
            file_paths[key] = file_path

            f = open(file_path, "w", encoding="utf-8-sig", newline="")
            writer = csv.writer(f, delimiter=";")
            writer.writerow(output_headers)

            open_files[key] = f
            writers[key] = writer
            row_counts[key] = 0

            return file_path

        def process_row(row):
            key = self._get_split_value(row, split_idx, selected_vals)
            if not key:
                return False

            output_row = self._project_row(row, output_indices)
            if deduplicate_rows:
                row_key = tuple(output_row)
                if row_key in seen_rows_by_key[key]:
                    self.duplicates_removed += 1
                    return False
                seen_rows_by_key[key].add(row_key)

            if key not in writers:
                if len(open_files) >= MAX_OPEN_FILES:
                    oldest_key = next(iter(open_files))
                    open_files[oldest_key].close()
                    del open_files[oldest_key]
                    del writers[oldest_key]

                _create_csv_file(key)

            writers[key].writerow(output_row)
            row_counts[key] = row_counts.get(key, 0) + 1
            return True

        try:
            self._process_rows_with_progress(
                reader, filter_idx, current_file, process_row, "Распределение по CSV..."
            )
        finally:
            for f in open_files.values():
                f.close()

        files_to_remove = []
        for key, count in row_counts.items():
            if count == 0 and key in file_paths:
                files_to_remove.append(file_paths[key])

        for file_path in files_to_remove:
            try:
                os.remove(file_path)
            except Exception:
                pass

        if files_to_remove:
            self.log_message.emit(
                f"Удалено пустых CSV файлов (без данных): {len(files_to_remove)}",
                QColor("blue"),
            )

        self.output_file_path = self.output_directory
        for path in file_paths.values():
            if path not in files_to_remove:
                self.generated_files.append(path)

    def _create_csv_pivot(self, input_file, base_name):
        try:
            processor = PivotTableProcessor(lambda _msg, _color: None)
            pivot_data = processor.create_pivot_data(
                input_file,
                self.pivot_settings,
                self.filter_column if self.filter_column != "Не фильтровать" else "",
                self.filter_values,
            )

            if pivot_data:
                output_path = os.path.join(
                    self.output_directory, f"{base_name}_Pivot.csv"
                )
                with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f, delimiter=";")

                    # Заголовки
                    row_headers = self.pivot_settings.get("rows", [])
                    writer.writerow(row_headers + ["Колонка", "Метрика", "Значение"])

                    for row_key, col_dict in pivot_data.items():
                        row_prefix = (
                            list(row_key) if isinstance(row_key, tuple) else [row_key]
                        )
                        for col_key, val_dict in col_dict.items():
                            col_prefix = (
                                list(col_key)
                                if isinstance(col_key, tuple)
                                else [col_key]
                            )
                            col_str = " / ".join(map(str, col_prefix))

                            for metric, value in val_dict.items():
                                writer.writerow(row_prefix + [col_str, metric, value])

                self.log_message.emit(
                    "Сводная таблица сохранена в отдельный CSV", QColor("green")
                )
        except (OSError, IOError, PermissionError) as e:
            self.log_message.emit(f"Ошибка CSV сводной: {e}", QColor("red"))

    def _convert_to_xlsx(
        self,
        input_file: str,
        encoding: str,
        delimiter: str,
        start_time: float,
        file_index: int = 0,
    ) -> str:
        """Стандартная конвертация в XLSX (перенесенная логика)."""
        # Имя выходного файла
        base_name = self._build_output_base_name(input_file, file_index)
        output_path = os.path.join(self.output_directory, f"{base_name}.xlsx")

        # Ожидаемый режим разделения на файлы
        expected_split_to_files = False
        if self.split_column and self.split_column != "Не разделять":
            if self.split_mode == "files":
                expected_split_to_files = True

        workbook = None
        result = "error"
        actual_split_to_files = expected_split_to_files

        # Определяем режим работы xlsxwriter:
        # - < ram_threshold: быстрый режим (временные файлы, баланс скорости и памяти)
        # - >= ram_threshold: constant_memory (минимум RAM, последовательная запись)
        # НЕ используем in_memory: True из-за долгого сжатия в конце
        use_constant_memory = self.total_rows >= self.ram_threshold
        try:
            t_convert = time.time()

            with open(input_file, "r", encoding=encoding, errors="replace") as f:
                reader = csv.reader(f, delimiter=delimiter)

                try:
                    headers = next(reader)
                except StopIteration:
                    self.log_message.emit("Пустой файл", QColor("red"))
                    return "error"

                split_idx = None
                filter_idx = None

                if self.split_column and self.split_column != "Не разделять":
                    try:
                        split_idx = headers.index(self.split_column)
                    except ValueError:
                        split_idx = None
                        self.log_message.emit(
                            f"Столбец для разделения '{self.split_column}' не найден в файле. Разделение отменено.", 
                            QColor("orange")
                        )

                if self.filter_column and self.filter_column != "Не фильтровать":
                    try:
                        filter_idx = headers.index(self.filter_column)
                    except ValueError:
                        filter_idx = None

                # Корректируем флаг фактического разделения на файлы
                if expected_split_to_files and split_idx is None:
                    actual_split_to_files = False

                # Создаём главный workbook, если мы НЕ разделяем на файлы (или если разделение отменилось)
                if not actual_split_to_files:
                    t_create_wb = time.time()
                    if use_constant_memory:
                        workbook = xlsxwriter.Workbook(
                            output_path,
                            {
                                "constant_memory": True,
                                "use_zip64": True,
                            },
                        )
                        self.log_message.emit(
                            f"Режим: экономия памяти (Constant Memory, {self.total_rows:,} строк)",
                            QColor("blue"),
                        )
                    else:
                        workbook = xlsxwriter.Workbook(output_path)
                        self.log_message.emit(
                            f"Режим: быстрый (временные файлы, {self.total_rows:,} строк)",
                            QColor("blue"),
                        )
                    self._timing['create_workbook'] = self._timing.get('create_workbook', 0.0) + (time.time() - t_create_wb)
                    self._init_formats(workbook)

                if split_idx is not None:
                    if self.split_mode == "files":
                        self._convert_with_split_to_files(
                            reader,
                            headers,
                            split_idx,
                            filter_idx,
                            base_name,
                            os.path.basename(input_file),
                        )
                    else:
                        self._convert_with_split(
                            workbook,
                            reader,
                            headers,
                            split_idx,
                            filter_idx,
                            os.path.basename(input_file),
                        )
                else:
                    self._convert_without_split(
                        workbook,
                        reader,
                        headers,
                        filter_idx,
                        os.path.basename(input_file),
                    )

            if self.stop_flag:
                result = "stopped"
            else:
                # Обработка сводной таблицы
                if self.pivot_settings:
                    try:
                        self.log_message.emit(
                            "Создание сводной таблицы...", QColor("blue")
                        )

                        # Читаем данные из файла
                        processor = PivotTableProcessor(lambda _msg, _color: None)
                        pivot_data = processor.create_pivot_data(
                            input_file,
                            self.pivot_settings,
                            self.filter_column
                            if self.filter_column != "Не фильтровать"
                            else "",
                            self.filter_values,
                        )

                        if pivot_data:
                            if self.split_mode == "files" and split_idx is not None:
                                # При разделении на файлы — сводная в отдельный файл
                                pivot_output_path = os.path.join(
                                    self.output_directory, f"{base_name}_Pivot.xlsx"
                                )
                                pivot_workbook = xlsxwriter.Workbook(pivot_output_path)

                                header_format = pivot_workbook.add_format(
                                    {
                                        "bold": self.styles.get("bold", False),
                                        "italic": self.styles.get("italic", False),
                                        "font_size": self.styles.get("font_size", 12),
                                        "font_name": self.styles.get(
                                            "font_name", "Arial"
                                        ),
                                        "bg_color": self.header_color,
                                        "align": "center",
                                        "valign": "vcenter",
                                    }
                                )
                                if self.styles.get("border", 0) == 1:
                                    header_format.set_border(1)

                                cached_formats = {"header": header_format}

                                processor.write_pivot_table(
                                    pivot_workbook,
                                    pivot_data,
                                    self.pivot_settings,
                                    cached_formats,
                                )
                                pivot_workbook.close()

                                self.log_message.emit(
                                    f"Сводная таблица: {os.path.basename(pivot_output_path)}",
                                    QColor("green"),
                                )
                                self.generated_files.append(pivot_output_path)
                            else:
                                # Сводная в текущий workbook
                                header_format = workbook.add_format(
                                    {
                                        "bold": self.styles.get("bold", False),
                                        "italic": self.styles.get("italic", False),
                                        "font_size": self.styles.get("font_size", 12),
                                        "font_name": self.styles.get(
                                            "font_name", "Arial"
                                        ),
                                        "bg_color": self.header_color,
                                        "align": "center",
                                        "valign": "vcenter",
                                    }
                                )
                                if self.styles.get("border", 0) == 1:
                                    header_format.set_border(1)

                                cached_formats = {"header": header_format}

                                processor.write_pivot_table(
                                    workbook,
                                    pivot_data,
                                    self.pivot_settings,
                                    cached_formats,
                                )

                                self.log_message.emit(
                                    "Сводная таблица создана", QColor("green")
                                )
                        else:
                            self.log_message.emit(
                                "Нет данных для сводной таблицы", QColor("orange")
                            )

                    except Exception as e:
                        self.log_message.emit(
                            f"Ошибка сводной таблицы: {str(e)}", QColor("red")
                        )

                result = "success"

        except Exception as e:
            self.log_message.emit(f"Ошибка конвертации: {str(e)}", QColor("red"))
            result = "error"

        finally:
            if workbook is not None:
                self._emit_progress_update(
                    os.path.basename(input_file), "Сохранение файла...", force=True
                )
                t_close = time.time()
                try:
                    workbook.close()
                except Exception:
                    pass
                cw_time = time.time() - t_close
                self._timing['close_workbook'] = cw_time
                self._timing['close_workbook_total'] = self._timing.get('close_workbook_total', 0) + cw_time

            self._cached_formats = {}

            # Record overall conversion time
            self._timing['conversion_total'] = time.time() - t_convert

        if (
                result != "success"
                and not actual_split_to_files
                and os.path.exists(output_path)
            ):
                try:
                    os.remove(output_path)
                except Exception:
                    pass

        if result == "success":
            if not actual_split_to_files:
                self.output_file_path = output_path
                self.generated_files.append(output_path)

                elapsed = time.time() - start_time
                minutes = int(elapsed // 60)
                seconds = int(elapsed % 60)

                self.log_message.emit(
                    f"Файл сохранён: {os.path.basename(output_path)} ({minutes:02d}:{seconds:02d})",
                    QColor("green"),
                )
            else:
                self.log_message.emit(
                    "Разделение на файлы завершено",
                    QColor("green"),
                )

        return result

    def _init_formats(self, workbook: xlsxwriter.Workbook):
        """Инициализирует кэш форматов."""
        self._cached_formats = {}

        # Формат заголовка
        header_format = workbook.add_format(
            {
                "bold": self.styles.get("bold", False),
                "italic": self.styles.get("italic", False),
                "font_size": self.styles.get("font_size", 12),
                "font_name": self.styles.get("font_name", "Arial"),
                "bg_color": self.header_color,
                "align": "center",
                "valign": "vcenter",
            }
        )

        if self.styles.get("border", 0) == 1:
            header_format.set_border(1)

        self._cached_formats["header"] = header_format

        # Формат ячейки
        cell_format = workbook.add_format(
            {
                "font_size": self.styles.get("font_size", 12),
                "font_name": self.styles.get("font_name", "Arial"),
            }
        )

        if self.styles.get("border", 0) == 1:
            cell_format.set_border(1)

        self._cached_formats["cell"] = cell_format

    def _process_rows_with_progress(
        self,
        reader,
        filter_idx: Optional[int],
        current_file: str,
        row_handler,
        operation_name: str = "Запись данных...",
    ):
        """Обрабатывает строки с фильтрацией, остановкой и обновлением прогресса."""
        t_process = time.time()
        update_freq = max(1, self.total_rows // 100)

        if filter_idx is not None and self.filter_values:
            filter_vals = self.filter_values
            for row_num, row in enumerate(reader, 1):
                if row_num % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                    break

                if filter_idx < len(row) and row[filter_idx] not in filter_vals:
                    continue

                if row_handler(row) is False:
                    continue
                self.processed_rows += 1
                if self.processed_rows % update_freq == 0:
                    self._emit_progress_update(current_file, operation_name)
        else:
            for row_num, row in enumerate(reader, 1):
                if row_num % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                    break

                if row_handler(row) is False:
                    continue
                self.processed_rows += 1
                if self.processed_rows % update_freq == 0:
                    self._emit_progress_update(current_file, operation_name)

        self._emit_progress_update(current_file, operation_name, force=True)
        self._timing[operation_name] = time.time() - t_process

    def _convert_with_split(
        self,
        workbook: xlsxwriter.Workbook,
        reader: csv.reader,
        headers: List[str],
        split_idx: int,
        filter_idx: Optional[int],
        current_file: str,
    ):
        """Конвертация с разделением по столбцу."""
        worksheets: Dict[str, Any] = {}
        sheet_row_counts: Dict[str, int] = {}

        used_names: Set[str] = set()
        MAX_ROWS = self.MAX_EXCEL_ROWS
        selected_vals = self.selected_values
        output_headers, output_indices = self._get_output_columns(headers)
        deduplicate_rows = self._should_deduplicate_rows(headers, output_indices)
        seen_rows_by_value = defaultdict(set)

        def _create_sheet_with_headers(sheet_name: str):
            ws = workbook.add_worksheet(sheet_name)
            for col, header in enumerate(output_headers):
                ws.write(0, col, header, self._cached_formats["header"])
            return ws, 1

        def process_row(row):
            value = self._get_split_value(row, split_idx, selected_vals)
            if not value:
                return False

            output_row = self._project_row(row, output_indices)
            if deduplicate_rows:
                row_key = tuple(output_row)
                if row_key in seen_rows_by_value[value]:
                    self.duplicates_removed += 1
                    return False
                seen_rows_by_value[value].add(row_key)

            if value not in worksheets:
                sheet_name = FileUtilities.sanitize_sheet_name(value, used_names)
                ws, row_count = _create_sheet_with_headers(sheet_name)
                worksheets[value] = ws
                sheet_row_counts[value] = row_count

            current_row = sheet_row_counts[value]
            if current_row >= MAX_ROWS:
                sheet_name = FileUtilities.sanitize_sheet_name(f"{value}_2", used_names)
                ws, current_row = _create_sheet_with_headers(sheet_name)
                worksheets[value] = ws
                sheet_row_counts[value] = current_row

            worksheets[value].write_row(
                current_row, 0, output_row, self._cached_formats["cell"]
            )
            sheet_row_counts[value] = current_row + 1
            return True

        self._process_rows_with_progress(reader, filter_idx, current_file, process_row)

    def _convert_with_split_to_files(
        self,
        reader: csv.reader,
        headers: List[str],
        split_idx: int,
        filter_idx: Optional[int],
        base_name: str,
        current_file: str,
    ):
        """
        Конвертация с разделением на отдельные XLSX файлы.
        Каждое выбранное значение → отдельный .xlsx файл с одним листом.
        """
        open_workbooks: Dict[str, xlsxwriter.Workbook] = {}
        open_worksheets: Dict[str, Any] = {}
        open_row_counts: Dict[str, int] = {}
        used_file_names: Set[str] = set()
        file_paths: Dict[str, str] = {}
        selected_vals = self.selected_values
        output_files: List[str] = []
        output_headers, output_indices = self._get_output_columns(headers)
        deduplicate_rows = self._should_deduplicate_rows(headers, output_indices)
        seen_rows_by_value = defaultdict(set)

        use_constant_memory = self.total_rows >= self.ram_threshold

        if use_constant_memory:
            self.log_message.emit(
                f"Режим разделения: экономия памяти (Constant Memory, {self.total_rows:,} строк)",
                QColor("blue"),
            )
        else:
            self.log_message.emit(
                f"Режим разделения: быстрый (временные файлы, {self.total_rows:,} строк)",
                QColor("blue"),
            )

        self._timing['create_workbook'] = 0.0

        def _create_workbook_for_value(value: str):
            """Создаёт новый workbook для значения."""
            t_create = time.time()
            safe_value = FileUtilities.sanitize_file_stem(value, used_file_names)
            sheet_name = FileUtilities.sanitize_sheet_name(safe_value, used_file_names)
            file_path = os.path.join(
                self.output_directory, f"{base_name}_{safe_value}.xlsx"
            )
            file_paths[value] = file_path

            if use_constant_memory:
                workbook = xlsxwriter.Workbook(
                    file_path, {"constant_memory": True, "use_zip64": True}
                )
            else:
                workbook = xlsxwriter.Workbook(file_path)

            header_format = workbook.add_format(
                {
                    "bold": self.styles.get("bold", False),
                    "italic": self.styles.get("italic", False),
                    "font_size": self.styles.get("font_size", 12),
                    "font_name": self.styles.get("font_name", "Arial"),
                    "bg_color": self.header_color,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
            if self.styles.get("border", 0) == 1:
                header_format.set_border(1)

            cell_format = workbook.add_format(
                {
                    "font_size": self.styles.get("font_size", 12),
                    "font_name": self.styles.get("font_name", "Arial"),
                }
            )
            if self.styles.get("border", 0) == 1:
                cell_format.set_border(1)

            worksheet = workbook.add_worksheet(sheet_name)

            for col, header in enumerate(output_headers):
                worksheet.write(0, col, header, header_format)

            open_workbooks[value] = workbook
            open_worksheets[value] = worksheet
            open_row_counts[value] = 1
            workbook._tsv_cell_format = cell_format

            output_files.append(file_path)
            self._timing['create_workbook'] += time.time() - t_create

        def process_row(row):
            value = self._get_split_value(row, split_idx, selected_vals)
            if not value:
                return False

            output_row = self._project_row(row, output_indices)
            if deduplicate_rows:
                row_key = tuple(output_row)
                if row_key in seen_rows_by_value[value]:
                    self.duplicates_removed += 1
                    return False
                seen_rows_by_value[value].add(row_key)

            if value not in open_worksheets:
                _create_workbook_for_value(value)

            if value not in open_worksheets or value not in open_row_counts:
                return False

            worksheet = open_worksheets[value]
            row_count = open_row_counts[value]
            workbook = open_workbooks[value]
            worksheet.write_row(row_count, 0, output_row, workbook._tsv_cell_format)
            open_row_counts[value] = row_count + 1
            return True

        self._process_rows_with_progress(reader, filter_idx, current_file, process_row)

        self._emit_progress_update(current_file, "Сохранение файлов...", force=True)
        t_close_start = time.time()
        # Закрываем все workbook
        for value, wb in open_workbooks.items():
            try:
                wb.close()
            except Exception:
                pass

        t_close = time.time() - t_close_start
        self._timing['close_workbook'] = t_close
        self._timing['close_workbook_total'] = self._timing.get('close_workbook_total', 0) + t_close

        # Удаляем файлы без данных (только заголовок)
        # row_count = 1 означает только заголовок, без строк данных
        files_to_remove = []
        for value, row_count in open_row_counts.items():
            if row_count == 1:
                if value in file_paths:
                    files_to_remove.append(file_paths[value])

        for file_path in files_to_remove:
            try:
                os.remove(file_path)
                if file_path in output_files:
                    output_files.remove(file_path)
            except Exception:
                pass

        if files_to_remove:
            self.log_message.emit(
                f"Удалено пустых файлов (без данных): {len(files_to_remove)}",
                QColor("blue"),
            )

        if len(open_workbooks) > 100:
            self.log_message.emit(
                f"Создано {len(open_workbooks)} файлов. При большом количестве "
                f"уникальных значений возможно превышение лимита открытых файлов ОС.",
                QColor("orange"),
            )

        self.output_file_path = self.output_directory
        self.generated_files.extend(output_files)

    def _convert_without_split(
        self,
        workbook: xlsxwriter.Workbook,
        reader: csv.reader,
        headers: List[str],
        filter_idx: Optional[int],
        current_file: str,
    ):
        """Конвертация без разделения."""
        worksheet = None
        row_count = 0
        sheet_num = 1

        used_names: Set[str] = set()
        MAX_ROWS = self.MAX_EXCEL_ROWS
        output_headers, output_indices = self._get_output_columns(headers)
        deduplicate_rows = self._should_deduplicate_rows(headers, output_indices)
        seen_rows = set()

        # Вспомогательная функция для обработки
        def process_row(row):
            nonlocal worksheet, row_count, sheet_num
            output_row = self._project_row(row, output_indices)
            if deduplicate_rows:
                row_key = tuple(output_row)
                if row_key in seen_rows:
                    self.duplicates_removed += 1
                    return False
                seen_rows.add(row_key)

            # Новый лист если нужно
            if worksheet is None or row_count >= MAX_ROWS:
                sheet_name = FileUtilities.sanitize_sheet_name(
                    "Все проекты" if sheet_num == 1 else f"Все проекты_{sheet_num}",
                    used_names,
                )
                worksheet = workbook.add_worksheet(sheet_name)
                row_count = 0
                sheet_num += 1

                # Заголовок
                for col, header in enumerate(output_headers):
                    worksheet.write(0, col, header, self._cached_formats["header"])
                row_count = 1

            worksheet.write_row(row_count, 0, output_row, self._cached_formats["cell"])
            row_count += 1
            return True

        self._process_rows_with_progress(reader, filter_idx, current_file, process_row)

    def stop(self):
        """Останавливает конвертацию."""
        self.stop_flag = True


# ============================================================================
# ПИВОТ-ТАБЛИЦЫ
# ============================================================================


class PivotTableProcessor:
    """
    Обработчик для создания сводных таблиц.
    """

    def __init__(self, log_callback=None):
        self.log_callback = log_callback or (lambda _msg, _color: None)

    def create_pivot_data(
        self,
        file_path: str,
        settings: Dict[str, Any],
        filter_column: str = "",
        filter_values: Set[str] = None,
    ) -> Optional[Dict]:
        """
        Создаёт данные для сводной таблицы.

        Args:
            file_path: Путь к файлу
            settings: Настройки сводной таблицы
            filter_column: Столбец для фильтра
            filter_values: Значения фильтра

        Returns:
            Словарь с данными сводной таблицы или None
        """
        try:
            self.log_callback("Создание сводной таблицы...", QColor("blue"))

            encoding = FileUtilities.get_encoding(file_path)
            delimiter = FileUtilities.get_delimiter(file_path)

            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                reader = csv.reader(f, delimiter=delimiter)

                try:
                    headers = next(reader)
                except StopIteration:
                    self.log_callback("Пустой файл", QColor("red"))
                    return None

                # Индексы
                try:
                    row_indices = [
                        headers.index(row) for row in settings.get("rows", [])
                    ]
                    col_indices = [
                        headers.index(col) for col in settings.get("columns", [])
                    ]
                    value_settings = settings.get("values", [])
                except ValueError as e:
                    self.log_callback(f"Ошибка: колонка не найдена: {e}", QColor("red"))
                    return None

                # Индекс фильтра
                filter_idx = None
                if filter_column and filter_column != "Не фильтровать":
                    try:
                        filter_idx = headers.index(filter_column)
                    except ValueError:
                        filter_idx = None

                # Подготовка к удалению дубликатов
                remove_duplicates = settings.get("remove_duplicates", False)
                seen_combinations = set()

                # Индексы для проверки уникальности (строки + столбцы + значения)
                dedup_indices = []
                if remove_duplicates:
                    dedup_indices = list(set(row_indices + col_indices))
                    for val_setting in value_settings:
                        try:
                            val_idx = headers.index(val_setting["field"])
                            if val_idx not in dedup_indices:
                                dedup_indices.append(val_idx)
                        except ValueError:
                            pass

                # Агрегация (Online Aggregation)
                def get_default_agg_state():
                    return {
                        "sum": 0.0,
                        "count": 0,
                        "max": float("-inf"),
                        "min": float("inf"),
                    }

                pivot_data = defaultdict(
                    lambda: defaultdict(lambda: defaultdict(get_default_agg_state))
                )
                duplicates_removed = 0

                for row in reader:
                    # Фильтр
                    if filter_idx is not None and filter_values:
                        if filter_idx < len(row):
                            if row[filter_idx] not in filter_values:
                                continue

                    # Удаление дубликатов (на лету)
                    if remove_duplicates:
                        combination = tuple(
                            row[idx] for idx in dedup_indices if idx < len(row)
                        )
                        if combination in seen_combinations:
                            duplicates_removed += 1
                            continue
                        seen_combinations.add(combination)

                    try:
                        row_key = (
                            tuple(row[idx] for idx in row_indices)
                            if row_indices
                            else ("Итого",)
                        )
                        col_key = (
                            tuple(row[idx] for idx in col_indices)
                            if col_indices
                            else ("Итого",)
                        )

                        for val_setting in value_settings:
                            try:
                                val_idx = headers.index(val_setting["field"])
                                value = row[val_idx] if val_idx < len(row) else ""
                                key = f"{val_setting['field']}_{val_setting['aggregation']}"

                                state = pivot_data[row_key][col_key][key]

                                if val_setting["aggregation"] == "Количество":
                                    state["count"] += 1
                                else:
                                    try:
                                        num_value = float(value) if value else 0.0
                                    except (ValueError, AttributeError):
                                        num_value = 0.0

                                    state["sum"] += num_value
                                    state["count"] += 1
                                    if num_value > state["max"]:
                                        state["max"] = num_value
                                    if num_value < state["min"]:
                                        state["min"] = num_value
                            except Exception:
                                continue
                    except Exception:
                        continue

            if remove_duplicates:
                self.log_callback(
                    f"Удалено дубликатов: {duplicates_removed}", QColor("blue")
                )

            # Агрегируем (вычисляем финальные значения)
            aggregated = {}
            for row_key in pivot_data:
                aggregated[row_key] = {}
                for col_key in pivot_data[row_key]:
                    aggregated[row_key][col_key] = {}

                    for val_setting in value_settings:
                        key = f"{val_setting['field']}_{val_setting['aggregation']}"
                        state = pivot_data[row_key][col_key][key]

                        if state["count"] > 0:
                            if val_setting["aggregation"] == "Сумма":
                                aggregated[row_key][col_key][key] = state["sum"]
                            elif val_setting["aggregation"] == "Среднее":
                                aggregated[row_key][col_key][key] = (
                                    state["sum"] / state["count"]
                                )
                            elif val_setting["aggregation"] == "Количество":
                                aggregated[row_key][col_key][key] = state["count"]
                            elif val_setting["aggregation"] == "Максимум":
                                aggregated[row_key][col_key][key] = state["max"]
                            elif val_setting["aggregation"] == "Минимум":
                                aggregated[row_key][col_key][key] = state["min"]
                        else:
                            aggregated[row_key][col_key][key] = 0

            return aggregated

        except Exception as e:
            self.log_callback(f"Ошибка создания сводной таблицы: {e}", QColor("red"))
            return None

    def write_pivot_table(
        self,
        workbook: xlsxwriter.Workbook,
        pivot_data: Dict,
        settings: Dict[str, Any],
        cached_formats: Dict[str, Any],
    ):
        """
        Записывает сводную таблицу в книгу Excel.

        Args:
            workbook: Книга Excel
            pivot_data: Данные сводной таблицы
            settings: Настройки
            cached_formats: Кэш форматов
        """
        if not pivot_data:
            return

        worksheet = workbook.add_worksheet("Сводная таблица")

        row_values = sorted(pivot_data.keys())
        col_values = sorted(
            set(col for row in pivot_data.values() for col in row.keys())
        )
        value_settings = settings.get("values", [])

        # Создаём форматы для сводной таблицы
        header_format = cached_formats.get("header")

        # Формат для строк (без заливки, выравнивание влево)
        row_label_format = workbook.add_format(
            {
                "font_size": self.styles.get("font_size", 12)
                if hasattr(self, "styles")
                else 12,
                "font_name": self.styles.get("font_name", "Arial")
                if hasattr(self, "styles")
                else "Arial",
                "align": "left",
                "valign": "vcenter",
            }
        )

        # Формат для значений (без заливки, выравнивание вправо)
        value_format = workbook.add_format(
            {
                "font_size": self.styles.get("font_size", 12)
                if hasattr(self, "styles")
                else 12,
                "font_name": self.styles.get("font_name", "Arial")
                if hasattr(self, "styles")
                else "Arial",
                "align": "right",
                "valign": "vcenter",
            }
        )

        # Заголовки строк
        current_col = 0
        for row_header in settings.get("rows", []):
            worksheet.write(0, current_col, row_header, header_format)
            current_col += 1

        # Заголовки столбцов
        for col_value in col_values:
            for val_setting in value_settings:
                header_text = f"{' / '.join(str(v) for v in col_value)} - {val_setting['field']} ({val_setting['aggregation']})"
                worksheet.write(0, current_col, header_text, header_format)
                current_col += 1

        # Данные
        for row_idx, row_value in enumerate(row_values, 1):
            current_col = 0

            # Значения строк (выравнивание влево, без заливки)
            if isinstance(row_value, tuple):
                for value in row_value:
                    worksheet.write(row_idx, current_col, str(value), row_label_format)
                    current_col += 1
            else:
                worksheet.write(row_idx, 0, str(row_value), row_label_format)
                current_col = 1

            # Агрегированные значения (выравнивание вправо, без заливки)
            for col_value in col_values:
                for val_setting in value_settings:
                    key = f"{val_setting['field']}_{val_setting['aggregation']}"
                    value = pivot_data[row_value][col_value].get(key, 0)

                    if isinstance(value, float):
                        worksheet.write(
                            row_idx, current_col, round(value, 2), value_format
                        )
                    else:
                        worksheet.write(row_idx, current_col, value, value_format)
                    current_col += 1

        # Авто-ширина столбцов
        for col_idx in range(current_col):
            worksheet.set_column(col_idx, col_idx, 15)

        self.log_callback("Сводная таблица создана", QColor("green"))
