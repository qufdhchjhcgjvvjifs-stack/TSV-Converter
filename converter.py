"""
Модуль бизнес-логики для конвертера TSV/CSV в Excel.
Содержит классы для конвертации, утилиты и вспомогательные функции.
"""

import os
import csv
import time
from typing import List, Dict, Any, Optional, Set
from collections import defaultdict
from dataclasses import dataclass

from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QColor

import xlsxwriter


# ============================================================================
# ТИПЫ ДАННЫХ ДЛЯ ПРОГРЕССА
# ============================================================================


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


# ============================================================================
# УТИЛИТЫ (продолжение)
# ============================================================================


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
                    if file_path.lower().endswith(".csv"):
                        try:
                            with open(file_path, "r", encoding="utf-8") as f:
                                f.read(8192)
                            return "utf-8"
                        except UnicodeDecodeError:
                            return "windows-1251"
                    else:
                        return "utf-8"
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

    # Архитектурные константы
    MAX_EXCEL_ROWS = 1000000  # Ограничение в 1млн строк на один лист Excel
    STOP_CHECK_INTERVAL = 2000  # Как часто (в строках) проверять флаг остановки конвертации

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
        selected_values: Optional[List[str]] = None,
        filter_column: str = "",
        filter_values: Optional[List[str]] = None,
        pivot_settings: Optional[Dict[str, Any]] = None,
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
        self.selected_values = selected_values or []
        self.filter_column = filter_column
        self.filter_values = filter_values or []
        self.pivot_settings = pivot_settings

        self.stop_flag = False
        self.output_file_path: Optional[str] = None
        self.total_rows = 0
        self.processed_rows = 0

        # Трекер прогресса
        self.progress_tracker = ProgressTracker()

        # Кэш форматов
        self._cached_formats: Dict[str, Any] = {}

    def run(self):
        """Основной метод потока."""
        try:
            self.log_message.emit("Начало конвертации...", QColor("blue"))
            self.processed_rows = 0

            # Валидация
            if not self.input_files:
                raise ValueError("Нет файлов для конвертации")

            if not os.path.isdir(self.output_directory):
                raise ValueError(f"Директория не существует: {self.output_directory}")

            # Подсчёт общего количества строк
            self._count_total_rows()

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

                conversion_result = self._convert_file(tsv_file, processed)
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

            self.log_message.emit(
                f"Конвертация завершена. Обработано файлов: {processed}",
                QColor("green"),
            )
            self.finished_signal.emit()

        except Exception as e:
            self.log_message.emit(f"Критическая ошибка: {str(e)}", QColor("red"))
            self.error.emit(str(e))

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

                    # Считаем строки (оптимизированные циклы)
                    if filter_idx is not None and self.filter_values:
                        filter_vals = self.filter_values  # Локальная ссылка быстрее
                        for i, row in enumerate(reader):
                            if i % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                                return
                            if filter_idx < len(row) and row[filter_idx] in filter_vals:
                                self.total_rows += 1
                    else:
                        for i, _ in enumerate(reader):
                            if i % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                                return
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
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        if file_index > 0:
            base_name = f"{base_name}_{file_index + 1}"

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
        update_freq = max(1, self.total_rows // 100)

        with open(output_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(headers)

            if filter_idx is not None and self.filter_values:
                filter_vals = self.filter_values
                for i, row in enumerate(reader):
                    if i % 2000 == 0 and self.stop_flag:
                        break
                    if filter_idx < len(row) and row[filter_idx] not in filter_vals:
                        continue
                    writer.writerow(row)
                    self.processed_rows += 1
                    if self.processed_rows % update_freq == 0:
                        self._emit_progress_update(current_file, "Запись CSV...")
            else:
                for i, row in enumerate(reader):
                    if i % 2000 == 0 and self.stop_flag:
                        break
                    writer.writerow(row)
                    self.processed_rows += 1
                    if self.processed_rows % update_freq == 0:
                        self._emit_progress_update(current_file, "Запись CSV...")

        self._emit_progress_update(current_file, "Запись CSV...", force=True)

        self.output_file_path = output_path

    def _write_split_csv(self, reader, headers, base_name, split_idx, filter_idx):
        # Для сплита приходится держать открытые файлы или собирать буфер
        # Чтобы не держать тысячи файлов, будем собирать в словаре списков (осторожно с памятью!)
        # Или, лучше, проход в один поток, но открытие/закрытие (медленно).
        # Оптимально: словарь открытых хэндлов (до лимита ОС).

        open_files = {}
        writers = {}
        used_file_names: Set[str] = set()
        selected_vals = self.selected_values  # Локальная ссылка
        current_file = f"{base_name}_*.csv"
        update_freq = max(1, self.total_rows // 100)

        try:
            # Разделяем циклы для оптимизации
            if filter_idx is not None and self.filter_values:
                filter_vals = self.filter_values
                for i, row in enumerate(reader):
                    if i % 2000 == 0 and self.stop_flag:
                        break

                    if filter_idx < len(row) and row[filter_idx] not in filter_vals:
                        continue

                    self._process_split_row(
                        row,
                        split_idx,
                        selected_vals,
                        headers,
                        base_name,
                        open_files,
                        writers,
                        used_file_names,
                    )
                    self.processed_rows += 1
                    if self.processed_rows % update_freq == 0:
                        self._emit_progress_update(
                            current_file, "Распределение по CSV..."
                        )
            else:
                for i, row in enumerate(reader):
                    if i % 2000 == 0 and self.stop_flag:
                        break

                    self._process_split_row(
                        row,
                        split_idx,
                        selected_vals,
                        headers,
                        base_name,
                        open_files,
                        writers,
                        used_file_names,
                    )
                    self.processed_rows += 1
                    if self.processed_rows % update_freq == 0:
                        self._emit_progress_update(
                            current_file, "Распределение по CSV..."
                        )

        finally:
            for f in open_files.values():
                f.close()

        self._emit_progress_update(current_file, "Распределение по CSV...", force=True)

        # Указываем путь к папке как результат
        self.output_file_path = self.output_directory

    def _process_split_row(
        self,
        row,
        split_idx,
        selected_vals,
        headers,
        base_name,
        open_files,
        writers,
        used_file_names,
    ):
        """Вспомогательный метод для записи строки в нужный файл."""
        key = row[split_idx] if split_idx < len(row) else "Unknown"
        if selected_vals and key not in selected_vals:
            key = "Остальные"

        # Санитизация имени файла (только если новый ключ)
        if key not in writers:
            safe_key = FileUtilities.sanitize_file_stem(key, used_file_names)
            file_path = os.path.join(
                self.output_directory, f"{base_name}_{safe_key}.csv"
            )
            f = open(file_path, "w", encoding="utf-8-sig", newline="")
            writer = csv.writer(f, delimiter=";")
            writer.writerow(headers)
            open_files[key] = f
            writers[key] = writer

        writers[key].writerow(row)

    def _create_csv_pivot(self, input_file, base_name):
        try:
            processor = PivotTableProcessor(lambda _msg, _color: None)
            pivot_data = processor.create_pivot_data(
                input_file,
                self.pivot_settings,
                self.filter_column if self.filter_column != "Не фильтровать" else "",
                self.filter_values or [],
            )

            if pivot_data:
                output_path = os.path.join(
                    self.output_directory, f"{base_name}_Pivot.csv"
                )
                with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f, delimiter=";")

                    # Заголовки
                    row_headers = self.pivot_settings.get("rows", [])
                    # col_headers - removed as unused
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
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        if file_index > 0:
            base_name = f"{base_name}_{file_index + 1}"
        output_path = os.path.join(self.output_directory, f"{base_name}.xlsx")

        workbook = None
        result = "error"
        try:
            # Создаём книгу Excel
            workbook = xlsxwriter.Workbook(
                output_path,
                {"constant_memory": True, "use_zip64": True, "in_memory": False},
            )

            # Инициализируем форматы
            self._init_formats(workbook)

            # Открываем исходный файл
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

                # Конвертация
                if split_idx is not None:
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
                # Обработка сводной таблицы (ДО закрытия книги!)
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
                            self.filter_values or [],
                        )

                        if pivot_data:
                            # Создаём форматы для сводной
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
            # Гарантированно закрываем workbook
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass

            # При ошибке удаляем частичный файл
            if result != "success" and os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception:
                    pass

        if result == "success":
            self.output_file_path = output_path

            # Время обработки
            elapsed = time.time() - start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)

            self.log_message.emit(
                f"Файл сохранён: {os.path.basename(output_path)} ({minutes:02d}:{seconds:02d})",
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
    ):
        """Обрабатывает строки с фильтрацией, остановкой и обновлением прогресса."""
        update_freq = max(1, self.total_rows // 100)

        if filter_idx is not None and self.filter_values:
            filter_vals = self.filter_values
            for row_num, row in enumerate(reader, 1):
                if row_num % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                    break

                if filter_idx < len(row) and row[filter_idx] not in filter_vals:
                    continue

                row_handler(row)
                self.processed_rows += 1
                if self.processed_rows % update_freq == 0:
                    self._emit_progress_update(current_file, "Запись данных...")
        else:
            for row_num, row in enumerate(reader, 1):
                if row_num % self.STOP_CHECK_INTERVAL == 0 and self.stop_flag:
                    break

                row_handler(row)
                self.processed_rows += 1
                if self.processed_rows % update_freq == 0:
                    self._emit_progress_update(current_file, "Запись данных...")

        self._emit_progress_update(current_file, "Запись данных...", force=True)

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
        others_worksheet = None
        others_row_count = 0

        used_names: Set[str] = set()
        MAX_ROWS = self.MAX_EXCEL_ROWS
        selected_vals = self.selected_values

        # Вспомогательная функция для обработки одной строки
        def process_row(row):
            nonlocal others_worksheet, others_row_count
            value = row[split_idx] if split_idx < len(row) else "Unknown"

            # Проверяем выбранное значение
            if selected_vals and value not in selected_vals:
                # В "Остальные"
                if others_worksheet is None or others_row_count >= MAX_ROWS:
                    sheet_name = FileUtilities.sanitize_sheet_name(
                        "Остальные", used_names
                    )
                    others_worksheet = workbook.add_worksheet(sheet_name)
                    others_row_count = 0

                    # Заголовок
                    for col, header in enumerate(headers):
                        others_worksheet.write(
                            others_row_count,
                            col,
                            header,
                            self._cached_formats["header"],
                        )
                    others_row_count = 1

                others_worksheet.write_row(
                    others_row_count, 0, row, self._cached_formats["cell"]
                )
                others_row_count += 1

            else:
                # В свой лист
                if value not in worksheets:
                    sheet_name = FileUtilities.sanitize_sheet_name(value, used_names)
                    worksheets[value] = workbook.add_worksheet(sheet_name)
                    sheet_row_counts[value] = 0

                    # Заголовок
                    for col, header in enumerate(headers):
                        worksheets[value].write(
                            0, col, header, self._cached_formats["header"]
                        )
                    sheet_row_counts[value] = 1

                current_row = sheet_row_counts[value]
                if current_row >= MAX_ROWS:
                    # Новый лист для этого значения
                    sheet_name = FileUtilities.sanitize_sheet_name(
                        f"{value}_2", used_names
                    )
                    worksheets[value] = workbook.add_worksheet(sheet_name)
                    current_row = 0

                    for col, header in enumerate(headers):
                        worksheets[value].write(
                            current_row, col, header, self._cached_formats["header"]
                        )
                    current_row = 1

                worksheets[value].write_row(
                    current_row, 0, row, self._cached_formats["cell"]
                )
                sheet_row_counts[value] = current_row + 1

        self._process_rows_with_progress(reader, filter_idx, current_file, process_row)

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

        # Вспомогательная функция для обработки
        def process_row(row):
            nonlocal worksheet, row_count, sheet_num
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
                for col, header in enumerate(headers):
                    worksheet.write(0, col, header, self._cached_formats["header"])
                row_count = 1

            worksheet.write_row(row_count, 0, row, self._cached_formats["cell"])
            row_count += 1

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
        filter_values: List[str] = None,
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

                # Агрегация
                pivot_data = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
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

                                if val_setting["aggregation"] == "Количество":
                                    pivot_data[row_key][col_key][
                                        f"{val_setting['field']}_{val_setting['aggregation']}"
                                    ].append(1)
                                else:
                                    try:
                                        num_value = float(value) if value else 0
                                        pivot_data[row_key][col_key][
                                            f"{val_setting['field']}_{val_setting['aggregation']}"
                                        ].append(num_value)
                                    except (ValueError, AttributeError):
                                        pivot_data[row_key][col_key][
                                            f"{val_setting['field']}_{val_setting['aggregation']}"
                                        ].append(0)
                            except Exception:
                                continue
                    except Exception:
                        continue

            if remove_duplicates:
                self.log_callback(
                    f"Удалено дубликатов: {duplicates_removed}", QColor("blue")
                )

            # Агрегируем
            aggregated = {}
            for row_key in pivot_data:
                aggregated[row_key] = {}
                for col_key in pivot_data[row_key]:
                    aggregated[row_key][col_key] = {}

                    for val_setting in value_settings:
                        key = f"{val_setting['field']}_{val_setting['aggregation']}"
                        values = pivot_data[row_key][col_key][key]

                        if values:
                            if val_setting["aggregation"] == "Сумма":
                                aggregated[row_key][col_key][key] = sum(values)
                            elif val_setting["aggregation"] == "Среднее":
                                aggregated[row_key][col_key][key] = sum(values) / len(
                                    values
                                )
                            elif val_setting["aggregation"] == "Количество":
                                aggregated[row_key][col_key][key] = len(values)
                            elif val_setting["aggregation"] == "Максимум":
                                aggregated[row_key][col_key][key] = max(values)
                            elif val_setting["aggregation"] == "Минимум":
                                aggregated[row_key][col_key][key] = min(values)
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
