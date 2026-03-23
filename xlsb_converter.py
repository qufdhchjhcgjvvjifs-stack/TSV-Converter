"""
Модуль конвертера XLSB.
Изолированная логика для формата XLSB с использованием pyxlsbwriter.
"""

import os
import csv
import time
from typing import List, Optional, Dict, Any
from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QColor

try:
    from pyxlsbwriter import XlsbWriter

    XLSB_SUPPORT = True
except ImportError:
    XLSB_SUPPORT = False
    XlsbWriter = None

from converter import FileUtilities


class XLSBConverter(QThread):
    """
    Конвертер TSV/CSV в XLSB.
    Полностью изолирован от логики XLSX/CSV.
    Реализует тот же интерфейс что и TSVToExcelConverter для совместимости.
    """

    update_progress = Signal(int)
    progress_data = Signal(object)
    log_message = Signal(str, QColor)
    finished_signal = Signal()
    error = Signal(str)

    def __init__(
        self,
        input_files: List[str],
        output_directory: str,
        output_format: str = "xlsb",
        auto_open: bool = False,
        auto_delete: bool = False,
        styles: Optional[Dict[str, Any]] = None,
        header_color: str = "#C8DCF0",
        filter_column: str = "",
        filter_values: Optional[List[str]] = None,
        total_rows: int = 0,
        split_column: str = "",
        selected_values: Optional[List[str]] = None,
    ):
        super().__init__()

        self.input_files = input_files
        self.output_directory = output_directory
        self.output_format = output_format
        self.auto_open = auto_open
        self.auto_delete = auto_delete
        self.styles = styles or {}
        self.header_color = header_color
        self.filter_column = filter_column
        self.filter_values = filter_values or []
        self.split_column = split_column
        self.selected_values = selected_values or []

        self.stop_flag = False
        self.output_file_path: Optional[str] = None
        self.total_rows = total_rows
        self.processed_rows = 0

        self.start_time: Optional[float] = None
        self.last_update_time: float = 0
        self.min_update_interval: float = 0.1

    def run(self):
        """Основной метод потока."""
        from converter import ProgressData

        try:
            if not XLSB_SUPPORT:
                raise ImportError("pyxlsbwriter не установлен")

            if not self.input_files:
                raise ValueError("Нет файлов для конвертации")

            if not os.path.isdir(self.output_directory):
                raise ValueError(f"Директория не существует: {self.output_directory}")

            self.start_time = time.time()
            self.last_update_time = self.start_time

            self.log_message.emit("Начало конвертации в XLSB...", QColor("blue"))

            if (
                self.styles.get("bold")
                or self.styles.get("italic")
                or self.styles.get("font_size")
            ):
                self.log_message.emit(
                    "Примечание: Формат XLSB не поддерживает стили (жирный, курсив, шрифт и т.д.)",
                    QColor("orange"),
                )

            self._count_total_rows()
            self.log_message.emit(
                f"Всего строк для обработки: {self.total_rows:,}", QColor("blue")
            )

            processed = 0
            for tsv_file in self.input_files:
                if self.stop_flag:
                    self.log_message.emit(
                        "Конвертация прервана пользователем", QColor("orange")
                    )
                    break

                progress_data = ProgressData(
                    current_file=os.path.basename(tsv_file),
                    current_operation="Чтение файла...",
                )
                self.progress_data.emit(progress_data)

                if self._convert_file(tsv_file):
                    processed += 1

            self.log_message.emit(
                f"Конвертация завершена. Обработано файлов: {processed}",
                QColor("green"),
            )
            self.finished_signal.emit()

        except Exception as e:
            self.log_message.emit(f"Критическая ошибка: {str(e)}", QColor("red"))
            self.error.emit(str(e))

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

                    filter_idx = None
                    if self.filter_column and self.filter_column != "Не фильтровать":
                        try:
                            filter_idx = headers.index(self.filter_column)
                        except ValueError:
                            filter_idx = None

                    if filter_idx is not None and self.filter_values:
                        filter_vals = self.filter_values
                        for i, row in enumerate(reader):
                            if i % 2000 == 0 and self.stop_flag:
                                return
                            if filter_idx < len(row) and row[filter_idx] in filter_vals:
                                self.total_rows += 1
                    else:
                        for i, _ in enumerate(reader):
                            if i % 2000 == 0 and self.stop_flag:
                                return
                            self.total_rows += 1

            except PermissionError:
                self.log_message.emit(
                    f"Ошибка доступа к файлу {os.path.basename(tsv_file)}",
                    QColor("red"),
                )
            except Exception as e:
                self.log_message.emit(
                    f"Ошибка подсчёта строк {os.path.basename(tsv_file)}: {e}",
                    QColor("red"),
                )

    def _convert_file(self, input_file: str) -> bool:
        """Конвертирует один файл. Потоковая запись с минимальным RAM."""
        import gc
        from converter import ProgressData

        start_time = time.time()
        ROWS_PER_SHEET = 1000000

        try:
            encoding = FileUtilities.get_encoding(input_file)
            delimiter = FileUtilities.get_delimiter(input_file)

            self.log_message.emit(
                f"Обработка {os.path.basename(input_file)} (кодировка: {encoding})",
                QColor("blue"),
            )

            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_path = os.path.join(self.output_directory, f"{base_name}.xlsb")

            if not XLSB_SUPPORT:
                raise ImportError("pyxlsbwriter не установлен")

            update_freq = (
                max(1, self.total_rows // 100) if self.total_rows > 0 else 5000
            )

            with XlsbWriter(output_path, compressionLevel=6) as writer:
                f = open(input_file, "r", encoding=encoding, errors="replace")
                reader = csv.reader(f, delimiter=delimiter)

                try:
                    headers = next(reader)
                except StopIteration:
                    self.log_message.emit("Пустой файл", QColor("red"))
                    f.close()
                    return False

                filter_idx = None
                if self.filter_column and self.filter_column != "Не фильтровать":
                    try:
                        filter_idx = headers.index(self.filter_column)
                    except ValueError:
                        filter_idx = None

                sheet_num = 1
                rows_written = 0
                row_in_sheet = 0

                while True:
                    chunk = [headers]

                    for row in reader:
                        if self.stop_flag:
                            f.close()
                            return False

                        if filter_idx is not None and self.filter_values:
                            if (
                                filter_idx < len(row)
                                and row[filter_idx] not in self.filter_values
                            ):
                                continue

                        chunk.append(row)
                        rows_written += 1
                        row_in_sheet += 1
                        self.processed_rows += 1

                        if rows_written % update_freq == 0:
                            current_time = time.time()
                            elapsed = (
                                current_time - self.start_time if self.start_time else 0
                            )
                            speed = self.processed_rows / elapsed if elapsed > 0 else 0
                            percent = (
                                int(
                                    (self.processed_rows / max(self.total_rows, 1))
                                    * 100
                                )
                                if self.total_rows > 0
                                else 0
                            )
                            eta = (
                                (self.total_rows - self.processed_rows) / speed
                                if speed > 0 and self.total_rows > 0
                                else 0
                            )

                            progress_data = ProgressData(
                                percent=min(percent, 99),
                                processed_rows=self.processed_rows,
                                total_rows=self.total_rows,
                                elapsed_seconds=elapsed,
                                eta_seconds=eta,
                                rows_per_second=speed,
                                current_file=os.path.basename(input_file),
                                current_operation="Запись данных...",
                            )
                            self.progress_data.emit(progress_data)
                            self.update_progress.emit(min(percent, 99))

                        if row_in_sheet >= ROWS_PER_SHEET:
                            break

                    if len(chunk) <= 1:
                        break

                    sheet_name = f"Данные_{sheet_num}" if sheet_num > 1 else "Data"
                    writer.add_sheet(sheet_name)
                    writer.write_sheet(chunk)
                    self.log_message.emit(
                        f"Лист {sheet_name}: записано {len(chunk) - 1:,} строк",
                        QColor("blue"),
                    )
                    del chunk
                    gc.collect()

                    if row_in_sheet < ROWS_PER_SHEET:
                        break

                    sheet_num += 1
                    row_in_sheet = 0

                f.close()
                gc.collect()

            self.log_message.emit(f"Записано строк: {rows_written:,}", QColor("green"))
            self.output_file_path = output_path

            elapsed = time.time() - start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.log_message.emit(
                f"Файл сохранён: {os.path.basename(output_path)} ({minutes:02d}:{seconds:02d})",
                QColor("green"),
            )

            return True

        except PermissionError:
            self.log_message.emit(
                f"Ошибка доступа к файлу {os.path.basename(input_file)}",
                QColor("red"),
            )
            return False
        except Exception as e:
            self.log_message.emit(
                f"Ошибка конвертации {os.path.basename(input_file)}: {str(e)}",
                QColor("red"),
            )
            return False

    def stop(self):
        """Останавливает конвертацию."""
        self.stop_flag = True
