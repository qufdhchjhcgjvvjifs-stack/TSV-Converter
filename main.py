"""
Главный файл запуска конвертера TSV/CSV в Excel.
Интегрирует GUI (gui.py) и бизнес-логику (converter.py).
"""

import sys
import os
import csv
import hashlib
from datetime import datetime
from collections import defaultdict

from PySide6.QtWidgets import QApplication, QMessageBox, QFileDialog, QDialog
from PySide6.QtCore import QTimer, QSettings, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices

from gui import (
    MainWindow,
    ColumnValuesDialog,
    PivotSettingsDialog,
    TSVPreviewDialog,
    apply_theme_to_messagebox,
    LoadingWorker,
)
from converter import TSVToExcelConverter, FileUtilities, ConversionConfig


class TSVConverterApp:
    """
    Основное приложение.
    Связывает GUI и бизнес-логику.
    """

    def __init__(self):
        # Настройки
        self.settings = QSettings("TSVConverter", "App")

        # Главное окно
        self.window = MainWindow()

        # Конвертер
        self.converter: TSVToExcelConverter = None

        # Таймер
        self._timer = QTimer()
        self._start_time = None

        # Воркеры подсчета строк (чтобы сборщик мусора не убил активный поток)
        self._active_workers = []

        # Загрузка настроек
        self._load_settings()

        # Подключение сигналов
        self._connect_signals()

        # Таймер
        self._timer.timeout.connect(self._update_timer)

    @staticmethod
    def _dedup_key(values):
        """Компактный ключ дедупликации без хранения всех строковых значений."""
        digest = hashlib.blake2b(digest_size=16)
        for value in values:
            data = value.encode("utf-8", errors="replace")
            digest.update(str(len(data)).encode("ascii"))
            digest.update(b":")
            digest.update(data)
        return digest.digest()

    def _connect_signals(self):
        """Подключает сигналы GUI к обработчикам."""
        # Кнопки главного окна
        self.window.conversion_started.connect(self._start_conversion)
        self.window.conversion_stopped.connect(self._stop_conversion)
        self.window.files_added.connect(self._on_files_added)

        # Сигналы действий
        self.window.open_converted_file_requested.connect(self._open_converted_file)
        self.window.delete_converted_file_requested.connect(self._delete_converted_file)
        self.window.export_report_requested.connect(self._export_report)
        self.window.preview_requested.connect(self._preview_file)
        self.window.pivot_settings_requested.connect(self._show_pivot_settings)

        # Сигналы
        self.window.settings_saved.connect(self._on_settings_saved)
        self.window.columns_changed.connect(self._on_columns_changed)

        # Комбобоксы
        self.window.file_split_column_combo.currentIndexChanged.connect(
            self._on_file_split_column_selected
        )
        self.window.split_column_combo.currentIndexChanged.connect(
            self._on_split_column_selected
        )
        self.window.filter_column_combo.currentIndexChanged.connect(
            self._on_filter_column_selected
        )

        # Конвертер

    def _load_settings(self):
        """Загружает сохранённые настройки."""
        theme = self.settings.value("theme", "Светлая")
        default_path = self.settings.value("default_path", "")
        auto_open = self.settings.value("auto_open", "false") == "true"
        auto_delete = self.settings.value("auto_delete", "false") == "true"
        ram_threshold = int(self.settings.value("ram_threshold", 500000))

        # Применяем к окну
        self.window._settings = {
            "theme": theme,
            "default_path": default_path,
            "auto_open": auto_open,
            "auto_delete": auto_delete,
            "ram_threshold": ram_threshold,
        }

        if default_path:
            self.window.output_path_edit.setText(default_path)

        self.window._apply_theme()

    def _save_settings(self):
        """Сохраняет настройки."""
        self.settings.setValue("theme", self.window._settings.get("theme", "Светлая"))
        self.settings.setValue(
            "default_path", self.window._settings.get("default_path", "")
        )
        self.settings.setValue(
            "auto_open", str(self.window._settings.get("auto_open", False))
        )
        self.settings.setValue(
            "auto_delete", str(self.window._settings.get("auto_delete", False))
        )
        self.settings.setValue(
            "ram_threshold", str(self.window._settings.get("ram_threshold", 500000))
        )

    # ========================================================================
    # ОБРАБОТЧИКИ СОБЫТИЙ
    # ========================================================================

    def _on_files_added(self, files: list):
        """Обработчик добавления файлов."""
        # Файлы уже обработаны в gui.py, просто логируем
        if files:
            self._log_message(
                f"Файлы готовы к конвертации: {len(files)} шт.", QColor("blue")
            )

    def _update_total_rows(self):
        """Подсчитывает общее количество строк (асинхронно)."""
        files = [
            self.window.file_list.item(i).text()
            for i in range(self.window.file_list.count())
        ]

        if not files:
            self.window.total_rows_label.setText("Строк: 0")
            return

        filter_col = self.window.filter_column_combo.currentText()
        filter_values = dict(self.window._filter_values)
        file_split_col = self.window.file_split_column_combo.currentText()
        sheet_split_col = self.window.split_column_combo.currentText()
        file_split_values = dict(self.window._file_split_values)
        sheet_split_values = dict(self.window._sheet_split_values)
        selected_columns = (
            list(self.window._selected_output_columns) if len(files) == 1 else []
        )

        # Обновляем UI
        self.window.total_rows_label.setText("Строк: Подсчет...")

        # Запускаем в отдельном потоке
        worker = LoadingWorker(
            self._count_rows_task,
            files,
            filter_col,
            filter_values,
            file_split_col,
            file_split_values,
            sheet_split_col,
            sheet_split_values,
            selected_columns,
        )
        worker.finished.connect(
            lambda total: self.window.total_rows_label.setText(f"Строк: {total:,}")
        )
        worker.error.connect(
            lambda err: self._log_message(
                f"Ошибка подсчёта строк: {err}", QColor("red")
            )
        )

        # Сначала очищаем завершенные потоки
        self._active_workers = [w for w in self._active_workers if w.isRunning()]
        # Затем сохраняем ссылку на новый, чтобы избежать "QThread: Destroyed"
        self._active_workers.append(worker)

        worker.start()

    @staticmethod
    def _count_rows_task(
        files,
        filter_col,
        filter_values,
        file_split_col="",
        file_split_values=None,
        sheet_split_col="",
        sheet_split_values=None,
        selected_columns=None,
    ):
        """Фоновая задача подсчета строк."""
        file_split_values = file_split_values or {}
        sheet_split_values = sheet_split_values or {}
        selected_columns = selected_columns or []
        total = 0
        for file_path in files:
            try:
                encoding = FileUtilities.get_encoding(file_path)
                delimiter = FileUtilities.get_delimiter(file_path)

                filter_idx = None
                filter_vals = None
                file_split_idx = None
                file_split_vals = set()
                sheet_split_idx = None
                sheet_split_vals = set()
                output_indices = None
                seen_rows = set()
                seen_rows_by_destination = defaultdict(set)

                with open(file_path, "r", encoding=encoding, errors="replace") as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    headers = next(reader, [])

                    if filter_col and filter_col != "Не фильтровать":
                        try:
                            filter_idx = headers.index(filter_col)
                            filter_vals = set(filter_values.get(filter_col, []))
                        except ValueError:
                            filter_idx = None

                    if file_split_col and file_split_col != "Не разделять на файлы":
                        try:
                            file_split_idx = headers.index(file_split_col)
                            file_split_vals = set(file_split_values.get(file_split_col, []))
                        except ValueError:
                            file_split_idx = None

                    if sheet_split_col and sheet_split_col != "Не разделять на листы":
                        try:
                            sheet_split_idx = headers.index(sheet_split_col)
                            sheet_split_vals = set(sheet_split_values.get(sheet_split_col, []))
                        except ValueError:
                            sheet_split_idx = None

                    if selected_columns and len(selected_columns) < len(headers):
                        header_to_index = {}
                        for index, header in enumerate(headers):
                            if header not in header_to_index:
                                header_to_index[header] = index
                        output_indices = [
                            header_to_index[column]
                            for column in selected_columns
                            if column in header_to_index
                        ]

                    for row in reader:
                        if filter_idx is not None and filter_vals:
                            if filter_idx >= len(row) or row[filter_idx] not in filter_vals:
                                continue

                        file_split_value = None
                        if file_split_idx is not None:
                            value = row[file_split_idx] if file_split_idx < len(row) else ""
                            if not value or not value.strip():
                                continue
                            file_split_value = (
                                "Остальные"
                                if file_split_vals and value not in file_split_vals
                                else value
                            )

                        sheet_split_value = None
                        if sheet_split_idx is not None:
                            value = row[sheet_split_idx] if sheet_split_idx < len(row) else ""
                            if not value or not value.strip():
                                continue
                            sheet_split_value = (
                                "Остальные"
                                if sheet_split_vals and value not in sheet_split_vals
                                else value
                            )

                        if output_indices:
                            row_key = TSVConverterApp._dedup_key(
                                row[index] if index < len(row) else ""
                                for index in output_indices
                            )
                            if file_split_value is not None or sheet_split_value is not None:
                                destination_key = (file_split_value, sheet_split_value)
                                if row_key in seen_rows_by_destination[destination_key]:
                                    continue
                                seen_rows_by_destination[destination_key].add(row_key)
                            else:
                                if row_key in seen_rows:
                                    continue
                                seen_rows.add(row_key)

                        total += 1
            except (OSError, IOError, UnicodeDecodeError):
                pass
        return total

    def _on_columns_changed(self):
        """Обновляет зависящие от состава столбцов данные."""
        self._update_total_rows()
        self._update_pivot_settings_filter()

    @staticmethod
    def _count_split_distribution_task(
        files,
        file_split_col,
        file_split_values,
        sheet_split_col,
        sheet_split_values,
        filter_col,
        filter_values_dict,
        selected_columns=None,
    ):
        """Фоновая задача подсчета распределения строк по значениям разделения."""
        counts = defaultdict(lambda: defaultdict(int))
        file_selected_set = set(file_split_values)
        sheet_selected_set = set(sheet_split_values)
        selected_columns = selected_columns or []
        seen_rows_by_destination = defaultdict(set)
        total_rows = 0
        
        filter_vals = None
        if filter_col and filter_col != "Не фильтровать":
            filter_vals = set(filter_values_dict.get(filter_col, []))

        def get_split_value(row, split_idx, selected_set):
            val = row[split_idx] if split_idx < len(row) else ""
            if not val or not val.strip():
                return ""
            if selected_set and val not in selected_set:
                return "Остальные"
            return val

        for file_path in files:
            try:
                encoding = FileUtilities.get_encoding(file_path)
                delimiter = FileUtilities.get_delimiter(file_path)

                with open(file_path, "r", encoding=encoding, errors="replace") as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    try:
                        headers = next(reader)
                        file_split_idx = None
                        sheet_split_idx = None

                        if file_split_col and file_split_col != "Не разделять на файлы":
                            if file_split_col not in headers:
                                continue
                            file_split_idx = headers.index(file_split_col)

                        if sheet_split_col and sheet_split_col != "Не разделять на листы":
                            if sheet_split_col not in headers:
                                continue
                            sheet_split_idx = headers.index(sheet_split_col)

                        output_indices = None
                        if selected_columns and len(selected_columns) < len(headers):
                            header_to_index = {}
                            for index, header in enumerate(headers):
                                if header not in header_to_index:
                                    header_to_index[header] = index
                            output_indices = [
                                header_to_index[column]
                                for column in selected_columns
                                if column in header_to_index
                            ]
                        
                        filter_idx = None
                        if filter_vals:
                            try:
                                filter_idx = headers.index(filter_col)
                            except ValueError:
                                filter_idx = None
                        
                        for row in reader:
                            # Фильтр
                            if filter_idx is not None:
                                if filter_idx >= len(row) or row[filter_idx] not in filter_vals:
                                    continue

                            file_key = ""
                            if file_split_idx is not None:
                                file_key = get_split_value(
                                    row, file_split_idx, file_selected_set
                                )
                                if not file_key:
                                    continue

                            sheet_key = ""
                            if sheet_split_idx is not None:
                                sheet_key = get_split_value(
                                    row, sheet_split_idx, sheet_selected_set
                                )
                                if not sheet_key:
                                    continue

                            if output_indices:
                                row_key = TSVConverterApp._dedup_key(
                                    row[idx] if idx < len(row) else ""
                                    for idx in output_indices
                                )
                                destination_key = (file_key, sheet_key)
                                if row_key in seen_rows_by_destination[destination_key]:
                                    continue
                                seen_rows_by_destination[destination_key].add(row_key)

                            counts[file_key][sheet_key] += 1
                            total_rows += 1
                                
                    except (ValueError, StopIteration):
                        continue
            except (OSError, IOError, UnicodeDecodeError):
                pass
        return {
            "file_column": file_split_col,
            "sheet_column": sheet_split_col,
            "counts": {key: dict(value) for key, value in counts.items()},
            "total_rows": total_rows,
        }

    def _on_split_distribution_calculated(self, result):
        """Обработчик завершения подсчета распределения."""
        self.window._hide_loading_overlay()

        counts = result.get("counts", {}) if isinstance(result, dict) else {}
        total_rows = result.get("total_rows", 0) if isinstance(result, dict) else 0
        file_column = result.get("file_column", "") if isinstance(result, dict) else ""
        sheet_column = result.get("sheet_column", "") if isinstance(result, dict) else ""
        has_file_split = bool(file_column and file_column != "Не разделять на файлы")
        has_sheet_split = bool(sheet_column and sheet_column != "Не разделять на листы")

        if not counts:
            self._log_message("Анализ разделения: нет данных для распределения", QColor("orange"))
            return

        self._log_message("=== Прогноз разделения ===", QColor("cyan"))
        if has_file_split:
            self._log_message(f"Файлы по столбцу: {file_column}", QColor("blue"))
        if has_sheet_split:
            self._log_message(f"Листы по столбцу: {sheet_column}", QColor("blue"))

        if has_file_split and has_sheet_split:
            sorted_files = sorted(
                counts.items(), key=lambda item: sum(item[1].values()), reverse=True
            )
            for file_idx, (file_value, sheet_counts) in enumerate(sorted_files, start=1):
                file_total = sum(sheet_counts.values())
                self._log_message(
                    f"{file_idx}. Файл: {file_value}: {file_total:,} строк",
                    QColor("gray"),
                )
                sorted_sheets = sorted(
                    sheet_counts.items(), key=lambda item: item[1], reverse=True
                )
                for sheet_idx, (sheet_value, count) in enumerate(sorted_sheets, start=1):
                    self._log_message(
                        f"   {sheet_idx}. Лист: {sheet_value}: {count:,} строк",
                        QColor("gray"),
                    )
        elif has_file_split:
            sorted_files = sorted(
                counts.items(), key=lambda item: sum(item[1].values()), reverse=True
            )
            for idx, (file_value, sheet_counts) in enumerate(sorted_files, start=1):
                self._log_message(
                    f"{idx}. Файл: {file_value}: {sum(sheet_counts.values()):,} строк",
                    QColor("gray"),
                )
        else:
            sheet_counts = counts.get("", {})
            for idx, (sheet_value, count) in enumerate(
                sorted(sheet_counts.items(), key=lambda item: item[1], reverse=True),
                start=1,
            ):
                self._log_message(
                    f"{idx}. Лист: {sheet_value}: {count:,} строк",
                    QColor("gray"),
                )

        item_type = "файлов" if has_file_split else "листов"
        item_count = len(counts) if has_file_split else len(counts.get("", {}))
        self._log_message("-" * 40, QColor("cyan"))
        self._log_message(f"Итого будет создано {item_type}: {item_count}", QColor("blue"))
        if has_file_split and has_sheet_split:
            total_sheets = sum(len(sheet_counts) for sheet_counts in counts.values())
            self._log_message(f"Итого листов внутри файлов: {total_sheets}", QColor("blue"))
        self._log_message(f"Всего строк к распределению: {total_rows:,}", QColor("blue"))

    def _on_split_distribution_error(self, error):
        """Обработчик ошибки подсчета распределения."""
        self.window._hide_loading_overlay()
        self._log_message(f"Ошибка анализа распределения: {error}", QColor("red"))

    def _on_split_column_selected(self, index: int):
        """Обработчик выбора столбца для разделения на листы."""
        self._on_split_dimension_selected(index, "sheets")

    def _on_file_split_column_selected(self, index: int):
        """Обработчик выбора столбца для разделения на файлы."""
        self._on_split_dimension_selected(index, "files")

    def _on_split_dimension_selected(self, index: int, dimension: str):
        """Общий обработчик выбора столбцов разделения."""
        if index <= 0:
            return

        combo = (
            self.window.file_split_column_combo
            if dimension == "files"
            else self.window.split_column_combo
        )
        storage = (
            self.window._file_split_values
            if dimension == "files"
            else self.window._sheet_split_values
        )
        column = combo.currentText()
        if not self.window.file_list.count():
            return

        try:
            # Получаем текущий фильтр
            filter_col = self.window.filter_column_combo.currentText()
            filter_vals_list = (
                list(self.window._filter_values.get(filter_col, []))
                if filter_col != "Не фильтровать"
                else []
            )

            # Создаём диалог с асинхронной загрузкой
            file_paths = [
                self.window.file_list.item(i).text()
                for i in range(self.window.file_list.count())
            ]
            dialog = ColumnValuesDialog(
                parent=self.window,
                file_paths=file_paths,
                column=column,
                filter_column=filter_col,
                filter_values=filter_vals_list,
                selected_values=list(storage.get(column, [])),
            )

            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected = dialog.get_selected_values()
                storage[column] = selected
                self._update_total_rows()
                
                # Показываем лоадер на главном окне
                self.window._show_loading_overlay("Анализ распределения данных...")

                file_split_col = self.window.file_split_column_combo.currentText()
                sheet_split_col = self.window.split_column_combo.currentText()
                file_split_values = (
                    self.window._file_split_values.get(file_split_col, [])
                    if file_split_col != "Не разделять на файлы"
                    else []
                )
                sheet_split_values = (
                    self.window._sheet_split_values.get(sheet_split_col, [])
                    if sheet_split_col != "Не разделять на листы"
                    else []
                )
                
                # Запускаем подсчет
                worker = LoadingWorker(
                    self._count_split_distribution_task,
                    file_paths,
                    file_split_col,
                    file_split_values,
                    sheet_split_col,
                    sheet_split_values,
                    filter_col,
                    self.window._filter_values,
                    list(self.window._selected_output_columns),
                )
                
                worker.finished.connect(self._on_split_distribution_calculated)
                worker.error.connect(
                    lambda err: self._on_split_distribution_error(err)
                )
                
                # Очищаем и сохраняем воркера
                self._active_workers = [w for w in self._active_workers if w.isRunning()]
                self._active_workers.append(worker)
                worker.start()
            else:
                combo.setCurrentIndex(0)

        except Exception as e:
            self._log_message(f"Ошибка: {e}", QColor("red"))
            combo.setCurrentIndex(0)

    def _on_filter_column_selected(self, index: int):
        """Обработчик выбора столбца для фильтра."""
        if index <= 0:
            column = self.window.filter_column_combo.currentText()
            if column in self.window._filter_values:
                del self.window._filter_values[column]
                self._update_total_rows()
                self._update_pivot_settings_filter()
            return

        column = self.window.filter_column_combo.currentText()
        if not self.window.file_list.count():
            return

        try:
            # Получаем текущий фильтр (для вложенного фильтра)
            filter_col = self.window.filter_column_combo.currentText()
            filter_vals = (
                list(self.window._filter_values.get(filter_col, []))
                if filter_col != "Не фильтровать"
                else []
            )

            # Создаём диалог с асинхронной загрузкой
            file_paths = [
                self.window.file_list.item(i).text()
                for i in range(self.window.file_list.count())
            ]
            dialog = ColumnValuesDialog(
                parent=self.window,
                file_paths=file_paths,
                column=column,
                filter_column=filter_col,
                filter_values=filter_vals,
                selected_values=list(self.window._filter_values.get(column, [])),
            )

            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected = dialog.get_selected_values()
                self.window._filter_values[column] = selected
                self._update_total_rows()
                self._update_pivot_settings_filter()
                self._log_message(
                    f"Применён фильтр: {', '.join(selected[:5])}...", QColor("green")
                )
            else:
                self.window.filter_column_combo.setCurrentIndex(0)
                if column in self.window._filter_values:
                    del self.window._filter_values[column]
                    self._update_total_rows()
                    self._update_pivot_settings_filter()

        except Exception as e:
            self._log_message(f"Ошибка: {e}", QColor("red"))
            self.window.filter_column_combo.setCurrentIndex(0)

    def _update_pivot_settings_filter(self):
        """Обновляет настройки фильтра в настройках сводной таблицы."""
        if self.window._pivot_settings:
            filter_col = self.window.filter_column_combo.currentText()
            if filter_col != "Не фильтровать":
                self.window._pivot_settings["filter_column"] = filter_col
                self.window._pivot_settings["filter_values"] = list(
                    self.window._filter_values.get(filter_col, [])
                )
            else:
                self.window._pivot_settings["filter_column"] = ""
                self.window._pivot_settings["filter_values"] = []

    def _start_conversion(self, config: ConversionConfig):
        """Запускает конвертацию."""
        # Создаём конвертер
        self.converter = TSVToExcelConverter(
            input_files=config.input_files,
            output_directory=config.output_directory,
            output_format=config.output_format,
            auto_open=config.auto_open,
            auto_delete=config.auto_delete,
            styles=config.styles,
            header_color=config.header_color,
            file_split_column=config.file_split_column,
            file_split_values=config.file_split_values,
            split_column=config.split_column,
            split_mode=config.split_mode,
            selected_values=config.selected_values,
            filter_column=config.filter_column,
            filter_values=config.filter_values,
            pivot_settings=config.pivot_settings,
            selected_columns=config.selected_columns,
            deduplicate_rows=config.deduplicate_rows,
            ram_threshold=config.ram_threshold,
        )

        # Подключаем сигналы
        self.converter.progress_data.connect(
            lambda data: self.window.progress_bar.setValue(data.percent)
        )
        self.converter.progress_data.connect(self.window.update_progress_details)
        self.converter.log_message.connect(self._log_message)
        self.converter.finished_signal.connect(self._on_conversion_finished)
        self.converter.stopped_signal.connect(self._on_conversion_stopped)
        self.converter.error.connect(self._on_conversion_error)

        self.window.open_file_btn.setEnabled(False)
        self.window.delete_file_btn.setEnabled(False)

        # Запускаем
        self.converter.start()

        # Таймер
        self._start_time = datetime.now()
        self._timer.start(1000)

        self._log_message("Конвертация запущена...", QColor("blue"))

    def _stop_conversion(self):
        """Останавливает конвертацию."""
        if self.converter and self.converter.isRunning():
            self.converter.stop()
            self._log_message("Остановка конвертации...", QColor("orange"))

    def _on_conversion_finished(self):
        """Обработчик завершения конвертации."""
        self._timer.stop()
        has_output = bool(self.converter and self.converter.output_file_path)
        self.window.open_file_btn.setEnabled(has_output)
        self.window.delete_file_btn.setEnabled(has_output)
        self.window.start_btn.setEnabled(True)
        self.window.stop_btn.setEnabled(False)
        self.window.progress_bar.setValue(100)

        # Сброс панели прогресса
        self.window.reset_progress_details()

        if self._start_time:
            elapsed = (datetime.now() - self._start_time).total_seconds()
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self._log_message(
                f"Конвертация завершена за {minutes:02d}:{seconds:02d}", QColor("green")
            )

        msgbox = self._show_message_box(
            QMessageBox.Icon.Information, "Информация", "Конвертация завершена успешно!"
        )
        msgbox.exec()

        if self.converter and self.converter.output_file_path:
            if self.converter.auto_open and os.path.exists(
                self.converter.output_file_path
            ):
                QDesktopServices.openUrl(
                    QUrl.fromLocalFile(self.converter.output_file_path)
                )

            if self.converter.auto_delete and self.converter.input_files:
                for src_file in self.converter.input_files:
                    if os.path.exists(src_file):
                        try:
                            os.remove(src_file)
                            self._log_message(
                                f"Удалён исходный файл: {os.path.basename(src_file)}",
                                QColor("orange"),
                            )
                        except Exception as e:
                            self._log_message(
                                f"Не удалось удалить {os.path.basename(src_file)}: {e}",
                                QColor("red"),
                            )

    def _on_conversion_stopped(self):
        """Обработчик пользовательской остановки конвертации."""
        self._timer.stop()
        self.window.open_file_btn.setEnabled(False)
        self.window.delete_file_btn.setEnabled(False)
        self.window.start_btn.setEnabled(True)
        self.window.stop_btn.setEnabled(False)

        # Сброс панели прогресса
        self.window.reset_progress_details()

        if self._start_time:
            elapsed = (datetime.now() - self._start_time).total_seconds()
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self._log_message(
                f"Конвертация остановлена через {minutes:02d}:{seconds:02d}",
                QColor("orange"),
            )

        msgbox = self._show_message_box(
            QMessageBox.Icon.Warning,
            "Предупреждение",
            "Конвертация остановлена пользователем. Результат может быть неполным.",
        )
        msgbox.exec()

    def _on_conversion_error(self, error: str):
        """Обработчик ошибки конвертации."""
        self._timer.stop()
        self.window.open_file_btn.setEnabled(False)
        self.window.delete_file_btn.setEnabled(False)
        self.window.start_btn.setEnabled(True)
        self.window.stop_btn.setEnabled(False)

        # Сброс панели прогресса
        self.window.reset_progress_details()

        msgbox = self._show_message_box(
            QMessageBox.Icon.Critical, "Ошибка", f"Ошибка конвертации:\n{error}"
        )
        msgbox.exec()

    def _update_timer(self):
        """Обновляет таймер."""
        if self._start_time:
            elapsed = (datetime.now() - self._start_time).total_seconds()
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.window.timer_label.setText(f"Время: {minutes:02d}:{seconds:02d}")

    def _show_pivot_settings(self):
        """Показывает настройки сводной таблицы."""
        if not self.window.file_list.count():
            msgbox = self._show_message_box(
                QMessageBox.Icon.Warning, "Предупреждение", "Сначала добавьте файл"
            )
            msgbox.exec()
            return

        try:
            file_path = self.window.file_list.item(0).text()
            encoding = FileUtilities.get_encoding(file_path)
            delimiter = FileUtilities.get_delimiter(file_path)

            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                reader = csv.reader(f, delimiter=delimiter)
                headers = next(reader)

            active_headers = self.window._get_active_output_columns()
            if active_headers:
                headers = active_headers

            dialog = PivotSettingsDialog(headers, self.window)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.window._pivot_settings = dialog.get_settings()
                # Добавляем текущие значения фильтра
                filter_col = self.window.filter_column_combo.currentText()
                if filter_col != "Не фильтровать":
                    self.window._pivot_settings["filter_column"] = filter_col
                    self.window._pivot_settings["filter_values"] = list(
                        self.window._filter_values.get(filter_col, [])
                    )
                self._log_message(
                    "Настройки сводной таблицы сохранены", QColor("green")
                )

        except Exception as e:
            self._log_message(f"Ошибка: {e}", QColor("red"))

    def _preview_file(self):
        """Предпросмотр файла."""
        current_item = self.window.file_list.currentItem()
        if not current_item:
            msgbox = self._show_message_box(
                QMessageBox.Icon.Warning, "Предупреждение", "Выберите файл"
            )
            msgbox.exec()
            return

        dialog = TSVPreviewDialog(current_item.text(), self.window)
        dialog.exec()

    def _on_settings_saved(self, settings: dict):
        """Обработчик сохранения настроек."""
        self.window._settings = settings
        self._save_settings()

        # Применяем путь по умолчанию к полю вывода
        default_path = settings.get("default_path", "")
        if default_path and isinstance(default_path, str):
            self.window.output_path_edit.setText(default_path)

    def _open_converted_file(self):
        """Открывает сконвертированный файл."""
        if self.converter and hasattr(self.converter, "output_file_path"):
            path = self.converter.output_file_path
            if path and os.path.exists(path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            else:
                msgbox = self._show_message_box(
                    QMessageBox.Icon.Warning, "Ошибка", "Файл не найден"
                )
                msgbox.exec()

    def _delete_converted_file(self):
        """Удаляет сконвертированные файлы."""
        if not self.converter:
            return

        files_to_delete = []
        if hasattr(self.converter, "generated_files") and self.converter.generated_files:
            files_to_delete = [f for f in self.converter.generated_files if os.path.exists(f)]
        elif hasattr(self.converter, "output_file_path") and self.converter.output_file_path:
            path = self.converter.output_file_path
            if path and os.path.isfile(path):
                files_to_delete = [path]

        if not files_to_delete:
            return

        msg_text = (
            f"Удалить сгенерированные файлы ({len(files_to_delete)} шт.)?"
            if len(files_to_delete) > 1
            else "Удалить сгенерированный файл?"
        )
        
        msgbox = self._show_message_box(
            QMessageBox.Icon.Question,
            "Подтверждение",
            msg_text,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        reply = msgbox.exec()

        if reply == QMessageBox.StandardButton.Yes:
            deleted_count = 0
            for path in files_to_delete:
                try:
                    os.remove(path)
                    deleted_count += 1
                except Exception as e:
                    self._log_message(f"Ошибка удаления файла {os.path.basename(path)}: {str(e)}", QColor("red"))

            self.window.open_file_btn.setEnabled(False)
            self.window.delete_file_btn.setEnabled(False)
            
            if deleted_count > 0:
                msg = f"Удалено файлов: {deleted_count}" if deleted_count > 1 else "Файл удалён"
                self._log_message(msg, QColor("red"))

    def _export_report(self):
        """Экспортирует отчёт в HTML."""
        import html

        # Получаем данные из логов
        if not self.window.log_text.toPlainText().strip():
            msgbox = self._show_message_box(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Нет данных для экспорта отчёта. Сначала выполните конвертацию.",
            )
            msgbox.exec()
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self.window, "Сохранить отчёт", "", "HTML файлы (*.html);;Все файлы (*.*)"
        )

        if save_path:
            try:
                # Добавляем .html если нет расширения
                if not save_path.lower().endswith(".html"):
                    save_path += ".html"

                # Формируем HTML
                log_content = self.window.log_text.toPlainText()
                escaped_log = html.escape(log_content)

                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Отчёт о конвертации</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        .container {{ max-width: 900px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1 {{ color: #2c5282; border-bottom: 2px solid #4299e1; padding-bottom: 10px; }}
        h2 {{ color: #2d3748; margin-top: 20px; }}
        .log {{ background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 15px; font-family: Consolas, monospace; font-size: 12px; white-space: pre-wrap; }}
        .meta {{ color: #718096; font-size: 12px; margin-top: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 15px 0; }}
        th, td {{ border: 1px solid #cbd5e0; padding: 8px; text-align: left; }}
        th {{ background: #edf2f7; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Отчёт о конвертации TSV/CSV в Excel</h1>
        <p class="meta">Время создания: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}</p>
        
        <h2>📋 Журнал событий</h2>
        <div class="log">{escaped_log}</div>
        
        <h2>ℹ️ Параметры конвертации</h2>
        <table>
            <tr><th>Параметр</th><th>Значение</th></tr>
            <tr><td>Файлов в списке</td><td>{self.window.file_list.count()}</td></tr>
            <tr><td>Формат вывода</td><td>{self.window.format_combo.currentText()}</td></tr>
            <tr><td>Папка сохранения</td><td>{self.window.output_path_edit.text() or "Не указана"}</td></tr>
            <tr><td>Разделение на файлы по столбцу</td><td>{self.window.file_split_column_combo.currentText()}</td></tr>
            <tr><td>Разделение на листы по столбцу</td><td>{self.window.split_column_combo.currentText()}</td></tr>
            <tr><td>Фильтр по столбцу</td><td>{self.window.filter_column_combo.currentText()}</td></tr>
        </table>
        
        <p class="meta">Сгенерировано TSV Converter v11.0</p>
    </div>
</body>
</html>"""

                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(html_content)

                # Открываем в браузере
                QDesktopServices.openUrl(QUrl.fromLocalFile(save_path))

                self._log_message(
                    f"Отчёт экспортирован: {os.path.basename(save_path)}",
                    QColor("green"),
                )

            except (OSError, IOError, PermissionError) as e:
                msgbox = self._show_message_box(
                    QMessageBox.Icon.Critical,
                    "Ошибка",
                    f"Ошибка при экспорте отчёта:\n{str(e)}",
                )
                msgbox.exec()
                self._log_message(f"Ошибка экспорта отчёта: {e}", QColor("red"))

    def _log_message(self, message: str, color: QColor = QColor("black")):
        """Добавляет сообщение в лог."""
        self.window.log_message(message, color)

    def _show_message_box(
        self,
        icon: QMessageBox.Icon,
        title: str,
        text: str,
        buttons: QMessageBox.StandardButton = QMessageBox.StandardButton.Ok,
    ):
        """Создаёт QMessageBox с применённой темой."""
        msgbox = QMessageBox(icon, title, text, buttons, self.window)
        apply_theme_to_messagebox(msgbox, self.window._is_dark_theme)
        return msgbox

    def run(self):
        """Запускает приложение."""
        self.window.show()


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================


def main():
    """Точка входа в приложение."""
    from PySide6.QtWidgets import QStyleFactory

    # Настройка High DPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("TSV Converter")
    app.setOrganizationName("TSV Converter")

    app.setStyle(QStyleFactory.create("Fusion"))

    # Запуск приложения
    converter_app = TSVConverterApp()
    converter_app.run()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
