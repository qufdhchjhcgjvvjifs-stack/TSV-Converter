"""
Главный файл запуска конвертера TSV/CSV в Excel.
Интегрирует GUI (gui.py) и бизнес-логику (converter.py).
"""

import sys
import os
import csv
from datetime import datetime

from PySide6.QtWidgets import QApplication, QMessageBox, QFileDialog, QDialog
from PySide6.QtCore import QTimer, QSettings, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices

from gui import (
    MainWindow,
    ColumnValuesDialog,
    PivotSettingsDialog,
    TSVPreviewDialog,
    apply_theme_to_messagebox,
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

        # Загрузка настроек
        self._load_settings()

        # Подключение сигналов
        self._connect_signals()

        # Таймер
        self._timer.timeout.connect(self._update_timer)

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
        # Кнопка настроек уже подключена в MainWindow._show_settings

        # Сигналы
        self.window.settings_saved.connect(self._on_settings_saved)

        # Комбобоксы
        self.window.split_column_combo.currentIndexChanged.connect(
            self._on_split_column_selected
        )
        self.window.filter_column_combo.currentIndexChanged.connect(
            self._on_filter_column_selected
        )

        # Конвертер
        # (будет подключено при создании)

    def _load_settings(self):
        """Загружает сохранённые настройки."""
        theme = self.settings.value("theme", "Светлая")
        default_path = self.settings.value("default_path", "")
        auto_open = self.settings.value("auto_open", "false") == "true"
        auto_delete = self.settings.value("auto_delete", "false") == "true"

        # Применяем к окну
        self.window._settings = {
            "theme": theme,
            "default_path": default_path,
            "auto_open": auto_open,
            "auto_delete": auto_delete,
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
        """Подсчитывает общее количество строк."""
        total = 0

        for i in range(self.window.file_list.count()):
            file_path = self.window.file_list.item(i).text()

            try:
                encoding = FileUtilities.get_encoding(file_path)
                delimiter = FileUtilities.get_delimiter(file_path)

                # Индекс фильтра
                filter_idx = None
                filter_vals = None

                filter_col = self.window.filter_column_combo.currentText()
                if filter_col and filter_col != "Не фильтровать":
                    with open(file_path, "r", encoding=encoding, errors="replace") as f:
                        reader = csv.reader(f, delimiter=delimiter)
                        headers = next(reader)
                        try:
                            filter_idx = headers.index(filter_col)
                            filter_vals = set(
                                self.window._filter_values.get(filter_col, [])
                            )
                        except ValueError:
                            filter_idx = None

                total += FileUtilities.count_rows(
                    file_path, delimiter, encoding, filter_idx, filter_vals
                )
            except (OSError, IOError, UnicodeDecodeError) as e:
                self._log_message(f"Ошибка подсчёта строк: {e}", QColor("red"))

        self.window.total_rows_label.setText(f"Строк: {total}")

    def _on_split_column_selected(self, index: int):
        """Обработчик выбора столбца для разделения."""
        if index <= 0:
            return

        column = self.window.split_column_combo.currentText()
        if not self.window.file_list.count():
            return

        try:
            # Получаем текущий фильтр
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
            )

            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected = dialog.get_selected_values()
                self.window._selected_column_values[column] = selected
                self._log_message(
                    f"Выбраны значения для разделения: {', '.join(selected[:5])}...",
                    QColor("green"),
                )
            else:
                self.window.split_column_combo.setCurrentIndex(0)

        except Exception as e:
            self._log_message(f"Ошибка: {e}", QColor("red"))
            self.window.split_column_combo.setCurrentIndex(0)

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
            split_column=config.split_column,
            selected_values=config.selected_values,
            filter_column=config.filter_column,
            filter_values=config.filter_values,
            pivot_settings=config.pivot_settings,
        )

        # Подключаем сигналы
        self.converter.update_progress.connect(self.window.progress_bar.setValue)
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
        """Удаляет сконвертированный файл."""
        if self.converter and hasattr(self.converter, "output_file_path"):
            path = self.converter.output_file_path
            if path and os.path.exists(path):
                msgbox = self._show_message_box(
                    QMessageBox.Icon.Question,
                    "Подтверждение",
                    "Удалить файл?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                reply = msgbox.exec()
                if reply == QMessageBox.StandardButton.Yes:
                    os.remove(path)
                    self.window.open_file_btn.setEnabled(False)
                    self.window.delete_file_btn.setEnabled(False)
                    self._log_message("Файл удалён", QColor("red"))

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
            <tr><td>Разделение по столбцу</td><td>{self.window.split_column_combo.currentText()}</td></tr>
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

    # Устанавливаем стиль Fusion для кроссплатформенной консистентности
    # и лучшего контроля над палитрой цветов
    app.setStyle(QStyleFactory.create("Fusion"))

    # Запуск приложения
    converter_app = TSVConverterApp()
    converter_app.run()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
