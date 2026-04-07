"""
Модуль GUI для конвертера TSV/CSV в Excel.
Использует только нативные возможности PySide6 без QSS.
Архитектура: модульная, с разделением на компоненты.
"""

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QToolButton,
    QLineEdit,
    QTextEdit,
    QProgressBar,
    QFileDialog,
    QMessageBox,
    QListWidget,
    QListWidgetItem,
    QComboBox,
    QCheckBox,
    QSpinBox,
    QFontComboBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QAbstractItemView,
    QDialog,
    QColorDialog,
    QSizePolicy,
    QGroupBox,
    QFrame,
    QStyledItemDelegate,
)
from PySide6.QtCore import (
    Qt,
    Signal,
    QTimer,
    QThread,
    QRect,
    Slot,
    QObject,
    QEvent,
    QSize,
)
from PySide6.QtGui import (
    QColor,
    QFont,
    QPalette,
    QDragEnterEvent,
    QDropEvent,
    QPainter,
    QPen,
    QBrush,
)

import os
import csv
from datetime import datetime
from typing import Optional, Dict, List, Set, Any

# Импортируем утилиты из converter.py
from converter import FileUtilities


# ============================================================================
# УТИЛИТЫ И БАЗОВЫЕ КЛАССЫ
# ============================================================================


class StylePalette:
    """
    Централизованное хранилище цветов и стилей.
    Позволяет легко менять тему приложения.
    """

    # Светлая тема (корпоративная) - улучшенные контрасты
    LIGHT_THEME = {
        "window_bg": QColor(240, 240, 240),
        "panel_bg": QColor(255, 255, 255),
        "text_primary": QColor(0, 0, 0),
        "text_secondary": QColor(60, 60, 60),
        "accent_primary": QColor(0, 100, 200),
        "accent_success": QColor(30, 130, 60),
        "accent_warning": QColor(180, 140, 20),
        "accent_danger": QColor(180, 40, 50),
        "border": QColor(0, 0, 0),
        "border_light": QColor(0, 0, 0),
        "hover": QColor(230, 230, 230),
        "selected": QColor(210, 220, 235),
        "shadow": QColor(0, 0, 0),
    }

    # Тёмная тема - улучшенные контрасты
    DARK_THEME = {
        "window_bg": QColor(30, 30, 30),
        "panel_bg": QColor(45, 45, 45),
        "text_primary": QColor(250, 250, 250),
        "text_secondary": QColor(180, 180, 180),
        "accent_primary": QColor(80, 140, 220),
        "accent_success": QColor(50, 150, 70),
        "accent_warning": QColor(210, 160, 50),
        "accent_danger": QColor(200, 60, 70),
        "border": QColor(90, 90, 90),
        "border_light": QColor(70, 70, 70),
        "hover": QColor(60, 60, 60),
        "selected": QColor(55, 65, 80),
        "shadow": QColor(0, 0, 0),
    }


def apply_native_style(
    widget: QWidget, palette: Dict[str, QColor], is_dark: bool = False
):
    """
    Применяет нативную палитру к виджету без использования QSS.
    Рекурсивно применяет тему ко всем дочерним виджетам.
    """
    pal = widget.palette()

    # Основные цвета
    pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
    pal.setColor(QPalette.ColorRole.AlternateBase, palette["hover"])
    pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])

    # Кнопки - используем panel_bg для фона, чтобы соответствовал окну
    pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Highlight, palette["accent_primary"])
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))

    # Цвета для состояний кнопок
    pal.setColor(QPalette.ColorRole.Light, palette["hover"])
    pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])

    pal.setColor(QPalette.ColorRole.Dark, palette["border"])
    pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
    pal.setColor(QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0)))

    widget.setPalette(pal)
    widget.setAutoFillBackground(True)

    # Применяем тему ко всем дочерним виджетам
    for child in widget.findChildren(QWidget):
        # Применяем тему к StyledCheckBox
        if isinstance(child, StyledCheckBox):
            child.set_theme(is_dark)

        # Для QListWidget и QTableWidget не используем QSS для индикаторов,
        # так как Qt6 не поддерживает SVG в QSS. Используем стандартную отрисовку.
        if isinstance(child, (QListWidget, QTableWidget)):
            # Просто сбрасываем QSS чтобы не конфликтовал
            child.setStyleSheet("")

        child_pal = child.palette()
        child_pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
        child_pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
        child_pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
        child_pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])
        # Кнопки и интерактивные элементы
        child_pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
        child_pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
        child_pal.setColor(QPalette.ColorRole.Light, palette["hover"])
        child_pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])
        # Границы
        child_pal.setColor(QPalette.ColorRole.Dark, palette["border"])
        child_pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
        child_pal.setColor(
            QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0))
        )
        child.setPalette(child_pal)
        child.setAutoFillBackground(True)


def create_separator_line(
    orientation: Qt.Orientation = Qt.Orientation.Horizontal,
) -> QFrame:
    """Создаёт линию-разделитель."""
    line = QFrame()
    line.setFrameShape(
        QFrame.Shape.HLine
        if orientation == Qt.Orientation.Horizontal
        else QFrame.Shape.VLine
    )
    line.setFrameShadow(QFrame.Shadow.Sunken)
    return line


# ============================================================================
# КАСТОМНЫЕ ВИДЖЕТЫ
# ============================================================================


class StyledCheckBox(QCheckBox):
    """
    Кастомный QCheckBox с корректной отрисовкой галочки.
    Использует paintEvent для ручной отрисовки индикатора.
    """

    def __init__(self, text: str = "", parent=None):
        super().__init__(text, parent)
        self._is_dark = False
        # Включаем hover эффект
        self.setAttribute(Qt.WidgetAttribute.WA_Hover)
        # Устанавливаем минимальную высоту для корректного отображения
        self.setMinimumHeight(24)
        # Отключаем стандартную отрисовку Windows 11 для предотвращения
        # переопределения цветов системной темой
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent)

    def set_theme(self, is_dark: bool):
        """Устанавливает тему для отрисовки."""
        self._is_dark = is_dark
        self.update()

    def paintEvent(self, event):
        """Переопределённая отрисовка для корректного отображения галочки."""
        from PySide6.QtGui import QPainter, QPen, QBrush
        from PySide6.QtCore import QRect, Qt

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Очищаем область перед отрисовкой (требуется для WA_OpaquePaintEvent)
        painter.fillRect(self.rect(), self.palette().color(QPalette.ColorRole.Window))

        # Определяем цвета в зависимости от темы
        if self._is_dark:
            border_color = QColor(144, 144, 144)
            bg_color = QColor(80, 80, 80)
            hover_color = QColor(96, 96, 96)
            check_color = QColor(120, 220, 120)
            text_color = (
                QColor(230, 230, 230) if self.isEnabled() else QColor(100, 100, 100)
            )
        else:
            border_color = QColor(0, 0, 0)
            bg_color = QColor(255, 255, 255)
            hover_color = QColor(224, 224, 224)
            check_color = QColor(35, 120, 55)
            text_color = QColor(0, 0, 0) if self.isEnabled() else QColor(150, 150, 150)

        # Проверяем состояние
        is_checked = self.isChecked()
        is_hovered = self.underMouse()
        is_enabled = self.isEnabled()
        is_down = self.isDown()

        # Получаем размеры
        indicator_size = 16
        indicator_rect = QRect(
            0, (self.height() - indicator_size) // 2, indicator_size, indicator_size
        )

        # Рисуем фон индикатора
        if is_down:
            painter.setBrush(QBrush(hover_color))
        elif is_hovered and is_enabled:
            painter.setBrush(QBrush(hover_color))
        else:
            painter.setBrush(QBrush(bg_color))

        painter.setPen(QPen(border_color, 1))
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.drawRoundedRect(indicator_rect.adjusted(0, 0, -1, -1), 3, 3)

        # Рисуем галочку если checked
        if is_checked:
            painter.setPen(QPen(check_color, 2))
            painter.setBrush(Qt.BrushStyle.NoBrush)

            # Координаты для галочки
            check_rect = indicator_rect.adjusted(4, 4, -4, -4)
            cx, cy = check_rect.center().x(), check_rect.center().y()

            # Рисуем галочку (две линии)
            painter.drawLine(int(cx - 3), int(cy), int(cx - 1), int(cy + 2))
            painter.drawLine(int(cx - 1), int(cy + 2), int(cx + 3), int(cy - 2))

        # Рисуем текст
        if self.text():
            painter.setPen(QPen(text_color))
            font = painter.font()
            font.setPointSize(9)
            painter.setFont(font)

            text_rect = QRect(
                indicator_size + 6, 0, self.width() - indicator_size - 6, self.height()
            )
            painter.drawText(
                text_rect,
                Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                self.text(),
            )


class StyledButton(QPushButton):
    """
    Кнопка с нативным стилем и поддержкой тем.
    """

    def __init__(self, text: str, parent=None, button_type: str = "default"):
        super().__init__(text, parent)
        self.button_type = button_type
        self._apply_style()

    def _apply_style(self):
        """Применяет стиль в зависимости от типа кнопки."""
        # Устанавливаем размерный шрифт
        font = self.font()
        font.setPointSize(10)
        self.setFont(font)

        # Устанавливаем политику размера
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.setMinimumHeight(32)


class PrimaryButton(StyledButton):
    """Основная кнопка действия (например, 'Начать')."""

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent, "primary")
        self.setAutoDefault(True)


class SecondaryButton(StyledButton):
    """Второстепенная кнопка."""

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent, "secondary")


class DangerButton(StyledButton):
    """Кнопка опасного действия (удаление, закрытие)."""

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent, "danger")


class FileListWidget(QListWidget):
    """
    Виджет списка файлов с поддержкой Drag & Drop.
    """

    files_dropped = Signal(object)  # list of file paths

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)

        # Настройка внешнего вида
        self.setMinimumHeight(110)
        self.setMaximumHeight(300)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        font = self.font()
        font.setPointSize(9)
        self.setFont(font)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile().lower()
                if file_path.endswith((".tsv", ".csv", ".txt")):
                    event.accept()
                    return
        event.ignore()

    def dragMoveEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile().lower()
                if file_path.endswith((".tsv", ".csv", ".txt")):
                    event.accept()
                    event.setDropAction(Qt.DropAction.CopyAction)
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".tsv", ".csv", ".txt")):
                files.append(file_path)
        if files:
            self.files_dropped.emit(list(files))
            event.acceptProposedAction()


class LogTextEdit(QTextEdit):
    """
    Виджет лога с автоматической прокруткой.
    """

    # Спокойные цвета для логов - СВЕТЛАЯ ТЕМА
    LIGHT_LOG_COLORS = {
        "info": QColor(50, 50, 50),  # Тёмно-серый для информации
        "success": QColor(35, 120, 55),  # Спокойный зелёный
        "warning": QColor(150, 110, 20),  # Спокойный янтарный
        "error": QColor(160, 40, 50),  # Спокойный красный
        "blue": QColor(40, 80, 140),  # Спокойный синий
        "orange": QColor(160, 90, 30),  # Спокойный оранжевый
    }

    # Спокойные цвета для логов - ТЁМНАЯ ТЕМА (контрастные, хорошо читаемые)
    DARK_LOG_COLORS = {
        "info": QColor(230, 230, 230),  # Светло-серый для информации
        "success": QColor(120, 220, 120),  # Светлый зелёный
        "warning": QColor(240, 200, 100),  # Светлый янтарный
        "error": QColor(240, 120, 120),  # Светлый красный
        "blue": QColor(160, 200, 255),  # Очень светлый синий (отлично читается)
        "orange": QColor(240, 170, 100),  # Светлый оранжевый
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Моноширинный шрифт для логов
        font = QFont("Consolas", 9)
        self.setFont(font)

        # Определяем тему от родителя
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        # Выбираем палитру в зависимости от темы
        self._log_colors = (
            self.DARK_LOG_COLORS if self._is_dark else self.LIGHT_LOG_COLORS
        )

        # Цвет текста по умолчанию
        self.setTextColor(self._log_colors["info"])

        # Минимальная высота
        self.setMinimumHeight(150)

    def append_colored(self, message: str, color: QColor):
        """Добавляет цветное сообщение."""
        timestamp = datetime.now().strftime("%H:%M:%S")

        # Преобразуем яркие цвета в спокойные (с учётом темы)
        log_color = self._get_calm_color(color)

        html = (
            f'<span style="color: {log_color.name()};">[{timestamp}] {message}</span>'
        )
        self.append(html)
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())

    def _get_calm_color(self, color: QColor) -> QColor:
        """Преобразует яркие цвета в спокойные аналоги с учётом темы."""
        color_name = color.name().lower()
        colors = self._log_colors

        # Ярко-зелёный -> спокойный зелёный
        if color_name in ["#28a745", "#40a640", "#388e3c", "#4caf50", "#45a049"]:
            return colors["success"]

        # Ярко-красный -> спокойный красный
        if color_name in ["#dc3545", "#d32f2f", "#c9302c", "#f44336", "#d9534f"]:
            return colors["error"]

        # Ярко-синий -> спокойный синий
        if color_name in ["#007bff", "#1976d2", "#0066cc", "#2196f3", "#0078d7"]:
            return colors["blue"]

        # Ярко-оранжевый/жёлтый -> спокойный оранжевый
        if color_name in ["#ffc107", "#ff9800", "#f57c00", "#ffeb3b", "#ff9800"]:
            return colors["warning"]

        # Чёрный/тёмный -> цвет информации
        if color_name in ["#000000", "#212529", "#333333"]:
            return colors["info"]

        # Остальные цвета возвращаем как есть
        return color

    def set_theme(self, is_dark: bool):
        """Обновляет палитру при смене темы."""
        self._is_dark = is_dark
        self._log_colors = self.DARK_LOG_COLORS if is_dark else self.LIGHT_LOG_COLORS
        self.setTextColor(self._log_colors["info"])


# ============================================================================
# ДИАЛОГОВЫЕ ОКНА
# ============================================================================


def apply_theme_to_dialog(dialog: QDialog, is_dark: bool):
    """
    Применяет тему к диалоговому окну и всем его виджетам.

    Args:
        dialog: Диалоговое окно
        is_dark: True для тёмной темы, False для светлой
    """
    palette = StylePalette.DARK_THEME if is_dark else StylePalette.LIGHT_THEME
    pal = dialog.palette()

    # Основные цвета диалога
    pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
    pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])
    # Кнопки - используем window_bg для соответствия цвету окна
    pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Dark, palette["border"])
    pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
    pal.setColor(QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0)))
    pal.setColor(QPalette.ColorRole.Light, palette["hover"])
    pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])

    dialog.setPalette(pal)
    dialog.setAutoFillBackground(True)

    # Применяем ко всем дочерним виджетам
    for widget in dialog.findChildren(QWidget):
        # Применяем тему к StyledCheckBox
        if isinstance(widget, StyledCheckBox):
            widget.set_theme(is_dark)

        # Применяем тему к CheckBoxListWidget (для чекбоксов в списках)
        if isinstance(widget, CheckBoxListWidget):
            widget.set_theme(is_dark)

        # Для QListWidget и QTableWidget не используем QSS для индикаторов,
        # так как Qt6 не поддерживает SVG в QSS. Используем стандартную отрисовку.
        if isinstance(widget, (QListWidget, QTableWidget)):
            widget.setStyleSheet("")

        widget_pal = widget.palette()
        widget_pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
        widget_pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
        widget_pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
        widget_pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])
        # Кнопки - используем window_bg для соответствия цвету окна
        widget_pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
        widget_pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
        widget_pal.setColor(QPalette.ColorRole.Light, palette["hover"])
        widget_pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])
        widget_pal.setColor(QPalette.ColorRole.Dark, palette["border"])
        widget_pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
        widget_pal.setColor(
            QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0))
        )
        widget.setPalette(widget_pal)
        widget.setAutoFillBackground(True)


def apply_theme_to_messagebox(msgbox: QMessageBox, is_dark: bool):
    """
    Применяет тему к QMessageBox.

    Args:
        msgbox: QMessageBox
        is_dark: True для тёмной темы, False для светлой
    """
    palette = StylePalette.DARK_THEME if is_dark else StylePalette.LIGHT_THEME
    pal = msgbox.palette()

    # Основные цвета
    pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
    pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
    pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
    pal.setColor(QPalette.ColorRole.Highlight, palette["accent_primary"])
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    pal.setColor(QPalette.ColorRole.Dark, palette["border"])
    pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
    pal.setColor(QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0)))
    pal.setColor(QPalette.ColorRole.Light, palette["hover"])
    pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])

    msgbox.setPalette(pal)
    msgbox.setAutoFillBackground(True)

    # Применяем ко всем дочерним виджетам
    for widget in msgbox.findChildren(QWidget):
        widget_pal = widget.palette()
        widget_pal.setColor(QPalette.ColorRole.Window, palette["window_bg"])
        widget_pal.setColor(QPalette.ColorRole.WindowText, palette["text_primary"])
        widget_pal.setColor(QPalette.ColorRole.Base, palette["panel_bg"])
        widget_pal.setColor(QPalette.ColorRole.Text, palette["text_primary"])
        widget_pal.setColor(QPalette.ColorRole.Button, palette["window_bg"])
        widget_pal.setColor(QPalette.ColorRole.ButtonText, palette["text_primary"])
        widget_pal.setColor(QPalette.ColorRole.Light, palette["hover"])
        widget_pal.setColor(QPalette.ColorRole.Midlight, palette["border_light"])
        widget_pal.setColor(QPalette.ColorRole.Dark, palette["border"])
        widget_pal.setColor(QPalette.ColorRole.Mid, palette["border_light"])
        widget_pal.setColor(
            QPalette.ColorRole.Shadow, palette.get("shadow", QColor(0, 0, 0))
        )
        widget.setPalette(widget_pal)
        widget.setAutoFillBackground(True)


def get_widget_theme_flag(widget: QWidget) -> bool:
    """Безопасно возвращает флаг тёмной темы для окна или диалога."""
    if hasattr(widget, "_is_dark"):
        return bool(getattr(widget, "_is_dark"))
    if hasattr(widget, "_is_dark_theme"):
        return bool(getattr(widget, "_is_dark_theme"))
    return False


# ============================================================================
# LOADING OVERLAY
# ============================================================================


class LoadingOverlay(QWidget):
    """
    Оверлей загрузки с анимированным спиннером.
    Отображается по центру родительского виджета, блокируя взаимодействие с GUI.
    Имеет полупрозрачный фон для видимости содержимого под ним.
    """

    def __init__(self, parent: QWidget, text: str = "Загрузка..."):
        super().__init__(parent)
        self._text = text
        self._angle = 0
        self._rotation_speed = 0

        # Настройка виджета
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)

        # Таймер анимации
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._update_animation)

        # Скрываем по умолчанию
        self.setVisible(False)

    def set_text(self, text: str):
        """Обновляет текст загрузки."""
        self._text = text
        self.update()

    def start_animation(self):
        """Запускает анимацию спиннера."""
        self._timer.start(16)  # ~60 FPS
        self._rotation_speed = 20
        self.setVisible(True)
        self.raise_()  # Поверх всех элементов

    def stop_animation(self):
        """Останавливает анимацию и скрывает оверлей."""
        self._timer.stop()
        self._rotation_speed = 0
        self.setVisible(False)

    def _update_animation(self):
        """Обновляет угол вращения спиннера."""
        self._angle = (self._angle + self._rotation_speed) % 360
        self.update()

    def paintEvent(self, event):
        """Отрисовка оверлея и спиннера."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Получаем тему от родителя
        is_dark = False
        if self.parent() and hasattr(self.parent(), "_is_dark_theme"):
            is_dark = self.parent()._is_dark_theme

        # Цвета в зависимости от темы (полупрозрачный фон alpha=150 для видимости GUI)
        if is_dark:
            overlay_color = QColor(30, 30, 30, 150)  # Более прозрачный
            spinner_color = QColor(100, 180, 255)
            text_color = QColor(230, 230, 230)
        else:
            overlay_color = QColor(240, 240, 240, 150)  # Более прозрачный
            spinner_color = QColor(0, 100, 200)
            text_color = QColor(20, 20, 20)

        # Полупрозрачный фон на весь родительский виджет
        if self.parent():
            parent_rect = self.parent().rect()
            painter.fillRect(parent_rect, QBrush(overlay_color))

        # Центрируем спиннер и текст
        center_x = self.width() // 2
        center_y = self.height() // 2

        # Рисуем спиннер (круг из сегментов)
        self._draw_spinner(painter, center_x, center_y, spinner_color)

        # Рисуем текст под спиннером
        painter.setPen(QPen(text_color))
        font = painter.font()
        font.setPointSize(12)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(
            QRect(center_x - 100, center_y + 60, 200, 30),
            Qt.AlignmentFlag.AlignCenter,
            self._text,
        )

    def _draw_spinner(
        self, painter: QPainter, center_x: int, center_y: int, color: QColor
    ):
        """Рисует анимированный спиннер."""
        radius = 40
        pen_width = 6

        # Создаём перо
        pen = QPen(color, pen_width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap)
        painter.setPen(pen)

        # Рисуем дугу (270 градусов)
        rect = QRect(center_x - radius, center_y - radius, radius * 2, radius * 2)

        # Поворачиваем на текущий угол
        painter.save()
        painter.translate(center_x, center_y)
        painter.rotate(self._angle)
        painter.translate(-center_x, -center_y)

        # Рисуем дугу от -135 до +135 градусов (270 градусов всего)
        painter.drawArc(rect, -135 * 16, 270 * 16)

        painter.restore()

    def resizeEvent(self, event):
        """При изменении размера перерисовываем."""
        super().resizeEvent(event)
        self.update()


class LoadingWorker(QThread):
    """
    Поток для загрузки и обработки файлов.
    Используется для отображения лоадера во время длительных операций.
    """

    finished = Signal(object)  # Результат обработки
    error = Signal(str)  # Ошибка
    progress = Signal(int, str)  # Прогресс (шаг, текст)

    def __init__(self, operation_callable, *args, **kwargs):
        super().__init__()
        self.operation = operation_callable
        self.args = args
        self.kwargs = kwargs

    def run(self):
        """Выполняет операцию в отдельном потоке."""
        try:
            result = self.operation(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class UniqueValuesWorkerSignals(QObject):
    """Класс сигналов для UniqueValuesWorker."""

    finished = Signal(object)  # set уникальных значений
    error = Signal(str)


class UniqueValuesWorker(QThread):
    """
    Поток для получения уникальных значений столбца.
    Используется в ColumnValuesDialog для загрузки значений без блокировки UI.
    """

    def __init__(
        self,
        file_paths: List[str],
        column: str,
        filter_column: str = "",
        filter_values: Optional[List[str]] = None,
    ):
        super().__init__()
        self.file_paths = file_paths
        self.column = column
        self.filter_column = filter_column
        self.filter_values = filter_values or []

        # Создаём сигналы через отдельный класс
        self.signals = UniqueValuesWorkerSignals()

    def run(self):
        """Получает уникальные значения столбца."""
        try:
            values = set()

            for file_path in self.file_paths:
                encoding = FileUtilities.get_encoding(file_path)
                delimiter = FileUtilities.get_delimiter(file_path)

                with open(file_path, "r", encoding=encoding, errors="replace") as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    headers = next(reader)

                    try:
                        col_idx = headers.index(self.column)
                        filter_idx = None

                        # Индекс фильтра
                        if (
                            self.filter_column
                            and self.filter_column != "Не фильтровать"
                        ):
                            try:
                                filter_idx = headers.index(self.filter_column)
                            except ValueError:
                                filter_idx = None
                                filter_vals = set()
                            else:
                                filter_vals = set(self.filter_values)

                        # Собираем уникальные значения
                        for row in reader:
                            if col_idx < len(row):
                                # Применяем фильтр если есть
                                if filter_idx is not None and filter_vals:
                                    if filter_idx < len(row):
                                        if row[filter_idx] not in filter_vals:
                                            continue

                                values.add(row[col_idx])

                    except ValueError:
                        continue

            # Emit через set
            self.signals.finished.emit(values)

        except Exception as e:
            self.signals.error.emit(str(e))


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setMinimumSize(500, 220)
        self.setModal(True)

        # Применяем тему родителя если есть
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        self._init_ui()
        # Применяем тему после создания UI
        apply_theme_to_dialog(self, self._is_dark)

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # Группа: Оформление
        appearance_group = self._create_appearance_group()
        layout.addWidget(appearance_group)

        # Группа: Пути и поведение
        behavior_group = self._create_behavior_group()
        layout.addWidget(behavior_group)

        # Прижимаем элементы к верху
        layout.addStretch(1)

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        save_btn = PrimaryButton("Сохранить")
        save_btn.clicked.connect(self.accept)
        button_layout.addWidget(save_btn)

        cancel_btn = SecondaryButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def _create_appearance_group(self) -> QGroupBox:
        group = QGroupBox("Оформление")
        layout = QHBoxLayout()

        layout.addWidget(QLabel("Тема оформления:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Светлая", "Тёмная"])
        layout.addWidget(self.theme_combo)

        # Растягиваем комбобокс, но не даем ему быть бесконечным, если нужно
        # В данном случае QHBoxLayout сам распределит место, или можно добавить stretch

        group.setLayout(layout)
        return group

    def _create_behavior_group(self) -> QGroupBox:
        group = QGroupBox("Пути и поведение")
        layout = QVBoxLayout()

        # Путь сохранения
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("Путь по умолчанию:"))
        self.default_path_edit = QLineEdit()
        self.default_path_edit.setPlaceholderText("Выберите папку...")
        path_layout.addWidget(self.default_path_edit)

        browse_btn = SecondaryButton("Обзор...")
        browse_btn.clicked.connect(self._browse_path)
        path_layout.addWidget(browse_btn)

        layout.addLayout(path_layout)

        # Чекбоксы
        self.auto_open_checkbox = StyledCheckBox("Автооткрытие файла после конвертации")
        layout.addWidget(self.auto_open_checkbox)

        self.auto_delete_checkbox = StyledCheckBox("Автоудаление исходного файла")
        layout.addWidget(self.auto_delete_checkbox)

        group.setLayout(layout)
        return group

    def _browse_path(self):
        directory = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if directory:
            self.default_path_edit.setText(directory)

    def get_settings(self) -> Dict[str, Any]:
        return {
            "theme": self.theme_combo.currentText(),
            "default_path": self.default_path_edit.text(),
            "auto_open": self.auto_open_checkbox.isChecked(),
            "auto_delete": self.auto_delete_checkbox.isChecked(),
        }

    def load_settings(self, settings: Dict[str, Any]):
        self.theme_combo.setCurrentText(settings.get("theme", "Светлая"))
        self.default_path_edit.setText(settings.get("default_path", ""))
        self.auto_open_checkbox.setChecked(settings.get("auto_open", False))
        self.auto_delete_checkbox.setChecked(settings.get("auto_delete", False))


class CheckBoxListWidget(QListWidget):
    """
    Кастомный QListWidget для чекбоксов.
    Использует CheckBoxListDelegate для корректной отрисовки с учётом темы.
    Хранит состояние чекбоксов независимо для предотвращения двойной обработки Qt.
    """

    def __init__(self, parent=None, is_dark: bool = False):
        super().__init__(parent)
        self._is_dark = is_dark
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setMouseTracking(True)

        # Устанавливаем кастомный delegate для отрисовки чекбоксов
        # Это предотвращает переопределение цветов системной темой Windows 11
        self._delegate = CheckBoxListDelegate(self, is_dark)
        self.setItemDelegate(self._delegate)

        # Хранилище состояний чекбоксов: row -> Qt.CheckState
        self._check_states: Dict[int, Qt.CheckState] = {}

        # Callback для уведомления об изменении состояния чекбокса
        self._on_check_state_changed = None

    def set_theme(self, is_dark: bool):
        """Обновляет тему для delegate."""
        self._is_dark = is_dark
        self._delegate.set_theme(is_dark)
        self.viewport().update()

    def set_item_check_state(self, row: int, state: Qt.CheckState):
        """Устанавливает состояние чекбокса для элемента."""
        self._check_states[row] = state
        self.viewport().update()
        if self._on_check_state_changed:
            self._on_check_state_changed()

    def get_item_check_state(self, row: int) -> Qt.CheckState:
        """Получает состояние чекбокса для элемента."""
        return self._check_states.get(row, Qt.CheckState.Unchecked)


class CheckBoxListDelegate(QStyledItemDelegate):
    """
    Кастомный delegate для отрисовки чекбоксов в QListWidget.
    Обеспечивает корректную отрисовку галочек в списках.
    """

    def __init__(self, parent=None, is_dark: bool = False):
        super().__init__(parent)
        self._is_dark = is_dark
        self._parent_widget = parent  # Сохраняем ссылку на родительский виджет

    def set_theme(self, is_dark: bool):
        """Устанавливает тему."""
        self._is_dark = is_dark

    def paint(self, painter, option, index):
        """Отрисовка элемента списка с чекбоксом."""
        from PySide6.QtGui import QPen, QBrush
        from PySide6.QtWidgets import QStyle

        painter.save()

        # Определяем цвета
        if self._is_dark:
            border_color = QColor(144, 144, 144)
            bg_color = QColor(80, 80, 80)
            hover_color = QColor(96, 96, 96)
            check_color = QColor(120, 220, 120)
            text_color = QColor(230, 230, 230)
        else:
            border_color = QColor(0, 0, 0)
            bg_color = QColor(255, 255, 255)
            hover_color = QColor(224, 224, 224)
            check_color = QColor(35, 120, 55)
            text_color = QColor(0, 0, 0)

        # Получаем состояние из хранилища виджета
        row = index.row()
        is_checked = False
        if self._parent_widget:
            is_checked = (
                self._parent_widget.get_item_check_state(row) == Qt.CheckState.Checked
            )

        is_hovered = option.state & QStyle.StateFlag.State_MouseOver
        is_selected = option.state & QStyle.StateFlag.State_Selected

        # Рисуем фон выделения
        if is_selected:
            painter.fillRect(
                option.rect,
                QBrush(
                    QColor(200, 220, 240) if not self._is_dark else QColor(60, 70, 90)
                ),
            )

        # Рисуем индикатор чекбокса
        indicator_size = 16
        indicator_rect = QRect(
            option.rect.left() + 4,
            option.rect.center().y() - indicator_size // 2,
            indicator_size,
            indicator_size,
        )

        # Фон индикатора
        if is_hovered:
            painter.setBrush(QBrush(hover_color))
        else:
            painter.setBrush(QBrush(bg_color))

        painter.setPen(QPen(border_color, 1))
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.drawRoundedRect(indicator_rect.adjusted(0, 0, -1, -1), 3, 3)

        # Галочка если checked
        if is_checked:
            painter.setPen(QPen(check_color, 2))
            painter.setBrush(Qt.BrushStyle.NoBrush)

            check_rect = indicator_rect.adjusted(4, 4, -4, -4)
            cx, cy = check_rect.center().x(), check_rect.center().y()

            painter.drawLine(int(cx - 3), int(cy), int(cx - 1), int(cy + 2))
            painter.drawLine(int(cx - 1), int(cy + 2), int(cx + 3), int(cy - 2))

        # Текст
        text_rect = QRect(
            indicator_rect.right() + 6,
            option.rect.top(),
            option.rect.width() - indicator_size - 10,
            option.rect.height(),
        )
        painter.setPen(QPen(text_color))
        font = painter.font()
        font.setPointSize(9)
        painter.setFont(font)
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            index.data(Qt.ItemDataRole.DisplayRole),
        )

        painter.restore()

    def sizeHint(self, option, index):
        """Размер элемента."""
        return QSize(200, 28)

    def editorEvent(self, event, model, option, index):
        """Обработка событий (клики)."""
        if event.type() == QEvent.Type.MouseButtonPress:
            mouse_pos = event.position().toPoint()
            indicator_size = 16
            indicator_rect = QRect(
                option.rect.left() + 4,
                option.rect.center().y() - indicator_size // 2,
                indicator_size,
                indicator_size,
            )

            if indicator_rect.contains(mouse_pos):
                # Получаем текущее состояние из виджета
                row = index.row()
                current = self._parent_widget.get_item_check_state(row)
                new_state = (
                    Qt.CheckState.Unchecked
                    if current == Qt.CheckState.Checked
                    else Qt.CheckState.Checked
                )
                # Обновляем состояние в виджете
                self._parent_widget.set_item_check_state(row, new_state)
                # Обновляем стандартную модель для совместимости
                model.setData(index, new_state, Qt.ItemDataRole.CheckStateRole)
                event.accept()
                return True

        # Для других событий — стандартная обработка
        return super().editorEvent(event, model, option, index)


class ColumnValuesDialog(QDialog):
    """
    Диалог выбора значений столбца.
    Поддерживает как синхронную загрузку значений, так и асинхронную через worker.
    """

    def __init__(
        self,
        values: Optional[Set[str]] = None,
        parent=None,
        file_paths: Optional[List[str]] = None,
        column: str = "",
        filter_column: str = "",
        filter_values: Optional[List[str]] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Выберите значения")
        self.setMinimumSize(400, 500)
        self.setModal(True)

        # Применяем тему родителя если есть
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        self._values = sorted(values) if values else []
        self._file_paths = file_paths or []
        self._column = column
        self._filter_column = filter_column
        self._filter_values = filter_values or []
        self._worker = None

        # Сначала инициализируем UI
        self._init_ui()
        apply_theme_to_dialog(self, self._is_dark)

        # Создаём лоадер
        self._loading_overlay = LoadingOverlay(self, "Загрузка...")

        # Если переданы file_paths и column, загружаем значения асинхронно
        if self._file_paths and self._column:
            self._load_values_async()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Поиск
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск...")
        self.search_edit.textChanged.connect(self._on_search_text_changed)
        layout.addWidget(self.search_edit)

        # Список с кастомным виджетом для чекбоксов
        # Передаём тему для корректной отрисовки чекбоксов в Windows 11
        self.value_list = CheckBoxListWidget(parent=self, is_dark=self._is_dark)
        self.value_list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.value_list._on_check_state_changed = self._update_info

        # Добавляем элементы с чекбоксами
        for value in self._values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.value_list.addItem(item)

        layout.addWidget(self.value_list)

        # Инфо
        self.info_label = QLabel("Выбрано: 0")
        layout.addWidget(self.info_label)

        # Кнопки
        button_layout = QHBoxLayout()

        select_all_btn = SecondaryButton("Выбрать все")
        select_all_btn.clicked.connect(self._select_all)
        button_layout.addWidget(select_all_btn)

        clear_all_btn = SecondaryButton("Очистить все")
        clear_all_btn.clicked.connect(self._clear_all)
        button_layout.addWidget(clear_all_btn)

        button_layout.addStretch()

        ok_btn = PrimaryButton("OK")
        ok_btn.clicked.connect(self.accept)
        button_layout.addWidget(ok_btn)

        layout.addLayout(button_layout)

        self._update_info()

    def _load_values_async(self):
        """Загружает значения асинхронно через worker."""
        self._show_loading()

        self._worker = UniqueValuesWorker(
            self._file_paths, self._column, self._filter_column, self._filter_values
        )
        self._worker.signals.finished.connect(self._on_values_loaded)
        self._worker.signals.error.connect(self._on_values_error)
        self._worker.start()

    @Slot(object)
    def _on_values_loaded(self, values):
        """Обработчик загрузки значений."""
        self._hide_loading()
        # Загружаем значения в список
        if isinstance(values, set):
            self.load_values(values)
        self._worker = None

    @Slot(str)
    def _on_values_error(self, error: str):
        """Обработчик ошибки загрузки."""
        self._hide_loading()
        self._worker = None

        msgbox = QMessageBox(
            QMessageBox.Icon.Critical,
            "Ошибка",
            f"Ошибка при загрузке значений:\n{error}",
            self,
        )
        apply_theme_to_messagebox(msgbox, self._is_dark)
        msgbox.exec()

    def _show_loading(self):
        """Показывает лоадер."""
        self._loading_overlay.resize(self.size())
        self._loading_overlay.start_animation()
        self.value_list.setEnabled(False)
        self.search_edit.setEnabled(False)

    def _hide_loading(self):
        """Скрывает лоадер."""
        self._loading_overlay.stop_animation()
        self.value_list.setEnabled(True)
        self.search_edit.setEnabled(True)

    def load_values(self, values: Set[str]):
        """Загружает значения в список."""
        self._values = sorted(values)

        # Очищаем список
        self.value_list.clear()
        self.value_list._check_states.clear()  # Сбрасываем состояния

        # Добавляем элементы с чекбоксами
        for i, value in enumerate(self._values):
            item = QListWidgetItem(value)
            # Убираем ItemIsUserCheckable чтобы Qt не обрабатывал клики сам
            flags = item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable
            item.setFlags(flags)
            self.value_list.addItem(item)
            # Инициализируем состояние как Unchecked
            self.value_list._check_states[i] = Qt.CheckState.Unchecked

        self._update_info()

    def _on_search_text_changed(self, text: str):
        """Обработчик изменения текста поиска."""
        for i in range(self.value_list.count()):
            item = self.value_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())
        self._update_info()

    def resizeEvent(self, event):
        """Обновляет размер лоадера при изменении размера диалога."""
        super().resizeEvent(event)
        if hasattr(self, "_loading_overlay"):
            self._loading_overlay.resize(self.size())

    def _select_all(self):
        """Выбирает все видимые элементы."""
        for i in range(self.value_list.count()):
            item = self.value_list.item(i)
            if not item.isHidden():
                self.value_list._check_states[i] = Qt.CheckState.Checked
        self.value_list.viewport().update()
        self._update_info()

    def _clear_all(self):
        """Очищает все видимые элементы."""
        for i in range(self.value_list.count()):
            item = self.value_list.item(i)
            if not item.isHidden():
                self.value_list._check_states[i] = Qt.CheckState.Unchecked
        self.value_list.viewport().update()
        self._update_info()

    def _update_info(self):
        total = sum(
            1
            for i in range(self.value_list.count())
            if not self.value_list.item(i).isHidden()
        )
        selected = sum(
            1
            for i in range(self.value_list.count())
            if not self.value_list.item(i).isHidden()
            and self.value_list._check_states.get(i, Qt.CheckState.Unchecked)
            == Qt.CheckState.Checked
        )
        self.info_label.setText(f"Выбрано: {selected} из {total}")

    def get_selected_values(self) -> List[str]:
        return [
            self.value_list.item(i).text()
            for i in range(self.value_list.count())
            if self.value_list._check_states.get(i, Qt.CheckState.Unchecked)
            == Qt.CheckState.Checked
        ]


class PivotSettingsDialog(QDialog):
    """
    Диалог настроек сводной таблицы.
    """

    def __init__(self, columns: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки сводной таблицы")
        self.setMinimumSize(700, 500)
        self.setModal(True)

        # Применяем тему родителя если есть
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        self._columns = columns
        self._init_ui()
        apply_theme_to_dialog(self, self._is_dark)

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # Чекбокс удаления дубликатов
        self.remove_duplicates_checkbox = StyledCheckBox("Удалить дубликаты")
        layout.addWidget(self.remove_duplicates_checkbox)

        # Секция выбора полей
        fields_layout = QHBoxLayout()

        # Фильтры
        filter_group = self._create_list_group("Фильтры", self._columns)
        fields_layout.addWidget(filter_group)

        # Строки
        row_group = self._create_list_group("Строки", self._columns)
        fields_layout.addWidget(row_group)

        # Столбцы
        col_group = self._create_list_group("Столбцы", self._columns)
        fields_layout.addWidget(col_group)

        layout.addLayout(fields_layout)

        # Значения
        values_group = QGroupBox("Значения и агрегации")
        values_layout = QVBoxLayout(values_group)

        self.values_table = QTableWidget()
        self.values_table.setColumnCount(3)
        self.values_table.setHorizontalHeaderLabels(["Значение", "Агрегация", ""])
        self.values_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.values_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.values_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        values_layout.addWidget(self.values_table)

        # Добавление значения
        add_layout = QHBoxLayout()
        add_layout.addWidget(QLabel("Значение:"))
        self.value_combo = QComboBox()
        self.value_combo.addItems(self._columns)
        add_layout.addWidget(self.value_combo)

        add_layout.addWidget(QLabel("Агрегация:"))
        self.agg_combo = QComboBox()
        self.agg_combo.addItems(
            ["Сумма", "Среднее", "Количество", "Максимум", "Минимум"]
        )
        add_layout.addWidget(self.agg_combo)

        add_btn = SecondaryButton("Добавить")
        add_btn.clicked.connect(self._add_value)
        add_layout.addWidget(add_btn)

        values_layout.addLayout(add_layout)
        layout.addWidget(values_group)

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        preview_btn = SecondaryButton("Предпросмотр")
        preview_btn.clicked.connect(self._show_preview)
        button_layout.addWidget(preview_btn)

        ok_btn = PrimaryButton("OK")
        ok_btn.clicked.connect(self._validate_and_accept)
        button_layout.addWidget(ok_btn)

        cancel_btn = SecondaryButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def _create_list_group(self, title: str, items: List[str]) -> QGroupBox:
        """Создаёт группу со списком для выбора."""
        group = QGroupBox(title)
        layout = QVBoxLayout(group)

        list_widget = QListWidget()
        list_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        list_widget.addItems(items)
        layout.addWidget(list_widget)

        # Сохраняем ссылку на list_widget для внешнего доступа
        if title == "Фильтры":
            self.filter_list_group = list_widget
        elif title == "Строки":
            self.row_list_group = list_widget
        elif title == "Столбцы":
            self.col_list_group = list_widget

        return group

    def _add_value(self):
        """Добавляет строку в таблицу значений."""
        row = self.values_table.rowCount()
        self.values_table.insertRow(row)

        # Значение
        value_combo = QComboBox()
        value_combo.addItems(self._columns)
        value_combo.setCurrentText(self.value_combo.currentText())
        self.values_table.setCellWidget(row, 0, value_combo)

        # Агрегация
        agg_combo = QComboBox()
        agg_combo.addItems(["Сумма", "Среднее", "Количество", "Максимум", "Минимум"])
        agg_combo.setCurrentText(self.agg_combo.currentText())
        self.values_table.setCellWidget(row, 1, agg_combo)

        # Кнопка удаления
        delete_btn = QPushButton("Удалить")
        delete_btn.clicked.connect(
            lambda: self.values_table.removeRow(
                self.values_table.indexAt(delete_btn.pos()).row()
            )
        )
        self.values_table.setCellWidget(row, 2, delete_btn)

    def _validate_and_accept(self):
        """Проверяет и принимает настройки."""
        if not (
            self.row_list_group.selectedItems() or self.col_list_group.selectedItems()
        ):
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Выберите хотя бы одно поле для строк или столбцов",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        if self.values_table.rowCount() == 0:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Добавьте хотя бы одно значение для агрегации",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        self.accept()

    def _show_preview(self):
        """Показывает предпросмотр сводной таблицы."""
        from PySide6.QtWidgets import QApplication

        settings = self.get_settings()

        if not (settings["rows"] or settings["columns"]):
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Выберите хотя бы одно поле для строк или столбцов",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        if not settings["values"]:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Добавьте хотя бы одно значение для агрегации",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        # Находим главное окно
        main_window = None
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, QMainWindow):
                main_window = widget
                break

        if not main_window:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Ошибка",
                "Не удалось найти главное окно",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        # Получаем значения фильтра из главного окна
        # Приоритет: сохранённые _pivot_settings, иначе текущие _filter_values
        pivot_settings = getattr(main_window, "_pivot_settings", None)

        if pivot_settings and pivot_settings.get("filter_column"):
            # Используем сохранённые настройки сводной таблицы
            filter_col_text = pivot_settings["filter_column"]
            filter_values = pivot_settings.get("filter_values", [])
        else:
            # Используем текущие значения фильтра
            filter_col_text = (
                main_window.filter_column_combo.currentText()
                if hasattr(main_window, "filter_column_combo")
                else "Не фильтровать"
            )
            if filter_col_text != "Не фильтровать" and hasattr(
                main_window, "_filter_values"
            ):
                filter_values = list(
                    main_window._filter_values.get(filter_col_text, [])
                )
            else:
                filter_values = []

        # Добавляем значения фильтра в настройки
        settings["filter_values"] = filter_values
        settings["filter_column"] = (
            filter_col_text if filter_col_text != "Не фильтровать" else ""
        )

        # Отладка

        # Показываем диалог предпросмотра
        preview_dialog = PivotPreviewDialog(main_window, settings)
        preview_dialog.exec()

    def get_settings(self) -> Dict[str, Any]:
        """Возвращает настройки сводной таблицы."""
        values_settings = []
        for row in range(self.values_table.rowCount()):
            value_combo = self.values_table.cellWidget(row, 0)
            agg_combo = self.values_table.cellWidget(row, 1)
            if value_combo and agg_combo:
                values_settings.append(
                    {
                        "field": value_combo.currentText(),
                        "aggregation": agg_combo.currentText(),
                    }
                )

        return {
            "filters": [item.text() for item in self.filter_list_group.selectedItems()],
            "rows": [item.text() for item in self.row_list_group.selectedItems()],
            "columns": [item.text() for item in self.col_list_group.selectedItems()],
            "values": values_settings,
            "remove_duplicates": self.remove_duplicates_checkbox.isChecked(),
        }


class PivotPreviewDialog(QDialog):
    """
    Диалог предпросмотра сводной таблицы.
    """

    def __init__(self, parent, settings: Dict[str, Any]):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр сводной таблицы")
        self.setMinimumSize(800, 600)
        self.setModal(True)

        # Применяем тему от главного окна
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        self._settings = settings
        self._pivot_data = None

        self._init_ui()
        apply_theme_to_dialog(self, self._is_dark)
        self._load_preview()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Таблица
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        export_btn = SecondaryButton("📊 Экспорт в Excel")
        export_btn.clicked.connect(self._export_to_excel)
        button_layout.addWidget(export_btn)

        copy_btn = SecondaryButton("Копировать")
        copy_btn.clicked.connect(self._copy_to_clipboard)
        button_layout.addWidget(copy_btn)

        close_btn = PrimaryButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)

    def _load_preview(self):
        """Загружает данные для предпросмотра."""
        from converter import PivotTableProcessor

        # Получаем главное окно (родитель)
        main_window = self.parent()
        if not main_window:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Ошибка",
                "Не удалось получить доступ к главному окну",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        # Получаем первый файл из списка
        file_list = main_window.file_list if hasattr(main_window, "file_list") else None
        if not file_list or file_list.count() == 0:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Нет файлов для предпросмотра\n\nДобавьте файл в главном окне",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        file_path = file_list.item(0).text()

        try:
            # Создаём процессор
            processor = PivotTableProcessor(lambda _msg, _color: None)

            # Получаем данные фильтра из настроек
            filter_col = self._settings.get("filter_column", "")
            filter_vals = self._settings.get("filter_values", [])

            self._pivot_data = processor.create_pivot_data(
                file_path, self._settings, filter_col, filter_vals
            )

            if not self._pivot_data:
                msgbox = QMessageBox(
                    QMessageBox.Icon.Warning,
                    "Предупреждение",
                    "Нет данных для отображения\n\nПроверьте:\n1. Выбраны ли поля для строк/столбцов\n2. Добавлены ли значения для агрегации\n3. Корректность настроек",
                    QMessageBox.StandardButton.Ok,
                    self,
                )
                apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
                msgbox.exec()
                return

            # Заполняем таблицу
            self._fill_table()

        except Exception as e:
            import traceback

            msgbox = QMessageBox(
                QMessageBox.Icon.Critical,
                "Ошибка",
                f"Ошибка при загрузке данных:\n{str(e)}\n\n{traceback.format_exc()}",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()

    def _fill_table(self):
        """Заполняет таблицу данными."""
        if not self._pivot_data:
            return

        row_values = sorted(self._pivot_data.keys())
        col_values = sorted(
            set(col for row in self._pivot_data.values() for col in row.keys())
        )
        value_settings = self._settings.get("values", [])

        # Определяем количество колонок
        header_cols = len(self._settings.get("rows", []))
        total_cols = header_cols + len(col_values) * len(value_settings)

        self.table_widget.setRowCount(len(row_values) + 1)
        self.table_widget.setColumnCount(total_cols)

        # Заголовки строк
        for col, header in enumerate(self._settings.get("rows", [])):
            self.table_widget.setItem(0, col, QTableWidgetItem(header))

        # Заголовки столбцов
        current_col = header_cols
        for col_value in col_values:
            for val_setting in value_settings:
                header_text = f"{' / '.join(str(v) for v in col_value)} - {val_setting['field']} ({val_setting['aggregation']})"
                item = QTableWidgetItem(header_text)
                item.setFont(QFont("Arial", 9, QFont.Weight.Bold))
                self.table_widget.setItem(0, current_col, item)
                current_col += 1

        # Данные
        for row_idx, row_value in enumerate(row_values, 1):
            current_col = 0

            # Значения строк
            if isinstance(row_value, tuple):
                for col_idx, value in enumerate(row_value):
                    self.table_widget.setItem(
                        row_idx, col_idx, QTableWidgetItem(str(value))
                    )
                current_col = len(row_value)
            else:
                self.table_widget.setItem(row_idx, 0, QTableWidgetItem(str(row_value)))
                current_col = 1

            # Агрегированные значения
            for col_value in col_values:
                for val_setting in value_settings:
                    key = f"{val_setting['field']}_{val_setting['aggregation']}"
                    value = self._pivot_data[row_value][col_value].get(key, 0)

                    if isinstance(value, float):
                        item = QTableWidgetItem(f"{value:.2f}")
                    else:
                        item = QTableWidgetItem(str(value))

                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                    self.table_widget.setItem(row_idx, current_col, item)
                    current_col += 1

        # Авто-ширина
        self.table_widget.resizeColumnsToContents()

    def _copy_to_clipboard(self):
        """Копирует данные в буфер обмена."""
        from PySide6.QtWidgets import QApplication

        rows = self.table_widget.rowCount()
        cols = self.table_widget.columnCount()

        data = []
        for row in range(rows):
            row_data = []
            for col in range(cols):
                item = self.table_widget.item(row, col)
                row_data.append(item.text() if item else "")
            data.append("\t".join(row_data))

        clipboard = QApplication.clipboard()
        clipboard.setText("\n".join(data))

    def _export_to_excel(self):
        """Экспортирует сводную таблицу в Excel."""
        from PySide6.QtWidgets import QFileDialog
        import xlsxwriter

        if not self._pivot_data:
            msgbox = QMessageBox(
                QMessageBox.Icon.Warning,
                "Предупреждение",
                "Нет данных для экспорта",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить сводную таблицу",
            "",
            "Excel файлы (*.xlsx);;Все файлы (*.*)",
        )

        if save_path:
            try:
                if not save_path.lower().endswith(".xlsx"):
                    save_path += ".xlsx"

                workbook = xlsxwriter.Workbook(save_path)
                worksheet = workbook.add_worksheet("Сводная таблица")

                # Форматы
                header_format = workbook.add_format(
                    {
                        "bold": True,
                        "bg_color": "#C8DCF0",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )

                cell_format = workbook.add_format({"border": 0})

                cell_format_right = workbook.add_format({"border": 0, "align": "right"})

                # Заголовки строк
                header_cols = len(self._settings.get("rows", []))
                for col, header in enumerate(self._settings.get("rows", [])):
                    worksheet.write(0, col, header, header_format)

                # Заголовки столбцов
                row_values = sorted(self._pivot_data.keys())
                col_values = sorted(
                    set(col for row in self._pivot_data.values() for col in row.keys())
                )
                value_settings = self._settings.get("values", [])

                current_col = header_cols
                for col_value in col_values:
                    for val_setting in value_settings:
                        header_text = f"{' / '.join(str(v) for v in col_value)} - {val_setting['field']} ({val_setting['aggregation']})"
                        worksheet.write(0, current_col, header_text, header_format)
                        current_col += 1

                # Данные
                for row_idx, row_value in enumerate(row_values, 1):
                    current_col = 0

                    # Значения строк
                    if isinstance(row_value, tuple):
                        for col_idx, value in enumerate(row_value):
                            worksheet.write(row_idx, col_idx, str(value), cell_format)
                        current_col = len(row_value)
                    else:
                        worksheet.write(row_idx, 0, str(row_value), cell_format)
                        current_col = 1

                    # Агрегированные значения
                    for col_value in col_values:
                        for val_setting in value_settings:
                            key = f"{val_setting['field']}_{val_setting['aggregation']}"
                            value = self._pivot_data[row_value][col_value].get(key, 0)

                            if isinstance(value, float):
                                worksheet.write(
                                    row_idx,
                                    current_col,
                                    round(value, 2),
                                    cell_format_right,
                                )
                            else:
                                worksheet.write(
                                    row_idx, current_col, value, cell_format_right
                                )
                            current_col += 1

                # Авто-ширина
                for col_idx in range(current_col):
                    worksheet.set_column(col_idx, col_idx, 15)

                workbook.close()

                msgbox = QMessageBox(
                    QMessageBox.Icon.Information,
                    "Информация",
                    f"Файл сохранён:\n{os.path.basename(save_path)}",
                    QMessageBox.StandardButton.Ok,
                    self,
                )
                apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
                msgbox.exec()

            except Exception as e:
                msgbox = QMessageBox(
                    QMessageBox.Icon.Critical,
                    "Ошибка",
                    f"Ошибка при экспорте:\n{str(e)}",
                    QMessageBox.StandardButton.Ok,
                    self,
                )
                apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
                msgbox.exec()


class TSVPreviewDialog(QDialog):
    """
    Диалог предпросмотра TSV/CSV файла.
    """

    def __init__(self, file_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Предпросмотр: {os.path.basename(file_path)}")
        self.setMinimumSize(800, 600)
        self.setModal(True)

        # Применяем тему родителя если есть
        self._is_dark = False
        if parent and hasattr(parent, "_is_dark_theme"):
            self._is_dark = parent._is_dark_theme

        self._file_path = file_path
        self._current_page = 1
        self._rows_per_page = 100
        self._headers = []
        self._total_rows = 0
        self._encoding = "utf-8"
        self._delimiter = "\t"

        self._active_search_text = ""
        self._active_search_column = -1
        self._search_match_positions = []

        self._init_ui()
        apply_theme_to_dialog(self, self._is_dark)
        self._load_file_info()
        self._load_page_data()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Поиск
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск:"))

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Текст для поиска...")
        self.search_edit.returnPressed.connect(self._search)
        search_layout.addWidget(self.search_edit)

        self.search_column_combo = QComboBox()
        self.search_column_combo.addItem("Все столбцы")
        search_layout.addWidget(self.search_column_combo)

        search_btn = PrimaryButton("Найти")
        search_btn.clicked.connect(self._search)
        search_layout.addWidget(search_btn)

        self.search_info_label = QLabel("")
        search_layout.addWidget(self.search_info_label)

        search_layout.addStretch()
        layout.addLayout(search_layout)

        # Таблица
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)

        # Навигация
        nav_layout = QHBoxLayout()

        prev_btn = SecondaryButton("← Пред.")
        prev_btn.clicked.connect(self._prev_page)
        nav_layout.addWidget(prev_btn)

        self.page_info_label = QLabel("Стр. 1")
        nav_layout.addWidget(self.page_info_label)

        next_btn = SecondaryButton("След. →")
        next_btn.clicked.connect(self._next_page)
        nav_layout.addWidget(next_btn)

        nav_layout.addWidget(QLabel("Строк на странице:"))
        self.rows_per_page_combo = QComboBox()
        self.rows_per_page_combo.addItems(["50", "100", "200", "500", "1000"])
        self.rows_per_page_combo.setCurrentText("100")
        self.rows_per_page_combo.currentTextChanged.connect(self._change_rows_per_page)
        nav_layout.addWidget(self.rows_per_page_combo)

        nav_layout.addStretch()

        close_btn = SecondaryButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        nav_layout.addWidget(close_btn)

        layout.addLayout(nav_layout)

    def _load_file_info(self):
        """Загружает информацию о файле."""
        try:
            self._encoding = FileUtilities.get_encoding(self._file_path)
            self._delimiter = FileUtilities.get_delimiter(self._file_path)

            with open(
                self._file_path, "r", encoding=self._encoding, errors="replace"
            ) as f:
                reader = csv.reader(f, delimiter=self._delimiter)
                headers = next(reader, [])
                self._headers = headers if headers else []

                self.search_column_combo.clear()
                self.search_column_combo.addItem("Все столбцы")
                self.search_column_combo.addItems(self._headers)

                # Подсчёт строк данных без заголовка
                self._total_rows = sum(1 for _ in reader)

            self._update_page_info()
        except (OSError, IOError, UnicodeDecodeError) as e:
            msgbox = QMessageBox(
                QMessageBox.Icon.Critical,
                "Ошибка",
                f"Ошибка чтения файла: {e}",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()

    def _load_page_data(self):
        """Загружает данные для текущей страницы."""
        try:
            start_row = (self._current_page - 1) * self._rows_per_page

            self.table_widget.clear()
            self.table_widget.setColumnCount(len(self._headers))
            self.table_widget.setHorizontalHeaderLabels(self._headers)
            self.table_widget.setRowCount(0)

            with open(
                self._file_path, "r", encoding=self._encoding, errors="replace"
            ) as f:
                reader = csv.reader(f, delimiter=self._delimiter)

                # Пропускаем заголовок
                next(reader, None)

                # Пропускаем до нужной строки
                for _ in range(start_row):
                    next(reader, None)

                # Читаем страницу
                for row_idx in range(self._rows_per_page):
                    values = next(reader, None)
                    if values is None:
                        break

                    self.table_widget.insertRow(row_idx)
                    for col_idx, value in enumerate(values):
                        item = QTableWidgetItem(str(value))
                        self.table_widget.setItem(row_idx, col_idx, item)

            self.table_widget.resizeColumnsToContents()

            if self._active_search_text:
                self._highlight_search_matches_on_page()
        except (OSError, IOError, UnicodeDecodeError) as e:
            msgbox = QMessageBox(
                QMessageBox.Icon.Critical,
                "Ошибка",
                f"Ошибка загрузки данных: {e}",
                QMessageBox.StandardButton.Ok,
                self,
            )
            apply_theme_to_messagebox(msgbox, get_widget_theme_flag(self))
            msgbox.exec()

    def _update_page_info(self):
        """Обновляет информацию о странице."""
        total_pages = max(
            1, (self._total_rows + self._rows_per_page - 1) // self._rows_per_page
        )
        self.page_info_label.setText(
            f"Стр. {self._current_page} из {total_pages} (Всего: {self._total_rows})"
        )

    def _prev_page(self):
        if self._current_page > 1:
            self._current_page -= 1
            self._load_page_data()
            self._update_page_info()

    def _next_page(self):
        total_pages = max(
            1, (self._total_rows + self._rows_per_page - 1) // self._rows_per_page
        )
        if self._current_page < total_pages:
            self._current_page += 1
            self._load_page_data()
            self._update_page_info()

    def _change_rows_per_page(self, value: str):
        self._rows_per_page = int(value)
        self._current_page = 1
        self._load_page_data()
        self._update_page_info()

    def _row_matches_search(
        self, values: list, search_text: str, column_idx: int
    ) -> bool:
        """Проверяет, соответствует ли строка критерию поиска."""
        search_lower = search_text.lower()
        if column_idx >= 0 and column_idx < len(values):
            return search_lower in str(values[column_idx]).lower()
        for value in values:
            if search_lower in str(value).lower():
                return True
        return False

    def _highlight_search_matches_on_page(self):
        """Подсвечивает ячейки с совпадениями на текущей странице."""
        if not self._active_search_text:
            return
        search_lower = self._active_search_text.lower()
        for row_idx in range(self.table_widget.rowCount()):
            for col_idx in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row_idx, col_idx)
                if item and search_lower in item.text().lower():
                    item.setBackground(QColor(255, 255, 0))
                else:
                    item.setBackground(QPalette().window())

    def _search(self):
        """Выполняет поиск по файлу."""
        search_text = self.search_edit.text()
        column_text = self.search_column_combo.currentText()

        if not search_text:
            self._active_search_text = ""
            self._search_match_positions = []
            self.search_info_label.setText("")
            self._load_page_data()
            return

        search_lower = search_text.lower()
        column_idx = -1
        if column_text != "Все столбцы" and self._headers:
            try:
                column_idx = self._headers.index(column_text)
            except ValueError:
                column_idx = -1

        self._active_search_text = search_text
        self._active_search_column = column_idx
        self._search_match_positions = []

        try:
            with open(
                self._file_path, "r", encoding=self._encoding, errors="replace"
            ) as f:
                reader = csv.reader(f, delimiter=self._delimiter)
                next(reader, None)

                for row_idx, row in enumerate(reader):
                    if self._row_matches_search(row, search_lower, column_idx):
                        self._search_match_positions.append(row_idx)

        except Exception as e:
            self.search_info_label.setText(f"Ошибка поиска: {e}")
            return

        if not self._search_match_positions:
            self.search_info_label.setText("Ничего не найдено")
            return

        first_match_row = self._search_match_positions[0]
        match_page = first_match_row // self._rows_per_page + 1

        total_matches = len(self._search_match_positions)
        self.search_info_label.setText(
            f"Найдено: {total_matches}, первая на стр. {match_page}"
        )

        self._current_page = match_page
        self._load_page_data()
        self._update_page_info()


# ============================================================================
# ГЛАВНОЕ ОКНО
# ============================================================================


class MainWindow(QMainWindow):
    """
    Главное окно приложения.
    Использует только layout-ы для размещения элементов.
    """

    # Сигналы для связи с бизнес-логикой
    conversion_started = Signal()
    conversion_stopped = Signal()
    files_added = Signal(object)  # list of file paths
    preview_requested = Signal()
    pivot_settings_requested = Signal()
    open_converted_file_requested = Signal()
    delete_converted_file_requested = Signal()
    export_report_requested = Signal()
    settings_saved = Signal(dict)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Конвертер TSV/CSV в Excel")
        self.setMinimumWidth(900)  # Только ширина, высоту контролирует layout

        # Хранилище состояния
        self._settings: Dict[str, Any] = {}
        self._header_color = QColor(200, 220, 240)
        self._is_dark_theme = False

        # Данные фильтра и сводной таблицы
        self._filter_values: Dict[str, List[str]] = {}
        self._selected_column_values: Dict[str, List[str]] = {}
        self._pivot_settings: Optional[Dict[str, Any]] = None

        # Для обновления прогресса
        self._current_progress_data = None

        # Лоадер загрузки
        self._loading_overlay = None
        self._file_worker = None

        self._init_ui()
        self._apply_theme()

    def _show_message_box(
        self,
        icon: QMessageBox.Icon,
        title: str,
        text: str,
        buttons: QMessageBox.StandardButton = QMessageBox.StandardButton.Ok,
    ) -> QMessageBox:
        """Создаёт QMessageBox с применённой темой."""
        msgbox = QMessageBox(icon, title, text, buttons, self)
        apply_theme_to_messagebox(msgbox, self._is_dark_theme)
        return msgbox

    def _init_ui(self):
        """Инициализация интерфейса."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setSpacing(12)
        self.main_layout.setContentsMargins(16, 16, 16, 16)

        # Верхняя панель с файлами
        self.file_panel = self._create_file_panel()
        self.main_layout.addWidget(self.file_panel)

        # Разделитель
        self.main_layout.addWidget(create_separator_line())

        # Панель настроек
        settings_panel = self._create_settings_panel()
        self.main_layout.addWidget(settings_panel)

        # Разделитель
        self.main_layout.addWidget(create_separator_line())

        # Панель управления
        control_panel = self._create_control_panel()
        self.main_layout.addWidget(control_panel)

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(24)
        self.main_layout.addWidget(self.progress_bar)

        # Спойлер журнала событий
        self.toggle_log_btn = QToolButton()
        self.toggle_log_btn.setText("▶ Показать журнал событий")
        self.toggle_log_btn.setCheckable(True)
        self.toggle_log_btn.setChecked(False)
        self.toggle_log_btn.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextBesideIcon
        )
        self.toggle_log_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.toggle_log_btn.setAutoRaise(True)
        self.toggle_log_btn.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.toggle_log_btn.toggled.connect(self._on_toggle_log)

        # Шрифт для кнопки спойлера чуть крупнее
        font = self.toggle_log_btn.font()
        font.setBold(True)
        self.toggle_log_btn.setFont(font)

        self.main_layout.addWidget(self.toggle_log_btn)

        # Лог (скрыт по умолчанию)
        self.log_group = QGroupBox("Журнал событий")
        self.log_group.setVisible(False)
        log_layout = QVBoxLayout(self.log_group)
        self.log_text = LogTextEdit()
        log_layout.addWidget(self.log_text)
        self.main_layout.addWidget(self.log_group)

        # Распорка для предотвращения растяжения панели файлов
        self._spacer = QWidget()
        self._spacer.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.main_layout.addWidget(self._spacer)

        # Нижняя панель кнопок
        bottom_panel = self._create_bottom_panel()
        self.main_layout.addWidget(bottom_panel)

        # Создаём лоадер после инициализации UI
        self._create_loading_overlay()

    def _on_toggle_log(self, checked: bool):
        """Переключение видимости лога."""
        self.log_group.setVisible(checked)
        self._spacer.setVisible(not checked)
        if checked:
            self.toggle_log_btn.setText("▼ Скрыть журнал событий")
            self.toggle_log_btn.setArrowType(Qt.ArrowType.DownArrow)
        else:
            self.toggle_log_btn.setText("▶ Показать журнал событий")
            self.toggle_log_btn.setArrowType(Qt.ArrowType.RightArrow)

        # Пересчитываем размер окна (с задержкой, чтобы layout успел обновиться)
        QTimer.singleShot(0, self.adjustSize)

    def _create_file_panel(self) -> QWidget:
        """Панель управления файлами."""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # Список файлов
        self.file_list = FileListWidget()
        self.file_list.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self.file_list, stretch=4)

        # Кнопки
        button_layout = QVBoxLayout()
        button_layout.setSpacing(8)

        self.add_files_btn = PrimaryButton("Добавить файлы")
        self.add_files_btn.clicked.connect(self._select_files)
        button_layout.addWidget(self.add_files_btn)

        self.preview_btn = SecondaryButton("Предпросмотр")
        self.preview_btn.clicked.connect(self._preview_file)
        button_layout.addWidget(self.preview_btn)

        self.remove_file_btn = SecondaryButton("Удалить")
        self.remove_file_btn.clicked.connect(self._remove_selected_file)
        button_layout.addWidget(self.remove_file_btn)

        button_layout.addStretch()
        layout.addLayout(button_layout, stretch=1)

        # Панель должна оставаться компактной, а не забирать лишнюю высоту окна.
        widget.setMaximumHeight(widget.minimumSizeHint().height())

        return widget

    def _create_settings_panel(self) -> QWidget:
        """Панель настроек конвертации."""
        widget = QWidget()
        layout = QGridLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setVerticalSpacing(10)
        layout.setHorizontalSpacing(16)

        # Строка 1: Формат и директория
        layout.addWidget(QLabel("Формат:"), 0, 0)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["xlsx", "csv"])
        layout.addWidget(self.format_combo, 0, 1)

        layout.addWidget(QLabel("Сохранить в:"), 0, 2)
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("Выберите папку...")
        layout.addWidget(self.output_path_edit, 0, 3)

        browse_btn = SecondaryButton("Обзор...")
        browse_btn.clicked.connect(self._select_output_directory)
        layout.addWidget(browse_btn, 0, 4)

        # Строка 2: Разделение и фильтр
        layout.addWidget(QLabel("Разделить по столбцу:"), 1, 0)
        self.split_column_combo = QComboBox()
        self.split_column_combo.addItem("Не разделять")
        layout.addWidget(self.split_column_combo, 1, 1)

        layout.addWidget(QLabel("Фильтр по столбцу:"), 1, 2)
        self.filter_column_combo = QComboBox()
        self.filter_column_combo.addItem("Не фильтровать")
        layout.addWidget(self.filter_column_combo, 1, 3)

        # Строка 3: Стили - с увеличенными отступами
        styles_layout = QHBoxLayout()
        styles_layout.setSpacing(12)  # Увеличенный отступ между элементами

        styles_layout.addWidget(QLabel("Стили:"))
        styles_layout.addSpacing(8)  # Дополнительный отступ после заголовка

        self.bold_checkbox = StyledCheckBox("Жирный")
        self.bold_checkbox.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        styles_layout.addWidget(self.bold_checkbox)
        styles_layout.addSpacing(4)

        self.italic_checkbox = StyledCheckBox("Курсив")
        self.italic_checkbox.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        styles_layout.addWidget(self.italic_checkbox)
        styles_layout.addSpacing(12)  # Отступ между группами

        styles_layout.addWidget(QLabel("Размер:"))
        self.font_size_spinbox = QSpinBox()
        self.font_size_spinbox.setRange(8, 24)
        self.font_size_spinbox.setValue(12)
        self.font_size_spinbox.setMinimumWidth(60)
        styles_layout.addWidget(self.font_size_spinbox)
        styles_layout.addSpacing(8)

        styles_layout.addWidget(QLabel("Шрифт:"))
        self.font_combo = QFontComboBox()
        self.font_combo.setMinimumWidth(150)
        styles_layout.addWidget(self.font_combo)
        styles_layout.addSpacing(12)  # Отступ между группами

        self.border_checkbox = StyledCheckBox("Границы")
        self.border_checkbox.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        styles_layout.addWidget(self.border_checkbox)
        styles_layout.addSpacing(8)

        self.header_color_btn = SecondaryButton("Цвет шапки")
        self.header_color_btn.clicked.connect(self._select_header_color)
        styles_layout.addWidget(self.header_color_btn)

        styles_layout.addStretch()
        layout.addLayout(styles_layout, 2, 0, 1, 5)

        return widget

    def _create_control_panel(self) -> QWidget:
        """Панель кнопок управления."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Верхняя строка: кнопки
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.start_btn = PrimaryButton("▶ Начать конвертацию")
        self.start_btn.clicked.connect(self._start_conversion)
        buttons_layout.addWidget(self.start_btn)

        self.stop_btn = SecondaryButton("⏹ Остановить")
        self.stop_btn.clicked.connect(self._stop_conversion)
        self.stop_btn.setEnabled(False)
        buttons_layout.addWidget(self.stop_btn)

        self.pivot_btn = SecondaryButton("📊 Сводная таблица")
        self.pivot_btn.clicked.connect(self._show_pivot_settings)
        buttons_layout.addWidget(self.pivot_btn)

        buttons_layout.addStretch()

        # Инфо
        self.total_rows_label = QLabel("Строк: 0")
        buttons_layout.addWidget(self.total_rows_label)

        self.timer_label = QLabel("Время: 00:00")
        buttons_layout.addWidget(self.timer_label)

        layout.addLayout(buttons_layout)

        # Панель детального прогресса (скрыта по умолчанию)
        self.progress_details_widget = QWidget()
        self.progress_details_widget.setVisible(False)
        details_layout = QGridLayout(self.progress_details_widget)
        details_layout.setContentsMargins(8, 8, 8, 8)
        details_layout.setVerticalSpacing(6)
        details_layout.setHorizontalSpacing(16)

        # Фон панели
        palette = self.progress_details_widget.palette()
        if self._is_dark_theme:
            palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
        else:
            palette.setColor(QPalette.ColorRole.Window, QColor(240, 245, 250))
        self.progress_details_widget.setPalette(palette)
        self.progress_details_widget.setAutoFillBackground(True)

        # Строка 1: Файл и операция
        details_layout.addWidget(QLabel("📁 Файл:"), 0, 0)
        self.progress_file_label = QLabel("—")
        self.progress_file_label.setStyleSheet("font-weight: bold;")
        self.progress_file_label.setWordWrap(False)
        self.progress_file_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        details_layout.addWidget(
            self.progress_file_label, 0, 1, 1, 1
        )  # row, col, rowSpan, colSpan

        details_layout.addWidget(QLabel("🔄 Операция:"), 0, 2)
        self.progress_operation_label = QLabel("—")
        self.progress_operation_label.setWordWrap(False)
        details_layout.addWidget(self.progress_operation_label, 0, 3)

        # Строка 2: Строки
        details_layout.addWidget(QLabel("📊 Обработано:"), 1, 0)
        self.progress_rows_label = QLabel("0 / 0")
        details_layout.addWidget(self.progress_rows_label, 1, 1)

        details_layout.addWidget(QLabel("⏱ Прошло:"), 1, 2)
        self.progress_elapsed_label = QLabel("00:00")
        details_layout.addWidget(self.progress_elapsed_label, 1, 3)

        # Строка 3: ETA и скорость
        details_layout.addWidget(QLabel("⏳ Осталось:"), 2, 0)
        self.progress_eta_label = QLabel("--:--")
        details_layout.addWidget(self.progress_eta_label, 2, 1)

        details_layout.addWidget(QLabel("⚡ Скорость:"), 2, 2)
        self.progress_speed_label = QLabel("0 строк/сек")
        details_layout.addWidget(self.progress_speed_label, 2, 3)

        layout.addWidget(self.progress_details_widget)

        return widget

    def _create_bottom_panel(self) -> QWidget:
        """Нижняя панель с кнопками действий."""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.open_file_btn = SecondaryButton("📂 Открыть файл")
        self.open_file_btn.clicked.connect(self._open_converted_file)
        self.open_file_btn.setEnabled(False)
        layout.addWidget(self.open_file_btn)

        self.delete_file_btn = DangerButton("🗑 Удалить файл")
        self.delete_file_btn.clicked.connect(self._delete_converted_file)
        self.delete_file_btn.setEnabled(False)
        layout.addWidget(self.delete_file_btn)

        layout.addStretch()

        self.export_report_btn = SecondaryButton("📄 Экспорт отчёта")
        self.export_report_btn.clicked.connect(self._export_report)
        layout.addWidget(self.export_report_btn)

        self.settings_btn = SecondaryButton("⚙ Настройки")
        self.settings_btn.clicked.connect(self._show_settings)
        layout.addWidget(self.settings_btn)

        self.close_btn = DangerButton("✕ Закрыть")
        self.close_btn.clicked.connect(self.close)
        layout.addWidget(self.close_btn)

        return widget

    # ========================================================================
    # ОБРАБОТЧИКИ СОБЫТИЙ
    # ========================================================================

    def _select_files(self):
        """Выбор файлов через диалог."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите файлы", "", "TSV/CSV файлы (*.tsv *.csv);;Все файлы (*.*)"
        )
        if files:
            self._add_files(files)

    def _on_files_dropped(self, files: List[str]):
        """Обработка перетаскивания файлов."""
        self._add_files(files)

    def _add_files(self, files: List[str]):
        """Добавление файлов в список с обработкой в фоне."""
        # Показываем лоадер
        self._show_loading_overlay("Обработка файлов...")

        # Блокируем интерфейс на время обработки
        self.setEnabled(False)

        # Запускаем worker для обработки файлов
        self._file_worker = LoadingWorker(self._process_files_operation, files)
        self._file_worker.finished.connect(self._on_files_processing_finished)
        self._file_worker.error.connect(self._on_files_processing_error)
        self._file_worker.start()

    def _process_files_operation(self, files: List[str]) -> tuple:
        """
        Операция обработки файлов (выполняется в фоне).
        Возвращает кортеж (файлы, общее количество строк, заголовки первого файла).
        """
        total_rows = 0
        first_headers = None

        files_with_rows = []

        for file_path in files:
            try:
                encoding = FileUtilities.get_encoding(file_path)
                delimiter = FileUtilities.get_delimiter(file_path)

                # Подсчёт строк
                rows = FileUtilities.count_rows(
                    file_path, delimiter, encoding, None, None
                )
                if rows > 0:
                    total_rows += rows
                    files_with_rows.append((file_path, rows))

                # Чтение заголовков первого файла
                if first_headers is None:
                    with open(file_path, "r", encoding=encoding, errors="replace") as f:
                        reader = csv.reader(f, delimiter=delimiter)
                        first_headers = next(reader)

            except Exception:
                # Продолжаем обработку даже если один файл не удался
                continue

        return (files_with_rows, total_rows, first_headers)

    def _on_files_processing_finished(self, result: tuple):
        """Обработчик завершения обработки файлов."""
        # Скрываем лоадер
        self._hide_loading_overlay()

        # Разблокируем интерфейс
        self.setEnabled(True)

        # Распаковываем результат
        files_with_rows, added_rows, headers = result

        # Добавляем файлы в список с сохранением количества строк в userData
        for file_path, row_count in files_with_rows:
            item = QListWidgetItem(file_path)
            item.setData(Qt.ItemDataRole.UserRole, row_count)
            self.file_list.addItem(item)

        # Обновляем состояние комбобоксов split/filter (блокируем при нескольких файлах)
        self._update_split_filter_state()

        # Обновляем комбобоксы столбцов (только если нет заголовков или есть новые)
        if headers:
            existing_count = self.file_list.count() - len(files_with_rows)
            if existing_count == 0:
                self._update_column_combos(headers)

        # Обновляем количество строк (прибавляем к существующему)
        current_text = self.total_rows_label.text()
        current_total = 0
        if "Строк: " in current_text:
            try:
                current_total = int(
                    current_text.replace("Строк: ", "").replace(",", "")
                )
            except ValueError:
                pass
        new_total = current_total + added_rows
        self.total_rows_label.setText(f"Строк: {new_total:,}")

        # Emit сигнал
        self.files_added.emit([f for f, _ in files_with_rows])

        self.log_message(
            f"Добавлено файлов: {len(files_with_rows)}, строк: {added_rows:,}, всего: {new_total:,}",
            QColor("green"),
        )

        # Сбрасываем worker
        self._file_worker = None

    def _update_column_combos(self, headers: List[str]):
        """Обновляет комбобоксы столбцов."""
        # Сохраняем текущие значения
        current_split = self.split_column_combo.currentText()
        current_filter = self.filter_column_combo.currentText()

        # Очищаем
        self.split_column_combo.clear()
        self.filter_column_combo.clear()

        self.split_column_combo.addItem("Не разделять")
        self.filter_column_combo.addItem("Не фильтровать")
        self.split_column_combo.addItems(headers)
        self.filter_column_combo.addItems(headers)

        # Восстанавливаем значения если возможно
        if current_split in headers:
            self.split_column_combo.setCurrentText(current_split)
        if current_filter in headers:
            self.filter_column_combo.setCurrentText(current_filter)

    def _on_files_processing_error(self, error: str):
        """Обработчик ошибки обработки файлов."""
        # Скрываем лоадер
        self._hide_loading_overlay()

        # Разблокируем интерфейс
        self.setEnabled(True)

        # Показываем ошибку
        msgbox = self._show_message_box(
            QMessageBox.Icon.Critical,
            "Ошибка",
            f"Ошибка при обработке файлов:\n{error}",
        )
        msgbox.exec()

        # Сбрасываем worker
        self._file_worker = None

    def _update_split_filter_state(self):
        """Обновляет состояние комбобоксов split/filter в зависимости от количества файлов."""
        file_count = self.file_list.count()
        if file_count > 1:
            self.split_column_combo.setEnabled(False)
            self.filter_column_combo.setEnabled(False)
            self.split_column_combo.setToolTip(
                "Функция недоступна при загрузке нескольких файлов. "
                "Split/filter работает только для одного файла."
            )
            self.filter_column_combo.setToolTip(
                "Функция недоступна при загрузке нескольких файлов. "
                "Split/filter работает только для одного файла."
            )
        else:
            self.split_column_combo.setEnabled(True)
            self.filter_column_combo.setEnabled(True)
            self.split_column_combo.setToolTip("")
            self.filter_column_combo.setToolTip("")

    def _remove_selected_file(self):
        """Удаление выбранного файла."""
        current_item = self.file_list.currentItem()
        if current_item:
            row = self.file_list.row(current_item)
            row_count = current_item.data(Qt.ItemDataRole.UserRole) or 0
            self.file_list.takeItem(row)

            # Вычитаем строки удалённого файла
            current_text = self.total_rows_label.text()
            if "Строк: " in current_text:
                try:
                    current_total = int(
                        current_text.replace("Строк: ", "").replace(",", "")
                    )
                    new_total = max(0, current_total - row_count)
                    self.total_rows_label.setText(f"Строк: {new_total:,}")
                except ValueError:
                    pass

            self._update_split_filter_state()
            self.log_message("Файл удалён из списка", QColor("orange"))

    def _preview_file(self):
        """Запрашивает предпросмотр выбранного файла."""
        self.preview_requested.emit()

    def _select_output_directory(self):
        """Выбор директории для сохранения."""
        directory = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if directory:
            self.output_path_edit.setText(directory)
            self.log_message(f"Папка сохранения: {directory}", QColor("green"))

    def _select_header_color(self):
        """Выбор цвета шапки."""
        dialog = QColorDialog(self._header_color, self)
        dialog.setWindowTitle("Выберите цвет")
        apply_theme_to_dialog(dialog, self._is_dark_theme)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            color = dialog.selectedColor()
            if color.isValid():
                self._header_color = color
                # Визуальная индикация
                pal = self.header_color_btn.palette()
                pal.setColor(QPalette.ColorRole.Button, color)
                self.header_color_btn.setPalette(pal)

    def _start_conversion(self):
        """Запуск конвертации."""
        # Валидация
        if self.file_list.count() == 0:
            msgbox = self._show_message_box(
                QMessageBox.Icon.Warning, "Ошибка", "Добавьте файлы для конвертации"
            )
            msgbox.exec()
            return

        if not self.output_path_edit.text():
            msgbox = self._show_message_box(
                QMessageBox.Icon.Warning, "Ошибка", "Выберите папку для сохранения"
            )
            msgbox.exec()
            return

        # Переключение состояния кнопок
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)

        self.log_message("Запуск конвертации...", QColor("blue"))

        # Сигнал для бизнес-логики
        self.conversion_started.emit()

    def _stop_conversion(self):
        """Остановка конвертации."""
        self.stop_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.conversion_stopped.emit()
        self.log_message("Конвертация остановлена пользователем", QColor("orange"))

    def _show_pivot_settings(self):
        """Запрашивает открытие настроек сводной таблицы."""
        self.pivot_settings_requested.emit()

    def _show_settings(self):
        """Показать настройки приложения."""
        dialog = SettingsDialog(self)
        dialog.load_settings(self._settings)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._settings = dialog.get_settings()
            self._apply_theme()
            self.settings_saved.emit(self._settings)
            self.log_message("Настройки сохранены", QColor("green"))

    def _open_converted_file(self):
        """Запрашивает открытие сконвертированного файла."""
        self.open_converted_file_requested.emit()

    def _delete_converted_file(self):
        """Запрашивает удаление сконвертированного файла."""
        self.delete_converted_file_requested.emit()

    def _export_report(self):
        """Запрашивает экспорт отчёта."""
        self.export_report_requested.emit()

    # ========================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
    # ========================================================================

    def log_message(self, message: str, color: QColor = QColor("black")):
        """Добавление сообщения в лог."""
        self.log_text.append_colored(message, color)

    def _create_loading_overlay(self):
        """Создаёт оверлей загрузки."""
        self._loading_overlay = LoadingOverlay(self, "Обработка файлов...")
        # Устанавливаем размер равный размеру центрального виджета
        if self.centralWidget():
            self._loading_overlay.resize(self.centralWidget().size())

    def _show_loading_overlay(self, text: str = "Обработка..."):
        """Показывает оверлей загрузки."""
        if self._loading_overlay:
            self._loading_overlay.set_text(text)
            self._loading_overlay.start_animation()
            # Обновляем размер
            if self.centralWidget():
                self._loading_overlay.resize(self.centralWidget().size())

    def _hide_loading_overlay(self):
        """Скрывает оверлей загрузки."""
        if self._loading_overlay:
            self._loading_overlay.stop_animation()

    def resizeEvent(self, event):
        """Обновляет размер лоадера при изменении размера окна."""
        super().resizeEvent(event)
        if self._loading_overlay:
            self._loading_overlay.resize(self.centralWidget().size())

    def _apply_theme(self):
        """Применение темы."""
        is_dark = self._settings.get("theme", "Светлая") == "Тёмная"
        self._is_dark_theme = is_dark
        palette = StylePalette.DARK_THEME if is_dark else StylePalette.LIGHT_THEME
        apply_native_style(self, palette, is_dark)

        # Обновляем тему лога
        self.log_text.set_theme(is_dark)

        # Обновляем тему панели прогресса
        if hasattr(self, "progress_details_widget"):
            palette = self.progress_details_widget.palette()
            if is_dark:
                palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
            else:
                palette.setColor(QPalette.ColorRole.Window, QColor(240, 245, 250))
            self.progress_details_widget.setPalette(palette)

    def update_progress_details(self, progress_data):
        """
        Обновляет детальную информацию о прогрессе.

        Args:
            progress_data: ProgressData из конвертера
        """
        from converter import ProgressData

        if not isinstance(progress_data, ProgressData):
            return

        # Сохраняем текущие данные
        self._current_progress_data = progress_data

        # Показываем панель прогресса
        if hasattr(self, "progress_details_widget"):
            self.progress_details_widget.setVisible(True)
            self.progress_file_label.setText(progress_data.current_file or "—")
            self.progress_operation_label.setText(
                progress_data.current_operation or "—"
            )
            self.progress_rows_label.setText(
                f"{progress_data.processed_rows:,} / {progress_data.total_rows:,}"
            )
            self.progress_elapsed_label.setText(progress_data.format_elapsed())
            self.progress_eta_label.setText(progress_data.format_eta())
            self.progress_speed_label.setText(progress_data.format_speed())

    def reset_progress_details(self):
        """Сбрасывает панель детального прогресса."""
        if hasattr(self, "progress_details_widget"):
            self.progress_details_widget.setVisible(False)
            self.progress_file_label.setText("—")
            self.progress_operation_label.setText("—")
            self.progress_rows_label.setText("0 / 0")
            self.progress_elapsed_label.setText("00:00")
            self.progress_eta_label.setText("--:--")
            self.progress_speed_label.setText("0 строк/сек")
        self._current_progress_data = None


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================


def create_main_window() -> MainWindow:
    """Фабричная функция для создания главного окна."""
    return MainWindow()


if __name__ == "__main__":
    import sys
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

    window = create_main_window()
    window.show()

    sys.exit(app.exec())
