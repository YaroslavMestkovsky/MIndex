import sys
import os

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QTextEdit,
    QLabel,
    QLineEdit,
    QFileDialog,
)
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtCore import Qt


class GUI(QMainWindow):
    def __init__(self, util):
        super().__init__()

        self.util = util

        self.setWindowTitle("Расчет индексов")
        self.setGeometry(100, 100, 600, 400)

        # Установка иконки приложения
        icon_path = "icon.png"
        self.setWindowIcon(QIcon(icon_path))

        # Основной контейнер
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Поле выбора директории
        self.directory_label = QLabel("Выберите директорию:")
        self.layout.addWidget(self.directory_label)

        self.directory_input = QLineEdit("Путь до папки с файлами КОНКУРЕНТЫ.xlsx, ДОЛИ.xlsx, ЦЕНЫ.xlsx и ЭТАЛОН.xlsx")
        self.directory_input.setReadOnly(True)
        self.directory_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.directory_input)

        self.browse_button = QPushButton("Выбрать директорию")
        self.browse_button.clicked.connect(self.select_directory)
        self.layout.addWidget(self.browse_button)

        # Кнопка "Обработать"
        self.process_button = QPushButton("Обработать")
        self.process_button.clicked.connect(self.start_processing)
        self.layout.addWidget(self.process_button)

        # Текстовое поле для логов
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.layout.addWidget(self.log_text)

    def select_directory(self):
        """Открывает диалог выбора директории и записывает путь в текстовое поле."""

        directory = QFileDialog.getExistingDirectory(self, "Выберите директорию")

        if directory:
            self.directory_input.setText(directory)

    def start_processing(self):
        """Начинает обработку: очищает лог и запускает логику."""

        if not self.directory_input.text():
            self.log_text.append("Ошибка: Директория не выбрана!")
            return

        self.util.manage()

    def update_log(self, msg):
        """Обновляет лог."""
        self.log_text.append(msg)
