import datetime
import os
from PyQt6.QtWidgets import (
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QTextEdit,
    QLabel,
    QLineEdit,
    QFileDialog,
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import (
    Qt,
    QThread,
    pyqtSignal,
)


class WorkerThread(QThread):
    """Поток для выполнения задачи."""

    log_signal = pyqtSignal(str)

    def __init__(self, util):
        super().__init__()
        self.util = util

    def run(self):
        """Основной метод потока."""

        try:
            self.util.manage(log_callback=self.log_signal.emit)
        except Exception as e:
            self.log_signal.emit(f"FAILURE: {str(e)}")


class GUI(QMainWindow):
    def __init__(self, util):
        super().__init__()

        self.util = util

        self.setWindowTitle("Расчет индексов")
        self.setGeometry(100, 100, 800, 400)

        icon_path = "icon.png"
        self.setWindowIcon(QIcon(icon_path))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.file_label = QLabel("Выберите файл (.xlsx):")
        self.layout.addWidget(self.file_label)

        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        self.file_input.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Центрируем текст
        self.layout.addWidget(self.file_input)

        self.browse_button = QPushButton("Выбрать файл")
        self.browse_button.clicked.connect(self.select_file)
        self.layout.addWidget(self.browse_button)

        self.cols_info_input = QLabel("Укажите порядок колонок в результирующем файле в указанном формате:")
        self.layout.addWidget(self.cols_info_input)

        self.cols_order_input = QLineEdit("ID,Конкурент,ЦК,Цена_Магнит,ЧВХ,Доля,Магнит/ЦК,Магнит*доля продаж,ЦК*доля продаж,ТГ20,ТГ21,ТГ22,ТГ23,ТП,Формат") # noqa e501
        self.cols_order_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.cols_order_input)

        self.log_info = QLabel("Процесс исполнения:")
        self.layout.addWidget(self.log_info)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.layout.addWidget(self.log_text)

        self.process_button = QPushButton("Обработать")
        self.process_button.clicked.connect(self.start_processing)
        self.layout.addWidget(self.process_button)

    def select_file(self):
        """Открывает диалог выбора директории и записывает путь в текстовое поле."""

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",  # Начальная директория (пустая строка означает текущую директорию)
            "Excel Files (*.xlsx);;All Files (*)"  # Фильтр для файлов
        )
        if file_path:
            self.file_input.setText(file_path)

    def start_processing(self):
        """Начинает обработку: очищает лог и запускает логику."""

        self.log_text.clear()

        if not self.file_input.text():
            self.log_text.append("Ошибка: Файл не выбран!")
        else:
            self.worker_thread = WorkerThread(self.util)
            self.worker_thread.log_signal.connect(self.update_log)
            self.worker_thread.start()

    def update_log(self, msg):
        """Обновляет лог."""

        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"{now}: {msg}")
