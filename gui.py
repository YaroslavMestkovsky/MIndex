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

        self.cols_info_input = QLineEdit("Укажите порядок колонок в результирующем файле в указанном формате.")
        self.cols_info_input.setReadOnly(True)
        self.cols_info_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.cols_info_input)

        self.cols_order_input = QLineEdit("ID,Конкурент,ЦК,Цена_Магнит,ЧВХ,Доля,Магнит/ЦК,Магнит*доля продаж,ЦК*доля продаж,ТГ20,ТГ21,ТГ22,ТГ23,ТП,Формат") # noqa e501
        self.cols_order_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.cols_order_input)

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

        self.log_text.clear()

        if not os.path.isdir(self.directory_input.text()):
            self.log_text.append("Ошибка: Директория не выбрана!")
        else:
            self.worker_thread = WorkerThread(self.util)
            self.worker_thread.log_signal.connect(self.update_log)
            self.worker_thread.start()

    def update_log(self, msg):
        """Обновляет лог."""

        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"{now}: {msg}")
