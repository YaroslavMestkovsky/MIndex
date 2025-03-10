import sys
import os
from PyQt6.QtWidgets import QApplication

from gui import GUI


class IndexUtil:
    def __init__(self):
        self.gui = None

        self.PARTS = "ДОЛИ.xlsx"
        self.PRICES = "ЦЕНЫ.xlsx"
        self.RIVALS = "КОНКУРЕНТЫ.xlsx"
        self.STANDARD = "ЭТАЛОН.xlsx"
    
        self.DOCS_MAP = {
            self.PARTS: "",
            self.PRICES: "",
            self.RIVALS: "",
            self.STANDARD: "",
        }

    def run(self):
        app = QApplication(sys.argv)
        self.gui = GUI(self)
        self.gui.show()
        sys.exit(app.exec())

    def manage(self):
        self.gui.log_text.clear()
        path = self.gui.directory_input.text()
        dir_list = os.listdir(path)
        all_docs_ok = True

        for document in self.DOCS_MAP.keys():
            if not document in dir_list:
                all_docs_ok = False
                self.gui.update_log(f"Обязательно наличие файла {document} в директории!")
            else:
                self.DOCS_MAP[document] = os.path.join(path, document)

        if all_docs_ok:
            self.gui.update_log("Файлы проверены, начинаем обработку.")

        else:
            self.gui.update_log("Проверьте директорию и повторите попытку.")


if __name__ == "__main__":
    util = IndexUtil()
    util.run()
