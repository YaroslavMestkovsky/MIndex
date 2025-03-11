import sys
import os
import pandas as pd

from PyQt6.QtWidgets import QApplication
from gui import GUI


class IndexUtil:
    def __init__(self):
        self.gui = None
        self.standard_df = None

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
        self.DOCS_RULES_MAP = {
            self.STANDARD: {
                'skip': 0, # Сколько строк скипнуть перед чтением
                'tags': {
                    'major_tag': "*dummy", # Имя тега, отмечающего существенные столбцы
                    'article_name': '*article', # Имя столбца с артиклями
                    'format_name': '*format', # Имя столбца с форматами
                },
            },
            self.RIVALS: {
                'skip': 5,
                'tags': {
                    'major_tag': '*rival',
                    'article_tag': '*article',
                },
            },
            self.PARTS: {
                'skip': 0,
                'tags': {
                    'major_tag': '*parts',
                    'article_name': '*article',
                    'format_name': '*format',
                },
            },
            self.PRICES: {
                'skip': 0,
                'tags': {
                    'major_tag': "*prices",
                    'article_name': '*article',
                    'format_name': '*format',
                },
            },
        }

    def run(self):
        app = QApplication(sys.argv)
        self.gui = GUI(self)
        self.gui.show()
        sys.exit(app.exec())
    #todo декоратор на ошибки

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
            self.gui.update_log("Файлы: OK")
            self.read_docs()
            self.gui.update_log("Файлы прочитаны.")
            self.merge_rivals_to_standard()

        else:
            self.gui.update_log("Файлы: FAILURE\nПроверьте директорию и повторите попытку.")

    def read_docs(self):
        """Чтение файлов."""

        for document in self.DOCS_MAP.keys():
            rules = self.DOCS_RULES_MAP[document]
            tags = rules['tags']
            df = pd.read_excel(self.DOCS_MAP[document], skiprows=rules['skip'])

            # Фильтруем по тегам
            filtered_columns = [col for col in df.columns if self.check_tags(col, tags.values())]
            df = df[filtered_columns]

            self.DOCS_MAP[document] = df

    def merge_rivals_to_standard(self):
        """Мерджим конкурентов в эталон."""
        ...

    @staticmethod
    def check_tags(col, tags):
        """Проверяем, есть ли в названии столбца тег."""
        return any((tag in col for tag in tags))


if __name__ == "__main__":
    util = IndexUtil()
    util.run()
