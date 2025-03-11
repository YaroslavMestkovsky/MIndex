import sys
import os
import pandas as pd

from PyQt6.QtWidgets import QApplication
from gui import GUI

# todo 1: Разобраться с удалением нулевых ЦК и цена магнит
# todo 2: Порядок и наименование столбцов согласно результату (КодТП переименовать в ID) !
# todo 3: Добавить метрики по времени

class IndexUtil:
    def __init__(self):
        self.gui = None
        self.standard_df = None

        self.PARTS = "ДОЛИ.xlsx"
        self.PRICES = "ЦЕНЫ.xlsx"
        self.RIVALS = "КОНКУРЕНТЫ.xlsx"
        self.STANDARD = "ЭТАЛОН.xlsx"

        self.article_tag = "*article"
        self.format_tag = "*format"
        self.parts_tag = "*parts"
        self.self_cost_tag = "*self_cost"
        self.prices_tag = "*prices"
        self.utility_col = "utility_col" # Объединение артикля + формата

        self.DOCS_MAP = {
            self.PARTS: "",
            self.PRICES: "",
            self.RIVALS: "",
            self.STANDARD: "",
        }
        self.DOCS_RULES_MAP = {
            self.STANDARD: {
                "skip": 0, # Сколько строк скипнуть перед чтением
                "tags": {
                    "major_tag": "*dummy", # Имя тега, отмечающего существенные столбцы
                    "article_name": self.article_tag, # Имя столбца с артиклями
                    "format_name": self.format_tag, # Имя столбца с форматами
                },
            },
            self.RIVALS: {
                "skip": 5,
                "tags": {
                    "major_tag": "*rival",
                    "article_tag": self.article_tag,
                    "format_name": "NONE",
                },
            },
            self.PARTS: {
                "skip": 0,
                "tags": {
                    "major_tag": self.parts_tag,
                    "article_name": self.article_tag,
                    "format_name": self.format_tag,
                },
            },
            self.PRICES: {
                "skip": 0,
                "tags": {
                    "major_tag": self.prices_tag,
                    "article_name": self.article_tag,
                    "format_name": self.format_tag,
                },
            },
        }

    def run(self):
        app = QApplication(sys.argv)
        self.gui = GUI(self)
        self.gui.show()
        sys.exit(app.exec())

    def manage(self, log_callback):
        path = self.gui.directory_input.text()
        dir_list = os.listdir(path)
        all_docs_ok = True

        for document in self.DOCS_MAP.keys():
            if not document in dir_list:
                all_docs_ok = False
                log_callback(f"Обязательно наличие файла {document} в директории!")
            else:
                self.DOCS_MAP[document] = os.path.join(path, document)

        if all_docs_ok:
            log_callback("Проверяем наличие файлов: OK")
            self.read_docs()
            log_callback("Читаем файлы: OK")
            self.rename_columns()
            self.merge_rivals_to_standard()
            log_callback("Добавляем конкурентов в эталон: OK")
            self.add_utility_col()
            self.brain_flip()
            log_callback("Переворачиваем таблицу: OK")
            self.merge_prices_to_standard()
            log_callback("Добавляем цены: OK")
            self.merge_parts_to_standard()
            log_callback("Добавляем доли: OK")
            #self.drop_null()
            log_callback("Пустые ЦК убраны: OK")
            self.calculate()
            log_callback("Производим расчеты: OK")
            self.clear_cols_names()

            self.DOCS_MAP[self.STANDARD].to_excel("result.xlsx", index=False)
            log_callback(f"Готово. Результат выгружен в папку: {os.path.join(os.path.curdir, 'result.xlsx')}")
        else:
            log_callback("Файлы: FAILURE\nПроверьте директорию и повторите попытку.")

    def read_docs(self):
        """Чтение файлов."""

        for document in self.DOCS_MAP.keys():
            rules = self.DOCS_RULES_MAP[document]
            tags = rules["tags"]
            df = pd.read_excel(self.DOCS_MAP[document], skiprows=rules["skip"])

            # Фильтруем по тегам
            filtered_columns = [col for col in df.columns if self.check_tags(col, tags.values())]
            df = df[filtered_columns][1:]

            self.DOCS_MAP[document] = df

    def rename_columns(self):
        """Переименовываем технические колонки."""

        for df in self.DOCS_MAP.values():
            columns = []

            for col in df.columns:
                if self.format_tag in col or self.article_tag in col:
                    columns.append(col.split("*")[1])
                else:
                    columns.append(col)

            df.columns = columns

    def add_utility_col(self):
        """Добавляем мердж формата с артиклем."""

        for doc, df in self.DOCS_MAP.items():
            try:
                format_row = (col for col in df.columns if self.format_tag.lstrip("*") in col).__next__()
                article_row = (col for col in df.columns if self.article_tag.lstrip("*") in col).__next__()
            except StopIteration:
                continue
            else:
                df[self.utility_col] = df.apply(lambda row: f"{row[format_row]}{row[article_row]}", axis=1)

                if not doc == self.STANDARD:
                    self.DOCS_MAP[doc] = df.drop(columns=[self.format_tag.lstrip("*"), self.article_tag.lstrip("*")])

    def brain_flip(self):
        """Переворачиваем таблицу, вытряхивая конкурентов... Что-то типа транспонирования?"""

        df = self.DOCS_MAP[self.STANDARD]
        major_tag = self.DOCS_RULES_MAP[self.RIVALS]["tags"]["major_tag"]

        df_melted = pd.melt(
            df,
            id_vars=[col for col in df.columns if major_tag not in col],
            value_vars=[col for col in df.columns if major_tag in col],
            var_name="Конкурент",
            value_name="ЦК",
        )

        df_melted = df_melted.sort_values(by=self.utility_col).reset_index(drop=True)
        self.DOCS_MAP[self.STANDARD] = df_melted

    def merge_rivals_to_standard(self):
        """Мерджим конкурентов в эталон."""

        self.DOCS_MAP[self.STANDARD] = pd.merge(
            self.DOCS_MAP[self.STANDARD],
            self.DOCS_MAP[self.RIVALS],
            on="article",
            how="left",
        )

    def merge_prices_to_standard(self):
        """Мерджим столбцы из цен в эталон."""

        self.DOCS_MAP[self.STANDARD] = pd.merge(
            self.DOCS_MAP[self.STANDARD],
            self.DOCS_MAP[self.PRICES],
            on=self.utility_col,
            how="left",
        )

    def merge_parts_to_standard(self):
        """Мерджим столбцы из долей в эталон."""

        self.DOCS_MAP[self.STANDARD] = pd.merge(
            self.DOCS_MAP[self.STANDARD],
            self.DOCS_MAP[self.PARTS],
            on=self.utility_col,
            how="left",
        )

    def drop_null(self):
        """Дроп строк, которые не нужны пустыми."""

        #todo пропали нулевые доли???
        df = self.DOCS_MAP[self.STANDARD]
        price_col = (
            col for col in df.columns
            if self.prices_tag.lstrip("*") in col
            and self.self_cost_tag not in col
        ).__next__()

        df = df.dropna(subset=['ЦК'])
        df = df.dropna(subset=[price_col])

        self.DOCS_MAP[self.STANDARD] = df

    def calculate(self):
        """Расчет дополнительных колонок."""

        df = self.DOCS_MAP[self.STANDARD]

        part_col = (col for col in df.columns if self.parts_tag.lstrip("*") in col).__next__()
        self_cost_col = (col for col in df.columns if self.self_cost_tag.lstrip("*") in col).__next__()

        df["ЦК*доля продаж"] = df["ЦК"]*df[part_col] #todo названия перепроверить
        df["Магнит*доля продаж"] = df[part_col]*df[self_cost_col] #todo названия перепроверить
        df["Магнит/ЦК"] = df[self_cost_col] / df["ЦК"]

    def clear_cols_names(self):
        """Избавление от тегов."""

        df = self.DOCS_MAP[self.STANDARD]

        columns = [col.split("*")[0] for col in df.columns]
        df.columns = columns
        df.rename(
            columns={
                self.article_tag.lstrip("*"): "КодТП",
                self.format_tag.lstrip("*"): "Формат",
            },
            inplace=True,
        )

        self.DOCS_MAP[self.STANDARD] = df

    @staticmethod
    def check_tags(col, tags):
        """Проверяем, есть ли в названии столбца тег."""
        return any((tag in col for tag in tags))


if __name__ == "__main__":
    util = IndexUtil()
    util.run()
