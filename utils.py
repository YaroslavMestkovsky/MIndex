import sys
import os
import pandas as pd

from PyQt6.QtWidgets import QApplication
from gui import GUI


class IndexUtil:
    def __init__(self):
        self.gui = None
        self.path = None
        self.standard_df = None
        self.columns_order = None

        self.PARTS = "ДОЛИ"
        self.PRICES = "ЦЕНЫ"
        self.RIVALS = "КОНКУРЕНТЫ"
        self.STANDARD = "ЭТАЛОН"

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
        self.path = self.gui.file_input.text()
        self.columns_order = self.gui.cols_order_input.text().split(",")

        log_callback("Читаем файл...")
        missing_sheets = self.read_file()

        if missing_sheets:
            log_callback(f"FAILURE. не хватает листов:{missing_sheets}")
            return

        self.rename_columns()

        log_callback("Добавляем конкурентов в эталон...")
        self.merge_rivals_to_standard()

        self.add_utility_col()

        log_callback("Переворачиваем таблицу...")
        self.brain_flip()

        log_callback("Добавляем цены...")
        self.merge_prices_to_standard()

        log_callback("Добавляем доли...")
        self.merge_parts_to_standard()

        log_callback("Убираем пустые ЦК/ЦМ...")
        self.drop_null()

        #log_callback("Убираем не-цифровые ЦК/ЦМ...")
        #self.drop_not_int()

        log_callback("Производим расчеты...")
        self.calculate()

        self.clear_cols_names()
        self.delete_cols()

        try:
            self.move_cols()
        except Exception as e:
            log_callback(f"Сортировка столбцов: FAILURE\nОшибка при сортировке: {e}")
        else:
            log_callback("Сортировка столбцов...")

        log_callback("Выгрузка в файл...")
        self.DOCS_MAP[self.STANDARD].to_excel("result.xlsx", index=False)
        log_callback(f"Готово. Результат выгружен в папку: {os.path.join(os.path.curdir, 'result.xlsx')}")

    def read_file(self):
        """Чтение файлов."""

        raw_doc = pd.ExcelFile(self.path)
        sheet_names = raw_doc.sheet_names
        missing_sheets = []

        for document in self.DOCS_MAP.keys():
            if not document in sheet_names:
                missing_sheets.append(document)
            else:
                rules = self.DOCS_RULES_MAP[document]
                tags = rules["tags"]
                df = raw_doc.parse(document, skiprows=rules["skip"])
                filtered_columns = [col for col in df.columns if self.check_tags(col, tags.values())]
                df = df[filtered_columns][1:]

                self.DOCS_MAP[document] = df

        if missing_sheets:
            return missing_sheets

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
        major_cols = [col for col in df if major_tag in col]
        id_cols = [col for col in df.columns if major_tag not in col]

        value_vars = [col.split("*")[0] for col in major_cols]
        renames = {tagged_col: tagged_col.split("*")[0] for tagged_col in major_cols}

        df.rename(
            columns=renames,
            inplace = True,
        )

        df_melted = pd.melt(
            df,
            id_vars=id_cols,
            value_vars=value_vars,
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

        df = self.DOCS_MAP[self.STANDARD]
        price_col = (
            col for col in df.columns
            if self.prices_tag.lstrip("*") in col
            and self.self_cost_tag not in col
        ).__next__()

        df = df.dropna(subset=['ЦК'])
        df = df.dropna(subset=[price_col])

        self.DOCS_MAP[self.STANDARD] = df

    # def drop_not_int(self):
    #     """Дроп строк, в которых должны быть цифры, но хранятся буквы или иные символы."""
    #
    #     def check_int(val):
    #         return isinstance(val, int)
    #
    #     df = self.DOCS_MAP[self.STANDARD]
    #     price_col = (
    #         col for col in df.columns
    #         if self.prices_tag.lstrip("*") in col
    #            and self.self_cost_tag not in col
    #     ).__next__()
    #
    #     mask = df[['ЦК', price_col]].applymap(check_int).all(axis=1)
    #     filtered_df = df[mask]
    #
    #     self.DOCS_MAP[self.STANDARD] = filtered_df

    def calculate(self):
        """Расчет дополнительных колонок."""

        df = self.DOCS_MAP[self.STANDARD]

        part_col = (col for col in df.columns if self.parts_tag.lstrip("*") in col).__next__()
        price_col = (
            col for col in df.columns
            if self.prices_tag.lstrip("*") in col
               and self.self_cost_tag not in col
        ).__next__()

        df["Магнит/ЦК"] = df[price_col] / df["ЦК"]
        df["MonP"] = df[part_col]*df[price_col]
        df["PRonP"] = df["ЦК"]*df[part_col]

    def clear_cols_names(self):
        """Избавление от тегов."""

        df = self.DOCS_MAP[self.STANDARD]

        self_cost_col = (col for col in df.columns if self.self_cost_tag.lstrip("*") in col).__next__().split("*")[0]
        price_col = (
            col for col in df.columns
            if self.prices_tag.lstrip("*") in col
            and self.self_cost_tag not in col
        ).__next__().split("*")[0]
        columns = [col.split("*")[0] for col in df.columns]

        df.columns = columns
        df.rename(
            columns={
                self.article_tag.lstrip("*"): "ID",
                self.format_tag.lstrip("*"): "Формат",
                price_col: "Цена_Магнит",
                self_cost_col: "ЧВХ",
                "PRonP": "ЦК*доля продаж",
                "MonP": "Магнит*доля продаж",
            },
            inplace=True,
        )

        self.DOCS_MAP[self.STANDARD] = df

    def delete_cols(self):
        """Удаляем ненужные колонки."""

        df = self.DOCS_MAP[self.STANDARD]
        self.DOCS_MAP[self.STANDARD] = df.drop(columns=[self.utility_col])

    def move_cols(self):
        """Перемещение столбцов в правильном порядке."""

        df = self.DOCS_MAP[self.STANDARD]
        self.DOCS_MAP[self.STANDARD] = df[self.columns_order]

    @staticmethod
    def check_tags(col, tags):
        """Проверяем, есть ли в названии столбца тег."""
        return any((tag in col for tag in tags))


if __name__ == "__main__":
    util = IndexUtil()
    util.run()
