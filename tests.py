import pandas as pd
from unittest import TestCase
from lib import find_column_unique_values, filter_data, export, load_excel_file

FILE_NAME = './test_data/test_1.xlsx'
SHEET_NAME = '工作表1'

# Just take the first row to extract the columns' names
df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, nrows=1)
col_str_dic = {column:str for column in df.columns}
df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, dtype=col_str_dic)


class TestLibFindColumnUniqueValues(TestCase):
    def test_column_does_not_exist(self):
        expected = "success"
        try:
            find_column_unique_values(df, 'NonExistentColumn')
        except ValueError as e:
            print(e)
            expected = "error"

        self.assertEqual(expected, "error")

    def test_func(self):
        find_column_unique_values(df, 'Supplier Name')
        find_column_unique_values(df, 'Raw Material #')
        find_column_unique_values(df, 'Color')


class TestLibFilterData(TestCase):
    def test_func(self):
        print(df["Raw Material #"].unique())
        output_df = filter_data(df, 'Marvel', '222', 'red')
        print(output_df["Raw Material #"].unique())


class TestLibExport(TestCase):
    def test_func(self):
        export(df, "output.xlsx", "filtered data",
               ["Style", "Component"])


class TestLoadExcelFile(TestCase):
    def test_func(self):
        (file_path, sheet_name) = load_excel_file()
        print(file_path)
        print(sheet_name)