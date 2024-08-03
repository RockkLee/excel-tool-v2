import os

import pandas as pd
from unittest import TestCase
from lib import find_column_unique_values, filter_data_3, export, load_excel_file, ExcelHelper
from dto import ColumnDto, SupplierNameCol, RawMaterialCol, ColorCol
from excel_tool import get_unique_filename

FILE_NAME = './test_data/test_1.xlsx'
SHEET_NAME = '工作表1'

# Just take the first row to extract the columns' names
df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, nrows=1)
col_str_dic = {column: str for column in df.columns}
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
        SupplierNameCol(['Marvel'])
        output_df = filter_data_3(df, SupplierNameCol(['Marvel']), RawMaterialCol(['222']), ColorCol(['red', 'Red']))
        print(output_df["Raw Material #"].unique())


class TestLibExport(TestCase):
    def test_func(self):
        usecols = ["Style", "Color"]
        export(df, "output", "filtered data", usecols)
        export(df, "output_no_filter", "filtered data")


class TestLoadExcelFile(TestCase):
    def test_func(self):
        (file_path, sheet_name) = load_excel_file()
        print(file_path)
        print(sheet_name)


class TestDto(TestCase):
    def test_func(self):
        abstract_class_test = "success"
        try:
            dto = ColumnDto()
        except TypeError as e:
            print(e)
            abstract_class_test = "error"
        self.assertEqual(abstract_class_test, "error")

        dto = SupplierNameCol(['test'])
        print(dto)
        self.assertEqual(dto.__name__, 'Supplier Name')
        self.assertEqual(dto.ls, ['test'])


class TestGetUniqueValues(TestCase):
    def test_filenames_do_not_exist(self):
        base_filename = "output_test"
        expected_filename = "output_test"
        result = get_unique_filename(base_filename, ".txt")
        self.assertEqual(result, expected_filename)

    def test_filenames_exist(self):
        try:
            # Create temporary files to simulate existing files
            with open("output_test.txt", 'w') as f:
                f.write("Temporary file")
            with open("output_test-1.txt", 'w') as f:
                f.write("Temporary file")

            base_filename = "output_test"
            expected_filename = "output_test-2"
            result = get_unique_filename(base_filename, ".txt")
            self.assertEqual(expected_filename, result)
        finally:
            # Clean up temporary files
            if os.path.exists("output_test.txt"):
                os.remove("output_test.txt")
            if os.path.exists("output_test-1.txt"):
                os.remove("output_test-1.txt")


class TestIntegration(TestCase):
    def test_func(self):
        excel_helper = ExcelHelper.__new__(ExcelHelper)
        excel_helper.df = df
        excel_helper.export(['Marvel'], ['222'], ['red', 'Red', 'blue'],
                            'test', 'test')
