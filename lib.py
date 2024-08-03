from tkinter import filedialog, simpledialog
from typing import Optional

import pandas as pd

from dto import SupplierNameCol, RawMaterialCol, ColorCol, ColumnDto, StyleCol, ComponentCol


def find_column_unique_values(df: pd.DataFrame, target_value: str) -> list:
    if target_value not in df.columns:
        raise ValueError(f"Value '{target_value}' not found in the DataFrame")

    unique_values = df[target_value].unique()
    return unique_values


def filter_data_3(df: pd.DataFrame,
                  col1: ColumnDto, col2: ColumnDto, col3: ColumnDto) -> pd.DataFrame:
    output_df = df[
        (df[col1.__name__].isin(col1.ls)) &
        (df[col2.__name__].isin(col2.ls)) &
        (df[col3.__name__].isin(col3.ls))
    ]
    return output_df


def filter_data_2(df: pd.DataFrame,
                  col1: ColumnDto, col2: ColumnDto) -> pd.DataFrame:
    output_df = df[
        (df[col1.__name__].isin(col1.ls)) &
        (df[col2.__name__].isin(col2.ls))
    ]
    return output_df


def export(df: pd.DataFrame, filename: str, sheetname: str, usecols: Optional[list] = None):
    if usecols is None:
        df.to_excel(f'{filename}.xlsx', sheet_name=sheetname, index=False)
        return

    df_selected = df[usecols]
    df_selected.to_excel(f'{filename}.xlsx', sheet_name=sheetname, index=False)


def load_excel_file() -> tuple:
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if not file_path:
        print("No file selected")
        return None, None

    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    # Prompt the user to select a sheet name
    sheet_name = simpledialog.askstring(
        "Select Sheet",
        f"Available sheets: {xls.sheet_names}\n Enter the sheet name:",
        initialvalue=xls.sheet_names[0]
    )
    if sheet_name not in xls.sheet_names:
        print(f"Sheet name '{sheet_name}' is not valid")
        return None, None

    return file_path, sheet_name


class ExcelHelper:
    def __init__(self):
        file_path, sheet_name = load_excel_file()
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
        col_str_dic = {column: str for column in df.columns}
        self.df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=col_str_dic)
        print(
            f"""\
            sheet_name: {sheet_name}
            file_path: {file_path}
            """
        )

    def get_col_unique_values(self, target_value: str) -> list:
        return find_column_unique_values(self.df, target_value)

    def export(
            self, supplier_names: list, raw_materials: list, colors: list,
            filename: str, sheetname: str
    ):
        output_df = filter_data_3(self.df,
                                  SupplierNameCol(supplier_names), RawMaterialCol(raw_materials), ColorCol(colors))

        styles = find_column_unique_values(output_df, 'Style')
        colors = find_column_unique_values(output_df, 'Color')
        print(f"styles: {styles}")
        print(f"colors: {colors}")
        print()
        output_df = filter_data_2(self.df, StyleCol(styles), ColorCol(colors))

        export(output_df, filename, sheetname)
