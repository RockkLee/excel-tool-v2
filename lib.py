from tkinter import filedialog, simpledialog

import pandas as pd


def find_column_unique_values(df: pd.DataFrame, target_value: str) -> list:
    if target_value not in df.columns:
        raise ValueError(f"Value '{target_value}' not found in the DataFrame")

    unique_values = df[target_value].unique()
    return unique_values


def filter_data(df: pd.DataFrame, supplier_name: str, raw_material: str, color: str) -> pd.DataFrame:
    output_df = df[
        (df['Supplier Name'] == supplier_name) & (df['Raw Material #'] == raw_material) & (df['Color'] == color)
    ]
    return output_df


def export(df: pd.DataFrame, filename, sheetname, usecols: list):
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

    def get_col_unique_values(self, target_value: str) -> list:
        return find_column_unique_values(self.df, target_value)

    def export(
            self, supplier_name: str, raw_material: str, color: str,
            filename: str, sheetname: str, usecols: list
    ):
        output_df = filter_data(self.df, supplier_name, raw_material, color)
        export(output_df, filename, sheetname, usecols)
