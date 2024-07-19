import pandas as pd


def find_column_unique_values(df : pd.DataFrame, target_value: str) -> list:
    if target_value not in df.columns:
        raise ValueError(f"Value '{target_value}' not found in the DataFrame")

    unique_values = df[target_value].unique()
    print(f"All unique values in column '{target_value}': {unique_values}")
    return unique_values


def filter_data(df: pd.DataFrame, supplier_name: str, raw_material: str, color: str) -> pd.DataFrame:
    output_df = df[
        (df['Supplier Name'] == supplier_name) & (df['Raw Material #'] == raw_material) & (df['Color'] == color)
    ]
    # output_df.to_excel('file.xlsx', sheet_name='Filtered Data')  # Saving to a new sheet called Filtered Data
    return output_df


def export(df: pd.DataFrame, usecols: list):
    df_selected = df[usecols]
    df_selected.to_excel('file.xlsx', sheet_name='Filtered Data')
