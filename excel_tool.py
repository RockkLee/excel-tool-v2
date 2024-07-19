import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
from pandas import ExcelFile
from lib import ExcelHelper


OUTPUT_FILENAME = "output.xlsx"
OUTPUT_SHEETNAME = "filtered data"
FILTERED_COLUMNS = ["Style", "Component"]


excel_helper = ExcelHelper()


def on_button_click(supplier_name_str, raw_material_str, color_str):
    print(f"Option Supplier_name: {supplier_name_str}")
    print(f"Option Raw_material: {raw_material_str}")
    print(f"Option Color: {color_str}")
    print()
    excel_helper.export(supplier_name_str, raw_material_str, color_str,
                        OUTPUT_FILENAME, OUTPUT_SHEETNAME, FILTERED_COLUMNS)


# Create the main window
root = tk.Tk()
root.title("Select Bars and Buttons")

# load options
# supplier_name_opt = ["supplier_name_opt 1", "supplier_name_opt 2", "supplier_name_opt 3"]
# raw_material_opt = ["raw_material_opt 1", "raw_material_opt 2", "raw_material_opt 3"]
# color_opt = ["color_opt 1", "color_opt 2", "color_opt 3"]
supplier_name_opt = excel_helper.get_col_unique_values("Supplier Name")
raw_material_opt = excel_helper.get_col_unique_values("Raw Material #")
color_opt = excel_helper.get_col_unique_values("Color")
# Create variables for the OptionMenus
supplier_name = tk.StringVar(value=supplier_name_opt[0])
raw_material = tk.StringVar(value=raw_material_opt[0])
color = tk.StringVar(value=color_opt[0])

# Create and place the OptionMenus and Buttons
# Supplier Name
supplier_label = ttk.Label(root, text="Supplier Name:")
supplier_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
option_menu1 = ttk.OptionMenu(root, supplier_name, supplier_name_opt[0], *supplier_name_opt)
option_menu1.grid(row=0, column=1, padx=5, pady=5, sticky="w")

# Raw Material
raw_material_label = ttk.Label(root, text="Raw Material:")
raw_material_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
option_menu2 = ttk.OptionMenu(root, raw_material, raw_material_opt[0], *raw_material_opt)
option_menu2.grid(row=1, column=1, padx=5, pady=5, sticky="w")

# Color
color_label = ttk.Label(root, text="Color:")
color_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
option_menu3 = ttk.OptionMenu(root, color, color_opt[0], *color_opt)
option_menu3.grid(row=2, column=1, padx=5, pady=5, sticky="w")

button = ttk.Button(root, text="Export",
                    command=lambda: on_button_click(supplier_name.get(), raw_material.get(), color.get()))
button.grid(row=3, column=0, padx=5, pady=5)

# Start the Tkinter event loop
if __name__ == "__main__":
    root.mainloop()
