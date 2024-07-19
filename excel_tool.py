import tkinter as tk
from tkinter import ttk, messagebox
from lib import ExcelHelper


OUTPUT_FILENAME = "output.xlsx"
OUTPUT_SHEETNAME = "filtered data"
FILTERED_COLUMNS = ["Style", "Component"]


# load an excel file
excel_helper = ExcelHelper()


def on_button_click(supplier_listbox, raw_material_listbox, color_listbox):
    supplier_names = [supplier_listbox.get(idx) for idx in supplier_listbox.curselection()]
    raw_materials = [raw_material_listbox.get(idx) for idx in raw_material_listbox.curselection()]
    colors = [color_listbox.get(idx) for idx in color_listbox.curselection()]

    print(f"Selected Supplier Names: {supplier_names}")
    print(f"Selected Raw Materials: {raw_materials}")
    print(f"Selected Colors: {colors}")
    print()
    if len(supplier_names) <= 0 or len(raw_materials) <= 0 or len(colors) <= 0:
        messagebox.showwarning(title="Warning", message="Please select at least one option for each box.")
        return

    excel_helper.export(supplier_names, raw_materials, colors,
                        OUTPUT_FILENAME, OUTPUT_SHEETNAME, FILTERED_COLUMNS)


# Create the main window
root = tk.Tk()
root.title("Select Bars and Buttons")

# Load options
supplier_name_opt = excel_helper.get_col_unique_values("Supplier Name")
raw_material_opt = excel_helper.get_col_unique_values("Raw Material #")
color_opt = excel_helper.get_col_unique_values("Color")


# Function to create and configure a Listbox
def create_listbox(options):
    listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, exportselection=0)
    for option in options:
        listbox.insert(tk.END, option)
    return listbox


# Create and place the Listboxes and Labels
# Supplier Name
supplier_label = ttk.Label(root, text="Supplier Name:")
supplier_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
supplier_listbox = create_listbox(supplier_name_opt)
supplier_listbox.grid(row=0, column=1, padx=5, pady=5, sticky="w")

# Raw Material
raw_material_label = ttk.Label(root, text="Raw Material:")
raw_material_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
raw_material_listbox = create_listbox(raw_material_opt)
raw_material_listbox.grid(row=1, column=1, padx=5, pady=5, sticky="w")

# Color
color_label = ttk.Label(root, text="Color:")
color_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
color_listbox = create_listbox(color_opt)
color_listbox.grid(row=2, column=1, padx=5, pady=5, sticky="w")

button = ttk.Button(root, text="Export",
                    command=lambda: on_button_click(supplier_listbox, raw_material_listbox, color_listbox))
button.grid(row=3, column=0, padx=5, pady=5)

# Start the Tkinter event loop
if __name__ == "__main__":
    root.mainloop()
