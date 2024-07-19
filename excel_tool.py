import tkinter as tk
from tkinter import ttk


def on_button_click(supplier_name, raw_material, color):
    selected_value = supplier_name.get()
    raw_selected_value = raw_material.get()
    color_selected_value = color.get()
    print(f"Selected option: {selected_value}")
    print(f"Selected option: {raw_selected_value}")
    print(f"Selected option: {color_selected_value}")
    print()


# Create the main window
root = tk.Tk()
root.title("Select Bars and Buttons")

# load options
supplier_name_opt = ["supplier_name_opt 1", "supplier_name_opt 2", "supplier_name_opt 3"]
raw_material_opt = ["raw_material_opt 1", "raw_material_opt 2", "raw_material_opt 3"]
color_opt = ["color_opt 1", "color_opt 2", "color_opt 3"]
# Create variables for the OptionMenus
supplier_name = tk.StringVar(value=supplier_name_opt[0])
raw_material = tk.StringVar(value=raw_material_opt[0])
color = tk.StringVar(value=color_opt[0])

# Create and place the OptionMenus and Buttons
option_menu1 = ttk.OptionMenu(root, supplier_name, supplier_name_opt[0], *supplier_name_opt)
option_menu1.grid(row=0, column=0, padx=5, pady=5)

option_menu2 = ttk.OptionMenu(root, raw_material, raw_material_opt[0], *raw_material_opt)
option_menu2.grid(row=0, column=2, padx=5, pady=5)

option_menu3 = ttk.OptionMenu(root, color, color_opt[0], *color_opt)
option_menu3.grid(row=0, column=4, padx=5, pady=5)

button3 = ttk.Button(root, text="Button 3", command=lambda: on_button_click(supplier_name, raw_material, color))
button3.grid(row=0, column=5, padx=5, pady=5)

# Start the Tkinter event loop
if __name__ == "__main__":
    root.mainloop()
