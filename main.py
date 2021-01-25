import os
from tkinter import *
from pity_calculator import PityCalculator

window = Tk()
window.title("Genshin Pity Script for Excel")
window.config(padx=20, pady=20)

title_label = Label(text="Genshin Pity Script Runner",
                    font=("Calibri", 14, "bold"))
title_label.grid(row=0, column=0)
title_label.config(pady=10)

select_label = Label(text="Select the Excel workbook to use:")
select_label.grid(row=1, column=0)

file_list = [f for f in os.listdir(os.getcwd())
             if os.path.isfile(f) and
             (f"{f}".endswith(".xlsx") or f"{f}".endswith(".xls"))
]
listbox = Listbox(height=len(file_list))
for i, f in enumerate(file_list):
    listbox.insert(i, f)

filename_selected = ""
def select_file_from_listbox(event):
    global filename_selected
    filename_selected = listbox.get(listbox.curselection())
listbox.bind("<<ListboxSelect>>", select_file_from_listbox)
listbox.grid(row=2, column=0)

def start_calculator():
    sheet_name = "Summon_Log"
    banner_col_header = "Banner"
    starlvl_col_header = "Star_Level"
    
    print(filename_selected, sheet_name, banner_col_header, starlvl_col_header)
    
    # FIXME: Exception in Tkinter callback. Try refactoring pity_calculator
    # to non-OOP implementation
    # if filename_selected != "":
    #     pity_calculator = PityCalculator(
    #         filename_selected, sheet_name, banner_col_header, starlvl_col_header
    #     )
    #     window.destroy()


run_button = Button(text="Run Script", command=start_calculator)
run_button.grid(row=3, column=0)

window.mainloop()
