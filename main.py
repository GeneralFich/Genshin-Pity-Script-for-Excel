from tkinter import *
from pity_calculator import PityCalculator

# TODO: Make GUI look good
# TODO: Link GUI elements to variables, functions, and objects
# TODO: Store default values for file name, etc. in .txt file

window = Tk()
window.title("Genshin Pity Script for Excel")
window.minsize(width=400, height=600)

title_label = Label(text="Genshin Pity Script Runner",
                    font=("Arial", 12, "bold"),
                    anchor="Left")
title_label.pack()

# USER MANUAL:
# 1) Store the Excel workbook in the same folder as this python file.
# 2) Change the file_name constant below to the name of the Excel workbook.
# 3) Change the sheet_name constant below to the name of the worksheet
# that contains the log of summon data.

filename_label = Label(text="Enter Excel file name with extension:",
                       font=("Arial", 12))
filename_label.pack()
filename_field = Entry()
filename_field.pack()

file_name = "genshin.xlsx"
sheet_name = "Summon_Log"

# 4) Change the two variables below to match the respective column names
# in Excel.
banner_col_header = "Banner"
starlvl_col_header = "Star_Level"

# 5) Make sure the workbook is NOT open in Excel; otherwise, the file will
# be protected and the script won't be able to write to the file.
# 6) You're ready! Run script and wait for completion. Then reopen your Excel.

calculator = PityCalculator(
    file_name, sheet_name, banner_col_header, starlvl_col_header
)

window.mainloop()
