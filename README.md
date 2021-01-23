# Genshin-Pity-Script-for-Excel

A Python-based script to automatically calculate the Pity countdown of Genshin Impact summons.

The script takes an Excel workbook and inserts the Pity countdown results into a new tab within the existing workbook.

Close the workbook in Excel before running the script.

# SUMMON LOG FORMAT

The user should use an Excel-based summon log that resemble something like the following:


| Banner | Star_Level | ... [extra columns as desired, will not impact script] |
| - | - | - |
| Character Event | 3 |   |
| Permanent | 4 |   |

* Place the Excel workbook in the same directory as **main.py** and **pity_calculator.py**.
* The summon log worksheet should be a standard table with headers (no non-table rows on top, etc.)
* Data required: (1) a column containing the banner for each summon, (2) a second column containing the star level of the summon result (3-5).
* If your banner and star level columns are named differently, remember to update the two **col_header** variables in **main.py**. Alternatively, you can change your column headers in Excel to match the script defaults for the two columns - "Banner" and "Star_Level" - respectively.
* Note: The relative positions of the "Banner" and "Star_Level" columns do not affect the script (i.e., they can be placed on any column in any order as long as they are on the same table and match the variable names in **main.py**).

## Example Workbook

For an example Excel workbook containing a Summon_Log, see **genshin.xlsx**.

# PYTHON SCRIPT SETUP:

* Store the Excel workbook in the same folder as the python files.
* In **main.py**:
  * Change the **file_name** variable to the name of the Excel workbook.
  * Change the **sheet_name** variable to the name of the worksheet containing the summon log.
* Change the two **col_header** variables to match the respective column names in your Excel log.
* Close the workbook in Excel **before** attempting to run **main.py**; otherwise, the Excel file will be protected and the script won't be able to write to the file.
* You're ready! Run **main.py** in the environment of your choice and wait for completion (typically 2-3 seconds). Then reopen your Excel and chech the "Summon_Log" tab.
