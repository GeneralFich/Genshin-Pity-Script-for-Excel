# Genshin-Pity-Script-for-Excel

A Python-based script to automatically calculate the Pity countdown of Genshin Impact summons.

The script takes an Excel workbook and inserts the Pity countdown results into a new tab within the existing workbook.

Close the workbook in Excel before running the script.


# SUMMON LOG FORMAT

The user should store an Excel-based summon log as follows:


| Banner | Star_Level | ... [extra columns as desired, will not impact script] |
| - | - | - |
|   |   |   |
|   |   |   |

* The relative positions of the "Banner" and "Star_Level" columns do not affect the script.
* If your banner and star level columns are named differently, remember to update the **COL_HEADER** constants in **main.py**.


# PYTHON SCRIPT SETUP:

* Store the Excel workbook in the same folder as the python files.
* In **main.py**:
  * Change the **FILENAME** constant to the name of the Excel workbook.
  * Change the **SHEET_NAME** constant to the name of the worksheet containing the summon log
* Change the two **COL_HEADER** variables below to match the respective column names in your Excel.
* Close the workbook in excel **before** attempting to run **main.py**; otherwise, the Excel file will be protected and the script won't be able to write to the file.
* You're ready! Run **main.py** in the environment of your choice and wait for completion (typically 2-3 seconds). Then reopen your Excel and chech the "Summon_Log" tab.
