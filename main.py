from os import error
from pity_calculator import PityCalculator

# USER MANUAL:
# 1) Store the Excel workbook in the same folder as this python file.
# 2) Change the FILENAME constant below to the name of the Excel workbook.
# 3) Change the SHEET_NAME constant below to the name of the worksheet
# that contains the log of summon data.
FILE_NAME = "genshin.xlsx"
SHEET_NAME = "Summon_Log"

# 4) Change the two COL_HEADER variables below to match the respective column
# names in Excel.
BANNER_COL_HEADER = "Banner"
STAR_LEVEL_COL_HEADER = "Star_Level"

# 5) Make sure the workbook is NOT open in Excel; otherwise, the file will
# be protected and the script won't be able to write to the file.
# 6) You're ready! Run script and wait for completion. Then reopen your Excel.


calculator = PityCalculator(
    FILE_NAME, SHEET_NAME, BANNER_COL_HEADER, STAR_LEVEL_COL_HEADER
)
