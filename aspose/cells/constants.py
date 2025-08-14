"""
Constants for Excel file format specifications and limits.
"""

# Excel worksheet limits
MAX_SHEET_NAME_LENGTH = 31
MAX_ROWS = 1048576  # Excel 2007+ limit
MAX_COLUMNS = 16384  # Excel 2007+ limit (XFD column)
MAX_CELL_CONTENT_LENGTH = 32767  # Maximum characters in a cell

# Column width limits
MIN_COLUMN_WIDTH = 0.0
MAX_COLUMN_WIDTH = 255.0
DEFAULT_COLUMN_WIDTH = 8.43

# Row height limits  
MIN_ROW_HEIGHT = 0.0
MAX_ROW_HEIGHT = 409.5
DEFAULT_ROW_HEIGHT = 15.0

# Style limits
MAX_FONT_SIZE = 409
MIN_FONT_SIZE = 1
DEFAULT_FONT_SIZE = 11

# File format extensions
EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm', '.xlsb'}
CSV_EXTENSIONS = {'.csv', '.tsv'}
TEXT_EXTENSIONS = {'.txt', '.tab'}

# Invalid characters for sheet names
INVALID_SHEET_NAME_CHARS = ['\\', '/', '?', '*', '[', ']', ':']

# Excel date constants
EXCEL_EPOCH_DATE = "1900-01-01"
EXCEL_EPOCH_DATETIME = "1900-01-01T00:00:00"

# Cell reference patterns
CELL_REF_PATTERN = r'^[A-Z]{1,3}[1-9]\d*$'
RANGE_REF_PATTERN = r'^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$'

# Default values
DEFAULT_SHEET_NAME = "Sheet"
DEFAULT_WORKBOOK_NAME = "Workbook"