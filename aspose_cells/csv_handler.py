"""
Aspose.Cells for Python - CSV Handler Module

This module provides CSV import and export functionality for workbooks.
It supports reading from and writing to CSV (Comma-Separated Values) files
with configurable delimiters, encodings, and type inference options.

Features:
- Export worksheet data to CSV format
- Import CSV data into a new workbook
- Configurable delimiter (comma, semicolon, tab, etc.)
- Configurable encoding (UTF-8, UTF-16, Latin-1, etc.)
- Automatic type inference for imported values
- Proper handling of quoted fields, escaped characters, and multiline values
"""

import csv
import io
import locale
import re
from datetime import datetime, date, time
from decimal import Decimal, ROUND_HALF_UP
import platform

_is_windows = platform.system() == 'Windows'
from typing import Optional, List, Any


def _system_encoding() -> str:
    """Return the system's preferred encoding (e.g. 'cp1252', 'gbk', 'utf-8')."""
    return locale.getpreferredencoding(False) or 'utf-8'


class CSVLoadOptions:
    """
    Options for loading CSV files.

    Attributes:
        delimiter (str): Field delimiter character. Default is ','.
        encoding (str): File encoding. Default is the system's preferred encoding.
        has_header (bool): Whether the first row contains headers. Default is False.
        quote_char (str): Character used for quoting fields. Default is '"'.
        escape_char (str): Character used for escaping. Default is None (uses doubling).
        skip_rows (int): Number of rows to skip at the beginning. Default is 0.
        auto_detect_types (bool): Automatically detect and convert value types. Default is True.
        date_formats (list): List of date format strings to try for parsing. Default formats included.
        true_values (list): Values to interpret as True. Default is ['true', 'yes', '1'].
        false_values (list): Values to interpret as False. Default is ['false', 'no', '0'].

    Examples:
        >>> options = CSVLoadOptions()
        >>> options.delimiter = ';'
        >>> options.encoding = 'utf-16'
        >>> wb = Workbook()
        >>> CSVHandler.load_csv(wb, 'data.csv', options)
    """

    def __init__(self):
        self.delimiter = ','
        self.encoding = _system_encoding()
        self.has_header = False
        self.quote_char = '"'
        self.escape_char = None
        self.skip_rows = 0
        self.auto_detect_types = True
        self.date_formats = [
            '%Y-%m-%d',
            '%Y/%m/%d',
            '%d-%m-%Y',
            '%d/%m/%Y',
            '%m-%d-%Y',
            '%m/%d/%Y',
            '%Y-%m-%d %H:%M:%S',
            '%Y/%m/%d %H:%M:%S',
            '%d-%m-%Y %H:%M:%S',
            '%d/%m/%Y %H:%M:%S',
        ]
        self.true_values = ['true', 'yes', '1', 'True', 'Yes', 'TRUE', 'YES']
        self.false_values = ['false', 'no', '0', 'False', 'No', 'FALSE', 'NO']


class CSVSaveOptions:
    """
    Options for saving CSV files.

    Attributes:
        delimiter (str): Field delimiter character. Default is ','.
        encoding (str): File encoding. Default is the system's preferred encoding.
        quote_char (str): Character used for quoting fields. Default is '"'.
        quoting (int): Quoting behavior (csv.QUOTE_MINIMAL, csv.QUOTE_ALL, etc.). Default is QUOTE_MINIMAL.
        line_terminator (str): Line ending character(s). Default is '\\r\\n'.
        include_header (bool): Whether to include column headers (if available). Default is False.
        worksheet_index (int): Index of the worksheet to export. Default is 0 (first worksheet).
        date_format (str): Format string for date values. Default is '%Y-%m-%d'.
        datetime_format (str): Format string for datetime values. Default is '%Y-%m-%d %H:%M:%S'.
        time_format (str): Format string for time values. Default is '%H:%M:%S'.
        write_bom (bool): Whether to write UTF-8 BOM at file start. Default is False.

    Examples:
        >>> options = CSVSaveOptions()
        >>> options.delimiter = '\\t'
        >>> options.encoding = 'utf-8'
        >>> CSVHandler.save_csv(workbook, 'output.csv', options)
    """

    def __init__(self):
        self.delimiter = ','
        self.encoding = _system_encoding()
        self.quote_char = '"'
        self.quoting = csv.QUOTE_MINIMAL
        self.line_terminator = '\r\n'
        self.include_header = False
        self.worksheet_index = 0
        self.date_format = '%Y-%m-%d'
        self.datetime_format = '%Y-%m-%d %H:%M:%S'
        self.time_format = '%H:%M:%S'
        self.write_bom = False


class CSVHandler:
    """
    Handles CSV import and export operations for workbooks.

    This class provides static methods to load CSV files into workbooks
    and save workbook data to CSV format.

    Examples:
        # Export workbook to CSV
        >>> wb = Workbook('data.xlsx')
        >>> CSVHandler.save_csv(wb, 'output.csv')

        # Import CSV to workbook
        >>> wb = Workbook()
        >>> CSVHandler.load_csv(wb, 'input.csv')
        >>> print(wb.worksheets[0].cells['A1'].value)
    """

    @staticmethod
    def save_csv(workbook, file_path: str, options: Optional[CSVSaveOptions] = None) -> None:
        """
        Saves a workbook worksheet to a CSV file.

        Args:
            workbook: The Workbook object to export.
            file_path (str): Path where the CSV file should be saved.
            options (CSVSaveOptions, optional): Export options. Uses defaults if None.

        Raises:
            IndexError: If the specified worksheet index is out of range.
            IOError: If the file cannot be written.

        Examples:
            >>> wb = Workbook('data.xlsx')
            >>> CSVHandler.save_csv(wb, 'output.csv')

            >>> options = CSVSaveOptions()
            >>> options.delimiter = ';'
            >>> CSVHandler.save_csv(wb, 'output.csv', options)
        """
        if options is None:
            options = CSVSaveOptions()

        # Validate worksheet index
        if options.worksheet_index >= len(workbook.worksheets):
            raise IndexError(f"Worksheet index {options.worksheet_index} out of range")

        worksheet = workbook.worksheets[options.worksheet_index]

        # Get the used range of the worksheet
        rows_data = CSVHandler._get_worksheet_data(worksheet)

        # Open file for writing
        with open(file_path, 'w', encoding=options.encoding, newline='') as f:
            # Write BOM if requested
            if options.write_bom and options.encoding.lower().replace('-', '') == 'utf8':
                f.write('\ufeff')

            writer = csv.writer(
                f,
                delimiter=options.delimiter,
                quotechar=options.quote_char,
                quoting=options.quoting,
                lineterminator=options.line_terminator
            )

            # Write data rows
            for row_data in rows_data:
                formatted_row = [
                    CSVHandler._format_cell_for_csv(cell, options, workbook, worksheet)
                    for cell in row_data
                ]
                writer.writerow(formatted_row)

    @staticmethod
    def save_csv_to_string(workbook, options: Optional[CSVSaveOptions] = None) -> str:
        """
        Saves a workbook worksheet to a CSV string.

        Args:
            workbook: The Workbook object to export.
            options (CSVSaveOptions, optional): Export options. Uses defaults if None.

        Returns:
            str: The CSV content as a string.

        Examples:
            >>> wb = Workbook('data.xlsx')
            >>> csv_content = CSVHandler.save_csv_to_string(wb)
        """
        if options is None:
            options = CSVSaveOptions()

        # Validate worksheet index
        if options.worksheet_index >= len(workbook.worksheets):
            raise IndexError(f"Worksheet index {options.worksheet_index} out of range")

        worksheet = workbook.worksheets[options.worksheet_index]

        # Get the used range of the worksheet
        rows_data = CSVHandler._get_worksheet_data(worksheet)

        # Write to string buffer
        output = io.StringIO()

        # Write BOM if requested
        if options.write_bom and options.encoding.lower().replace('-', '') == 'utf8':
            output.write('\ufeff')

        writer = csv.writer(
            output,
            delimiter=options.delimiter,
            quotechar=options.quote_char,
            quoting=options.quoting,
            lineterminator=options.line_terminator
        )

        # Write data rows
        for row_data in rows_data:
            formatted_row = [
                CSVHandler._format_cell_for_csv(cell, options, workbook, worksheet)
                for cell in row_data
            ]
            writer.writerow(formatted_row)

        return output.getvalue()

    @staticmethod
    def load_csv(workbook, file_path: str, options: Optional[CSVLoadOptions] = None) -> None:
        """
        Loads a CSV file into a workbook.

        The CSV data is loaded into the first worksheet of the workbook,
        replacing any existing data.

        Args:
            workbook: The Workbook object to load data into.
            file_path (str): Path to the CSV file to load.
            options (CSVLoadOptions, optional): Import options. Uses defaults if None.

        Raises:
            FileNotFoundError: If the CSV file does not exist.
            IOError: If the file cannot be read.

        Examples:
            >>> wb = Workbook()
            >>> CSVHandler.load_csv(wb, 'data.csv')

            >>> options = CSVLoadOptions()
            >>> options.delimiter = ';'
            >>> options.encoding = 'latin-1'
            >>> CSVHandler.load_csv(wb, 'data.csv', options)
        """
        if options is None:
            options = CSVLoadOptions()

        with open(file_path, 'r', encoding=options.encoding, newline='') as f:
            CSVHandler._load_csv_from_reader(workbook, f, options)

    @staticmethod
    def load_csv_from_string(workbook, csv_content: str, options: Optional[CSVLoadOptions] = None) -> None:
        """
        Loads CSV data from a string into a workbook.

        Args:
            workbook: The Workbook object to load data into.
            csv_content (str): The CSV content as a string.
            options (CSVLoadOptions, optional): Import options. Uses defaults if None.

        Examples:
            >>> wb = Workbook()
            >>> csv_data = "Name,Age\\nAlice,30\\nBob,25"
            >>> CSVHandler.load_csv_from_string(wb, csv_data)
        """
        if options is None:
            options = CSVLoadOptions()

        # Remove BOM if present
        if csv_content.startswith('\ufeff'):
            csv_content = csv_content[1:]

        f = io.StringIO(csv_content)
        CSVHandler._load_csv_from_reader(workbook, f, options)

    @staticmethod
    def _load_csv_from_reader(workbook, reader, options: CSVLoadOptions) -> None:
        """
        Internal method to load CSV data from a file-like reader.

        Args:
            workbook: The Workbook object to load data into.
            reader: File-like object to read from.
            options (CSVLoadOptions): Import options.
        """
        # Get the first worksheet (or create one if needed)
        if len(workbook.worksheets) == 0:
            from .worksheet import Worksheet
            workbook._worksheets.append(Worksheet("Sheet1"))

        worksheet = workbook.worksheets[0]

        # Clear existing cell data
        worksheet.cells._cells.clear()

        # Create CSV reader
        csv_reader = csv.reader(
            reader,
            delimiter=options.delimiter,
            quotechar=options.quote_char,
            escapechar=options.escape_char
        )

        # Skip rows if specified
        for _ in range(options.skip_rows):
            try:
                next(csv_reader)
            except StopIteration:
                return

        # Read header if specified
        headers = None
        if options.has_header:
            try:
                headers = next(csv_reader)
                # Write headers to first row
                for col_idx, header in enumerate(headers, start=1):
                    worksheet.cells.cell(row=1, column=col_idx).value = header
            except StopIteration:
                return

        # Starting row (1 if no header, 2 if header)
        start_row = 2 if options.has_header else 1

        # Read data rows
        for row_idx, row in enumerate(csv_reader, start=start_row):
            for col_idx, value in enumerate(row, start=1):
                # Auto-detect types if enabled
                if options.auto_detect_types:
                    value = CSVHandler._parse_value(value, options)

                worksheet.cells.cell(row=row_idx, column=col_idx).value = value

    @staticmethod
    def _get_worksheet_data(worksheet) -> List[List[Any]]:
        """
        Extracts all cell data from a worksheet as a 2D list.

        Args:
            worksheet: The Worksheet object to extract data from.

        Returns:
            list: 2D list of Cell objects (or None), organized by rows.
        """
        from .cells import Cells

        cells_dict = worksheet.cells._cells

        min_row = 1
        min_col = 1
        max_row = 0
        max_col = 0

        if hasattr(worksheet, '_dimension'):
            min_row, min_col, max_row, max_col = worksheet._dimension
        else:
            if not cells_dict:
                return []

            # Find the dimensions by parsing cell references
            for ref in cells_dict.keys():
                row, col = Cells.coordinate_from_string(ref)
                if row > max_row:
                    max_row = row
                if col > max_col:
                    max_col = col

        if max_row == 0 or max_col == 0:
            return []

        # Build 2D array
        rows_data = []
        for row_idx in range(min_row, max_row + 1):
            row_data = []
            for col_idx in range(min_col, max_col + 1):
                ref = Cells.coordinate_to_string(row_idx, col_idx)
                cell = cells_dict.get(ref)
                row_data.append(cell)
            rows_data.append(row_data)

        # Trim trailing rows that are effectively empty, matching Excel's
        # CSV export behavior. A row is considered non-empty if any cell
        # has a value, a formula, or non-default styling (e.g. a non-default font,
        # border, fill, or alignment).
        while len(rows_data) > 0:
            if any(CSVHandler._cell_has_content(c) for c in rows_data[-1]):
                break
            rows_data.pop()

        return rows_data

    @staticmethod
    def _cell_has_content(cell) -> bool:
        """Check if a cell has meaningful content for CSV export purposes."""
        if cell is None:
            return False
        if cell.value is not None:
            return True
        if hasattr(cell, 'formula') and cell.formula is not None:
            return True
        if hasattr(cell, 'style'):
            style = cell.style
            # Check for non-default font
            if hasattr(style, 'font'):
                font = style.font
                if font.name != 'Calibri' or font.size != 11 or font.bold or font.italic:
                    return True
            # Check for non-default borders
            if hasattr(style, 'borders'):
                for side in ('left', 'right', 'top', 'bottom'):
                    border = getattr(style.borders, side, None)
                    if border and hasattr(border, 'line_style') and border.line_style not in (None, 'none'):
                        return True
            # Check for non-default fill
            if hasattr(style, 'fill'):
                fill = style.fill
                if hasattr(fill, 'pattern_type') and fill.pattern_type not in (None, 'none'):
                    return True
            # Check for non-default alignment
            if hasattr(style, 'alignment'):
                align = style.alignment
                if (getattr(align, 'horizontal', 'general') not in ('general', None) or
                    getattr(align, 'vertical', 'bottom') not in ('bottom', None) or
                    getattr(align, 'wrap_text', False)):
                    return True
        return False


    @staticmethod
    def _format_cell_for_csv(cell, options: CSVSaveOptions, workbook=None, worksheet=None) -> str:
        """
        Formats a cell for CSV output, applying number formats when appropriate.

        Args:
            cell: The Cell object (or None).
            options (CSVSaveOptions): Export options.
            workbook: The Workbook object (for formula evaluation).
            worksheet: The Worksheet object (for formula evaluation context).

        Returns:
            str: The formatted string value.
        """
        if cell is None:
            return ''

        number_format = None
        if hasattr(cell, 'style') and hasattr(cell.style, 'number_format'):
            number_format = cell.style.number_format

        # Get the cell value
        value = cell.value

        # If value is None but cell has a formula, try to evaluate it
        if value is None and hasattr(cell, 'formula') and cell.formula is not None:
            if workbook is not None:
                try:
                    from .formula_evaluator import FormulaEvaluator
                    evaluator = FormulaEvaluator(workbook)
                    value = evaluator.evaluate(cell.formula, worksheet)
                except Exception:
                    pass  # If evaluation fails, value remains None

        return CSVHandler._format_value_for_csv(value, options, number_format)

    @staticmethod
    def _format_value_for_csv(value: Any, options: CSVSaveOptions, number_format: Optional[str] = None) -> str:
        """
        Formats a cell value for CSV output.

        Args:
            value: The cell value to format.
            options (CSVSaveOptions): Export options.
            number_format (str, optional): Cell number format string.

        Returns:
            str: The formatted string value.
        """
        if value is None:
            return ''

        # Handle datetime types
        if isinstance(value, datetime):
            return value.strftime(options.datetime_format)

        if isinstance(value, date):
            return value.strftime(options.date_format)

        if isinstance(value, time):
            return value.strftime(options.time_format)

        # Handle boolean
        if isinstance(value, bool):
            return 'TRUE' if value else 'FALSE'

        # Handle numeric types
        if isinstance(value, (int, float)):
            formatted_number = CSVHandler._format_number_with_format(value, number_format)
            if formatted_number is not None:
                return formatted_number
            # Check if it's an integer stored as float
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            # Format float to match Excel's display precision (avoids artifacts like 1.1000000000000001)
            return CSVHandler._format_float_like_excel(value)

        # Handle strings
        return str(value)

    @staticmethod
    def _format_float_like_excel(value: float) -> str:
        """
        Formats a float value to match Excel's default display precision.

        Excel displays up to 15 significant digits and removes trailing zeros.
        This avoids floating-point artifacts like 1.1000000000000001.

        Args:
            value (float): The float value to format.

        Returns:
            str: The formatted float string.
        """
        # Excel's "General" format displays up to 10 significant digits in CSV export.
        # Using 10 significant digits matches Excel's CSV output behaviour.
        formatted = f'{value:.10g}'

        # Handle edge case where very small numbers might show as scientific notation
        # Excel shows these as 0 in CSV
        if 'e' in formatted.lower() and abs(value) < 1e-10:
            return '0'

        return formatted

    @staticmethod
    def _format_number_with_format(value: float, format_code: Optional[str]) -> Optional[str]:
        """
        Formats a numeric value using an Excel-style number format string.

        Args:
            value (float): The numeric value to format.
            format_code (str, optional): The Excel number format code.

        Returns:
            str or None: Formatted value if a usable format is provided, otherwise None.
        """
        if format_code is None:
            return None

        if format_code == '' or format_code.lower() == 'general' or format_code == '@':
            return None

        sections = format_code.split(';')
        section = sections[0] if sections else format_code
        value_to_format = value

        if len(sections) > 1:
            if value < 0:
                section = sections[1]
                value_to_format = abs(value)
            elif value == 0 and len(sections) > 2:
                section = sections[2]
            else:
                section = sections[0]

        # Strip bracketed tokens like [Red], [>100], [$-409]
        section = re.sub(r'\[[^\]]+\]', '', section)

        # Check if this is a date/time format before checking for number placeholders
        if CSVHandler._is_date_format(section):
            return CSVHandler._format_date_with_excel_format(value, section)

        # If no placeholders, return literal section
        if not re.search(r'[0#?]', section):
            return CSVHandler._clean_format_literal(section)

        first_idx = None
        last_idx = None
        for idx, ch in enumerate(section):
            if ch in '0#?':
                if first_idx is None:
                    first_idx = idx
                last_idx = idx

        if first_idx is None:
            return CSVHandler._clean_format_literal(section)

        prefix_raw = section[:first_idx]
        suffix_raw = section[last_idx + 1:]
        prefix = CSVHandler._clean_format_literal(prefix_raw)
        suffix = CSVHandler._clean_format_literal(suffix_raw)

        has_percent = '%' in section
        if has_percent:
            value_to_format *= 100

        # Scientific notation
        if 'E' in section or 'e' in section:
            decimals = 0
            match = re.search(r'\.(?P<frac>[0#?]+)[eE]', section)
            if match:
                decimals = len(match.group('frac'))
            formatted = f"{value_to_format:.{decimals}E}"
            return f"{prefix}{formatted}{suffix}"

        # Standard number format
        number_pattern = section[first_idx:last_idx + 1]
        pattern_clean = re.sub(r'[^0#.,]', '', number_pattern)
        if '.' in pattern_clean:
            int_part, frac_part = pattern_clean.split('.', 1)
        else:
            int_part, frac_part = pattern_clean, ''

        use_grouping = ',' in int_part
        min_decimals = frac_part.count('0')
        max_decimals = sum(1 for ch in frac_part if ch in '0#')

        if max_decimals == 0:
            rounded = int(Decimal(str(value_to_format)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
            formatted = f'{rounded:,}' if use_grouping else str(rounded)
        else:
            quant = Decimal('0.' + '0' * max_decimals)
            rounded_d = Decimal(str(value_to_format)).quantize(quant, rounding=ROUND_HALF_UP)
            rounded_f = float(rounded_d)
            formatted = f'{rounded_f:,.{max_decimals}f}' if use_grouping else f'{rounded_f:.{max_decimals}f}'
            if max_decimals > min_decimals and '.' in formatted:
                int_text, frac_text = formatted.split('.', 1)
                frac_text = frac_text.rstrip('0')
                if len(frac_text) < min_decimals:
                    frac_text = frac_text.ljust(min_decimals, '0')
                formatted = int_text if frac_text == '' else f"{int_text}.{frac_text}"

        return f"{prefix}{formatted}{suffix}"

    @staticmethod
    def _clean_format_literal(text: str) -> str:
        """
        Cleans format literals by removing Excel formatting directives.

        Args:
            text (str): Raw format text.

        Returns:
            str: Cleaned literal text.
        """
        result = []
        idx = 0
        while idx < len(text):
            ch = text[idx]
            if ch == '"':
                idx += 1
                while idx < len(text) and text[idx] != '"':
                    result.append(text[idx])
                    idx += 1
                idx += 1
                continue
            if ch == '_':
                # Underscore means "add a space the width of the following character"
                # In CSV output, we represent this as a single space
                result.append(' ')
                idx += 2
                continue
            if ch == '*':
                # Asterisk means "repeat the following character to fill" - skip in CSV
                idx += 2
                continue
            if ch == '\\':
                if idx + 1 < len(text):
                    result.append(text[idx + 1])
                    idx += 2
                else:
                    idx += 1
                continue
            result.append(ch)
            idx += 1
        return ''.join(result)

    @staticmethod
    def _is_date_format(format_code: str) -> bool:
        """
        Checks if an Excel format code is a date/time format.

        Args:
            format_code (str): The Excel format code (with bracketed tokens already removed).

        Returns:
            bool: True if the format code is a date/time format, False otherwise.
        """
        if not format_code:
            return False

        # Common date/time format characters in Excel
        # y, m, d for date; h, s for time; AM/PM for 12-hour format
        # Note: 'm' can be month or minute depending on context
        date_chars = set('yYdDhHsS')
        format_lower = format_code.lower()

        # Check for date-specific patterns
        if any(c in format_code for c in date_chars):
            return True

        # Check for month patterns (m or mm not adjacent to h or s means month)
        # If 'm' appears and there's also 'y' or 'd', it's definitely a date
        if 'm' in format_lower:
            if 'y' in format_lower or 'd' in format_lower:
                return True
            # If 'm' appears without 'h' or 's', it's likely a date (standalone month)
            if 'h' not in format_lower and 's' not in format_lower:
                return True

        return False

    @staticmethod
    def _format_date_with_excel_format(serial_date: float, format_code: str) -> str:
        """
        Formats an Excel serial date number using an Excel date format code.

        Args:
            serial_date (float): The Excel serial date number.
            format_code (str): The Excel date format code.

        Returns:
            str: The formatted date string.
        """
        # Convert Excel serial date to Python datetime
        # Excel base date is December 30, 1899 (to account for the 1900 leap year bug)
        excel_epoch = datetime(1899, 12, 30)

        try:
            days = int(serial_date)
            fraction = serial_date - days
            seconds = int(round(fraction * 86400))

            from datetime import timedelta
            dt = excel_epoch + timedelta(days=days, seconds=seconds)
        except (ValueError, OverflowError):
            # If conversion fails, return the raw number
            return str(serial_date)

        # Convert Excel format code to Python strftime format
        python_format = CSVHandler._excel_date_format_to_python(format_code)

        try:
            return dt.strftime(python_format)
        except ValueError:
            # If formatting fails, return ISO format
            return dt.strftime('%Y-%m-%d')

    @staticmethod
    def _excel_date_format_to_python(excel_format: str) -> str:
        """
        Converts an Excel date format code to a Python strftime format string.

        Args:
            excel_format (str): The Excel date format code.

        Returns:
            str: The equivalent Python strftime format string.
        """
        result = []
        i = 0
        fmt = excel_format

        while i < len(fmt):
            ch = fmt[i]
            ch_lower = ch.lower()

            # Year patterns
            if ch_lower == 'y':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'y':
                    count += 1
                if count >= 4:
                    result.append('%Y')  # 4-digit year
                else:
                    result.append('%y')  # 2-digit year
                i += count
                continue

            # Month patterns (m or mm when not after h)
            if ch_lower == 'm':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'm':
                    count += 1

                # Determine if this is month or minute based on context
                # Look backward for 'h' - if found, this is minutes
                is_minute = False
                for j in range(i - 1, -1, -1):
                    if fmt[j].lower() == 'h':
                        is_minute = True
                        break
                    if fmt[j].lower() in 'yds':
                        break

                if is_minute:
                    result.append('%M')  # Minutes (always 2-digit)
                elif count >= 4:
                    result.append('%B')  # Full month name
                elif count == 3:
                    result.append('%b')  # Abbreviated month name
                elif count == 2:
                    result.append('%m')  # 2-digit month
                else:
                    result.append('%#m' if _is_windows else '%-m')  # 1 or 2-digit month
                i += count
                continue

            # Day patterns
            if ch_lower == 'd':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'd':
                    count += 1
                if count >= 4:
                    result.append('%A')  # Full weekday name
                elif count == 3:
                    result.append('%a')  # Abbreviated weekday name
                elif count == 2:
                    result.append('%d')  # 2-digit day
                else:
                    result.append('%#d' if _is_windows else '%-d')  # 1 or 2-digit day
                i += count
                continue

            # Hour patterns
            if ch_lower == 'h':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 'h':
                    count += 1
                # Check if AM/PM is present in format for 12-hour vs 24-hour
                if 'am' in fmt.lower() or 'pm' in fmt.lower():
                    result.append('%I')  # 12-hour format
                else:
                    result.append('%H')  # 24-hour format
                i += count
                continue

            # Second patterns
            if ch_lower == 's':
                count = 1
                while i + count < len(fmt) and fmt[i + count].lower() == 's':
                    count += 1
                result.append('%S')
                i += count
                continue

            # AM/PM patterns
            if ch_lower == 'a' and i + 1 < len(fmt) and fmt[i + 1].lower() == 'm':
                # Check for AM/PM or A/P
                if i + 3 < len(fmt) and fmt[i:i+4].lower() == 'am/p' and fmt[i+4].lower() == 'm':
                    result.append('%p')
                    i += 5
                elif i + 1 < len(fmt) and fmt[i:i+2].lower() == 'am':
                    result.append('%p')
                    i += 2
                else:
                    result.append(ch)
                    i += 1
                continue

            # Skip escaped characters
            if ch == '\\' and i + 1 < len(fmt):
                result.append(fmt[i + 1])
                i += 2
                continue

            # Skip quoted strings
            if ch == '"':
                i += 1
                while i < len(fmt) and fmt[i] != '"':
                    result.append(fmt[i])
                    i += 1
                i += 1  # Skip closing quote
                continue

            # Pass through other characters (like /, -, :, space)
            result.append(ch)
            i += 1

        return ''.join(result)

    @staticmethod
    def _parse_value(value_str: str, options: CSVLoadOptions) -> Any:
        """
        Parses a string value and attempts to convert it to the appropriate type.

        Args:
            value_str (str): The string value to parse.
            options (CSVLoadOptions): Import options containing type detection settings.

        Returns:
            The parsed value (int, float, bool, datetime, date, or str).
        """
        # Handle empty strings
        if not value_str or value_str.strip() == '':
            return None

        value_str = value_str.strip()

        # Try boolean
        if value_str in options.true_values:
            return True
        if value_str in options.false_values:
            return False

        # Try integer
        try:
            # Check for integer format (no decimal point, no exponential)
            if '.' not in value_str and 'e' not in value_str.lower():
                return int(value_str)
        except ValueError:
            pass

        # Try float
        try:
            return float(value_str)
        except ValueError:
            pass

        # Try date/datetime formats
        for fmt in options.date_formats:
            try:
                dt = datetime.strptime(value_str, fmt)
                # Return date if no time component, datetime otherwise
                if '%H' in fmt or '%I' in fmt:
                    return dt
                else:
                    return dt.date()
            except ValueError:
                continue

        # Return as string
        return value_str


def load_csv_workbook(file_path: str, options: Optional[CSVLoadOptions] = None):
    """
    Convenience function to create a new Workbook from a CSV file.

    Args:
        file_path (str): Path to the CSV file to load.
        options (CSVLoadOptions, optional): Import options. Uses defaults if None.

    Returns:
        Workbook: A new Workbook containing the CSV data.

    Examples:
        >>> wb = load_csv_workbook('data.csv')
        >>> print(wb.worksheets[0].cells['A1'].value)
    """
    from .workbook import Workbook

    wb = Workbook()
    CSVHandler.load_csv(wb, file_path, options)
    return wb


def save_workbook_as_csv(workbook, file_path: str, options: Optional[CSVSaveOptions] = None) -> None:
    """
    Convenience function to save a Workbook to a CSV file.

    Args:
        workbook: The Workbook object to export.
        file_path (str): Path where the CSV file should be saved.
        options (CSVSaveOptions, optional): Export options. Uses defaults if None.

    Examples:
        >>> wb = Workbook('data.xlsx')
        >>> save_workbook_as_csv(wb, 'output.csv')
    """
    CSVHandler.save_csv(workbook, file_path, options)
