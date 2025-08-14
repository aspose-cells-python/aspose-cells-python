"""
CSV file writer for saving workbook data to CSV format.
"""

import csv
import io
from typing import List, Optional, TYPE_CHECKING
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook
    from ...worksheet import Worksheet


class CsvWriter:
    """Writer for CSV files."""
    
    def __init__(self):
        pass
    
    def write(self, file_path: str, data: List[List[CellValue]], **kwargs) -> None:
        """Write data to CSV file."""
        delimiter = kwargs.get('delimiter', ',')
        quotechar = kwargs.get('quotechar', '"')
        encoding = kwargs.get('encoding', 'utf-8')
        
        try:
            with open(file_path, 'w', newline='', encoding=encoding) as file:
                writer = csv.writer(file, delimiter=delimiter, quotechar=quotechar, 
                                   quoting=csv.QUOTE_MINIMAL)
                
                for row in data:
                    formatted_row = []
                    for cell in row:
                        formatted_row.append(self._format_cell_value(cell))
                    writer.writerow(formatted_row)
                    
        except Exception as e:
            raise ValueError(f"Error writing CSV file: {e}")
    
    def write_workbook(self, file_path: str, workbook: 'Workbook', **kwargs) -> None:
        """Write workbook data to CSV file."""
        sheet_name = kwargs.get('sheet_name')
        
        # Get target worksheet
        if sheet_name and sheet_name in workbook._worksheets:
            worksheet = workbook._worksheets[sheet_name]
        else:
            worksheet = workbook.active
        
        if not worksheet or not worksheet._cells:
            # Write empty file
            with open(file_path, 'w', newline='', encoding=kwargs.get('encoding', 'utf-8')) as file:
                pass
            return
        
        # Convert worksheet to data
        data = self._worksheet_to_data(worksheet)
        self.write(file_path, data, **kwargs)
    
    def _worksheet_to_data(self, worksheet: 'Worksheet') -> List[List[CellValue]]:
        """Convert worksheet to list of rows."""
        max_row = worksheet.max_row
        max_col = worksheet.max_column
        
        if max_row == 0 or max_col == 0:
            return []
        
        data = []
        for row in range(1, max_row + 1):
            row_data = []
            for col in range(1, max_col + 1):
                cell = worksheet._cells.get((row, col))
                if cell and cell.value is not None:
                    row_data.append(cell.value)
                else:
                    row_data.append(None)
            
            # Skip completely empty rows unless they're in the middle
            if any(val is not None for val in row_data) or row < max_row:
                data.append(row_data)
        
        return data
    
    def _format_cell_value(self, value: CellValue) -> str:
        """Format cell value for CSV output."""
        if value is None:
            return ""
        elif isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        elif isinstance(value, (int, float)):
            return str(value)
        else:
            return str(value)