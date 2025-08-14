"""
CSV file reader for loading CSV data into workbook format.
"""

import csv
from typing import Dict, List, Optional, Union, TYPE_CHECKING
from pathlib import Path
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook


class CsvReader:
    """Reader for CSV files."""
    
    def __init__(self):
        pass
    
    def read(self, file_path: str, **kwargs) -> List[List[CellValue]]:
        """Read CSV file and return data as list of rows."""
        delimiter = kwargs.get('delimiter', ',')
        quotechar = kwargs.get('quotechar', '"')
        encoding = kwargs.get('encoding', 'utf-8')
        has_header = kwargs.get('has_header', False)
        
        try:
            with open(file_path, 'r', encoding=encoding, newline='') as file:
                reader = csv.reader(file, delimiter=delimiter, quotechar=quotechar)
                
                data = []
                for row in reader:
                    # Convert each cell value
                    converted_row = []
                    for cell in row:
                        converted_row.append(self._convert_cell_value(cell))
                    data.append(converted_row)
                
                return data
                
        except FileNotFoundError:
            raise FileNotFoundError(f"CSV file not found: {file_path}")
        except Exception as e:
            raise ValueError(f"Error reading CSV file: {e}")
    
    def _convert_cell_value(self, value: str) -> CellValue:
        """Convert string value to appropriate Python type."""
        if not value or value.strip() == "":
            return None
        
        value = value.strip()
        
        # Try boolean first
        if value.upper() in ('TRUE', 'FALSE'):
            return value.upper() == 'TRUE'
        
        # Try integer
        try:
            if '.' not in value and 'e' not in value.lower():
                return int(value)
        except ValueError:
            pass
        
        # Try float
        try:
            return float(value)
        except ValueError:
            pass
        
        # Return as string
        return value