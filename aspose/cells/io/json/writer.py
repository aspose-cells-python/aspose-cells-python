"""
JSON file writer for saving workbook data to JSON format.
"""

import json
from typing import Dict, List, Optional, Union, TYPE_CHECKING
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook
    from ...worksheet import Worksheet


class JsonWriter:
    """Writer for JSON files."""
    
    def __init__(self):
        pass
    
    def write(self, file_path: str, data: Union[List[Dict], Dict], **kwargs) -> None:
        """Write data to JSON file."""
        pretty_print = kwargs.get('pretty_print', False)
        encoding = kwargs.get('encoding', 'utf-8')
        
        try:
            with open(file_path, 'w', encoding=encoding) as file:
                if pretty_print:
                    json.dump(data, file, indent=2, ensure_ascii=False)
                else:
                    json.dump(data, file, ensure_ascii=False)
                    
        except Exception as e:
            raise ValueError(f"Error writing JSON file: {e}")
    
    def write_workbook(self, file_path: str, workbook: 'Workbook', **kwargs) -> None:
        """Write workbook data to JSON file."""
        include_empty_cells = kwargs.get('include_empty_cells', False)
        all_sheets = kwargs.get('all_sheets', False)
        sheet_name = kwargs.get('sheet_name')
        
        if sheet_name:
            # Export specific sheet
            if sheet_name in workbook._worksheets:
                worksheet = workbook._worksheets[sheet_name]
                result = self._convert_worksheet(worksheet, include_empty_cells)
            else:
                result = []
        elif all_sheets:
            # Export all sheets with sheet names as keys
            result = {}
            for name, worksheet in workbook._worksheets.items():
                sheet_data = self._convert_worksheet(worksheet, include_empty_cells)
                result[name] = sheet_data
        else:
            # Export only active sheet as simple list
            result = self._convert_worksheet(workbook.active, include_empty_cells)
        
        self.write(file_path, result, **kwargs)
    
    def _convert_worksheet(self, worksheet: 'Worksheet', include_empty_cells: bool = False) -> List[Dict[str, Union[str, int, float, bool, None]]]:
        """Convert worksheet to list of row dictionaries."""
        if not worksheet._cells and not include_empty_cells:
            return []
        
        # Find actual data bounds
        if not worksheet._cells:
            return []
        
        max_row = worksheet.max_row
        max_col = worksheet.max_column
        
        if max_row == 0 or max_col == 0:
            return []
        
        # Convert to list of dictionaries
        result = []
        
        # Generate column headers (A, B, C, etc.)
        headers = []
        for col in range(1, max_col + 1):
            col_name = ""
            temp_col = col - 1
            while temp_col >= 0:
                col_name = chr(ord('A') + (temp_col % 26)) + col_name
                temp_col = temp_col // 26 - 1
            headers.append(col_name)
        
        # Process all data rows
        for row in range(1, max_row + 1):
            row_data = {}
            has_data = False
            
            for col in range(1, max_col + 1):
                cell = worksheet._cells.get((row, col))
                header = headers[col - 1] if col <= len(headers) else f"Column{col}"
                
                if cell and cell.value is not None:
                    row_data[header] = self._convert_cell_value(cell.value)
                    has_data = True
                elif include_empty_cells:
                    row_data[header] = None
            
            if has_data or include_empty_cells:
                result.append(row_data)
        
        return result
    
    def _convert_cell_value(self, value: CellValue) -> Union[str, int, float, bool, None]:
        """Convert cell value to JSON-serializable format."""
        if value is None:
            return None
        elif isinstance(value, (str, int, float, bool)):
            return value
        else:
            return str(value)
    
    def save_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """Save workbook to JSON file - unified interface method."""
        self.write_workbook(file_path, workbook, **options)