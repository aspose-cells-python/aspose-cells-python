"""
JSON file reader for loading JSON data into workbook format.
"""

import json
from typing import Dict, List, Optional, Union, Any, TYPE_CHECKING
from pathlib import Path
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook


class JsonReader:
    """Reader for JSON files."""
    
    def __init__(self):
        pass
    
    def read(self, file_path: str, **kwargs) -> Union[List[List[CellValue]], Dict[str, List[List[CellValue]]]]:
        """Read JSON file and return data in tabular format."""
        encoding = kwargs.get('encoding', 'utf-8')
        
        try:
            with open(file_path, 'r', encoding=encoding) as file:
                data = json.load(file)
            
            return self._convert_json_to_tabular(data)
                
        except FileNotFoundError:
            raise FileNotFoundError(f"JSON file not found: {file_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {e}")
        except Exception as e:
            raise ValueError(f"Error reading JSON file: {e}")
    
    def _convert_json_to_tabular(self, data: Any) -> Union[List[List[CellValue]], Dict[str, List[List[CellValue]]]]:
        """Convert JSON data to tabular format."""
        if isinstance(data, dict):
            # Check if it's a multi-sheet format (keys are sheet names)
            if all(isinstance(v, list) for v in data.values()):
                result = {}
                for sheet_name, sheet_data in data.items():
                    result[sheet_name] = self._convert_list_to_rows(sheet_data)
                return result
            else:
                # Single object, convert to single row
                return self._convert_dict_to_rows(data)
        elif isinstance(data, list):
            return self._convert_list_to_rows(data)
        else:
            # Single value, create single cell
            return [[self._convert_value(data)]]
    
    def _convert_list_to_rows(self, data_list: List[Any]) -> List[List[CellValue]]:
        """Convert list of objects/values to rows."""
        if not data_list:
            return []
        
        if isinstance(data_list[0], dict):
            # List of objects - create header from keys
            headers = list(data_list[0].keys()) if data_list else []
            rows = [headers]  # Add header row
            
            for item in data_list:
                row = []
                for header in headers:
                    value = item.get(header) if isinstance(item, dict) else None
                    row.append(self._convert_value(value))
                rows.append(row)
            
            return rows
        else:
            # List of simple values - convert to single column
            return [[self._convert_value(item)] for item in data_list]
    
    def _convert_dict_to_rows(self, data_dict: Dict[str, Any]) -> List[List[CellValue]]:
        """Convert dictionary to rows (key-value pairs)."""
        rows = []
        for key, value in data_dict.items():
            if isinstance(value, (list, dict)):
                # Complex value, convert to JSON string
                value_str = json.dumps(value, ensure_ascii=False)
                rows.append([key, value_str])
            else:
                rows.append([key, self._convert_value(value)])
        return rows
    
    def _convert_value(self, value: Any) -> CellValue:
        """Convert JSON value to appropriate cell value."""
        if value is None:
            return None
        elif isinstance(value, bool):
            return value
        elif isinstance(value, (int, float)):
            return value
        elif isinstance(value, str):
            return value
        else:
            # Complex types, convert to JSON string
            return json.dumps(value, ensure_ascii=False)
    
    def load_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """Load JSON file into workbook object."""
        data = self.read(file_path, **options)
        
        # Clear existing worksheets
        workbook._worksheets.clear()
        workbook._active_sheet = None
        
        if isinstance(data, dict):
            # Multi-sheet format
            for sheet_name, sheet_rows in data.items():
                worksheet = workbook.create_sheet(sheet_name)
                self._populate_worksheet(worksheet, sheet_rows)
        else:
            # Single sheet format
            worksheet = workbook.create_sheet("Sheet1")
            self._populate_worksheet(worksheet, data)
    
    def _populate_worksheet(self, worksheet, rows: List[List[CellValue]]) -> None:
        """Populate worksheet with row data."""
        for row_idx, row_data in enumerate(rows, 1):
            for col_idx, cell_value in enumerate(row_data, 1):
                if cell_value is not None:
                    worksheet.cell(row_idx, col_idx, cell_value)