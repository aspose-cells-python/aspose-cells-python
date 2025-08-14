"""
JSON converter for exporting Excel data to JSON format.
"""

import json
from typing import Dict, List, Optional, Union, TYPE_CHECKING
from ..io.json import JsonWriter

if TYPE_CHECKING:
    from ..workbook import Workbook


class JsonConverter:
    """Convert Excel workbook data to JSON format."""
    
    def __init__(self):
        self._writer = JsonWriter()
    
    def convert_workbook(self, workbook: 'Workbook', **kwargs) -> str:
        """Convert entire workbook to JSON string."""
        pretty_print = kwargs.get('pretty_print', False)
        include_empty_cells = kwargs.get('include_empty_cells', False)
        all_sheets = kwargs.get('all_sheets', False)
        sheet_name = kwargs.get('sheet_name')
        
        if sheet_name:
            # Export specific sheet
            if sheet_name in workbook._worksheets:
                worksheet = workbook._worksheets[sheet_name]
                result = self._writer._convert_worksheet(worksheet, include_empty_cells)
            else:
                result = []
        elif all_sheets:
            # Export all sheets with sheet names as keys
            result = {}
            for name, worksheet in workbook._worksheets.items():
                sheet_data = self._writer._convert_worksheet(worksheet, include_empty_cells)
                result[name] = sheet_data
        else:
            # Export only active sheet as simple list
            result = self._writer._convert_worksheet(workbook.active, include_empty_cells)
        
        if pretty_print:
            return json.dumps(result, indent=2, ensure_ascii=False)
        else:
            return json.dumps(result, ensure_ascii=False)