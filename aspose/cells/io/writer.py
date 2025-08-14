"""
Excel file writer maintaining OOXML protocol compliance with proper styling.
"""

from typing import TYPE_CHECKING
from pathlib import Path

from ..utils import FileFormatError

if TYPE_CHECKING:
    from ..workbook import Workbook, Worksheet
    from ..cell import Cell




class ExcelWriter:
    """Excel file writer with OOXML protocol support and proper styling."""
    
    def __init__(self):
        pass
    
    def save_workbook(self, workbook: 'Workbook', filename: str, format: str = 'xlsx', **kwargs):
        """Save workbook to file in specified format."""
        if format == 'xlsx':
            from .xlsx.writer import XlsxWriter
            xlsx_writer = XlsxWriter()
            xlsx_writer.save_workbook(workbook, filename, **kwargs)
        elif format == 'csv':
            self._save_csv(workbook, filename, **kwargs)
        else:
            raise FileFormatError(f"Unsupported save format: {format}")
    
    
    
    def _save_csv(self, workbook: 'Workbook', filename: str, **kwargs):
        """Save active sheet as CSV."""
        active_sheet = workbook.active
        if not active_sheet:
            return
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            for row in active_sheet.rows:
                values = []
                for cell in row:
                    value = cell.value
                    if value is None:
                        values.append("")
                    elif isinstance(value, str) and (',' in value or '"' in value or '\n' in value):
                        # CSV escaping
                        escaped = value.replace('"', '""')
                        values.append(f'"{escaped}"')
                    else:
                        values.append(str(value))
                f.write(','.join(values) + '\n')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
