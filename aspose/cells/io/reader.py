"""
Excel file reader maintaining OOXML protocol compatibility.
"""

from typing import TYPE_CHECKING
from pathlib import Path

from ..utils import FileFormatError

if TYPE_CHECKING:
    from ..workbook import Workbook


class ExcelReader:
    """Excel file reader with OOXML protocol support."""
    
    def __init__(self):
        pass
    
    def load_workbook(self, workbook: 'Workbook', filename: str):
        """Load Excel file into workbook object."""
        from .xlsx.reader import XlsxReader
        
        # Check file extension and delegate to appropriate reader
        file_path = Path(filename)
        if file_path.suffix.lower() in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            xlsx_reader = XlsxReader()
            xlsx_reader.load_workbook(workbook, filename)
        else:
            raise FileFormatError(f"Unsupported file format: {file_path.suffix}")
    
    
    
    
    
