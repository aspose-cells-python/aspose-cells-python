"""
File format definitions and utilities for Excel workbook operations.
"""

from typing import Union
from pathlib import Path
from enum import Enum
from datetime import datetime

# Define common cell value types
CellValue = Union[str, int, float, bool, datetime, None]


class FileFormat(Enum):
    """File format enumeration for save and export operations."""
    
    XLSX = "xlsx"
    CSV = "csv"
    JSON = "json"
    MARKDOWN = "markdown"
    
    @classmethod
    def from_extension(cls, filename: Union[str, Path]) -> 'FileFormat':
        """Infer format from file extension."""
        ext = Path(filename).suffix.lower()
        format_map = {
            '.xlsx': cls.XLSX,
            '.csv': cls.CSV,
            '.json': cls.JSON,
            '.md': cls.MARKDOWN,
            '.markdown': cls.MARKDOWN,
        }
        return format_map.get(ext, cls.XLSX)
    
    @classmethod
    def get_supported_formats(cls) -> list['FileFormat']:
        """Get list of all supported file formats."""
        return list(cls)
    
    @property
    def extension(self) -> str:
        """Get file extension for this format."""
        extension_map = {
            FileFormat.XLSX: '.xlsx',
            FileFormat.CSV: '.csv',
            FileFormat.JSON: '.json',
            FileFormat.MARKDOWN: '.md'
        }
        return extension_map.get(self, '.xlsx')
    
    @property
    def mime_type(self) -> str:
        """Get MIME type for this format."""
        mime_map = {
            self.XLSX: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            self.CSV: 'text/csv',
            self.JSON: 'application/json',
            self.MARKDOWN: 'text/markdown'
        }
        return mime_map.get(self, 'application/octet-stream')


class ConversionOptions:
    """Options for format conversion operations."""
    
    def __init__(self, 
                 sheet_name: str = None,
                 include_headers: bool = True,
                 all_sheets: bool = False,
                 max_col_width: int = 50,
                 table_alignment: str = 'left',
                 preserve_formatting: bool = False,
                 **kwargs):
        self.sheet_name = sheet_name
        self.include_headers = include_headers
        self.all_sheets = all_sheets
        self.max_col_width = max_col_width
        self.table_alignment = table_alignment
        self.preserve_formatting = preserve_formatting
        self.extra_options = kwargs