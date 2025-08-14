"""
Workbook implementation with unified API and multiple file format support.
"""

from typing import Dict, List, Optional, Union
from pathlib import Path

from .worksheet import Worksheet
from .formats import FileFormat, ConversionOptions
from .utils import (
    sanitize_sheet_name,
    WorksheetNotFoundError,
    FileFormatError,
    ExportError
)




class WorksheetCollection:
    """Collection manager for worksheets with multiple access patterns."""
    
    def __init__(self, workbook: 'Workbook'):
        self._workbook = workbook
    
    def add(self, name: str) -> Worksheet:
        """Add new worksheet with specified name."""
        clean_name = sanitize_sheet_name(name)
        if clean_name in self._workbook._worksheets:
            # Generate unique name
            counter = 1
            base_name = clean_name
            while clean_name in self._workbook._worksheets:
                clean_name = f"{base_name}_{counter}"
                counter += 1
        
        worksheet = Worksheet(self._workbook, clean_name)
        self._workbook._worksheets[clean_name] = worksheet
        return worksheet
    
    def remove(self, name: Union[str, int, Worksheet]):
        """Remove worksheet by name, index, or object."""
        if isinstance(name, Worksheet):
            name = name.name
        elif isinstance(name, int):
            sheets = list(self._workbook._worksheets.values())
            if 0 <= name < len(sheets):
                name = sheets[name].name
            else:
                raise WorksheetNotFoundError(f"Worksheet index {name} out of range")
        
        if name not in self._workbook._worksheets:
            raise WorksheetNotFoundError(f"Worksheet '{name}' not found")
        
        # Don't allow removing the last worksheet
        if len(self._workbook._worksheets) <= 1:
            raise WorksheetNotFoundError("Cannot remove the last worksheet")
        
        # Update active sheet if necessary
        if self._workbook._active_sheet and self._workbook._active_sheet.name == name:
            remaining_sheets = [ws for ws in self._workbook._worksheets.values() if ws.name != name]
            self._workbook._active_sheet = remaining_sheets[0]
        
        del self._workbook._worksheets[name]
    
    def __getitem__(self, key: Union[str, int]) -> Worksheet:
        """Get worksheet by name or index."""
        if isinstance(key, str):
            if key not in self._workbook._worksheets:
                raise WorksheetNotFoundError(f"Worksheet '{key}' not found")
            return self._workbook._worksheets[key]
        elif isinstance(key, int):
            sheets = list(self._workbook._worksheets.values())
            if 0 <= key < len(sheets):
                return sheets[key]
            else:
                raise WorksheetNotFoundError(f"Worksheet index {key} out of range")
        else:
            raise WorksheetNotFoundError(f"Invalid worksheet key: {key}")
    
    def __len__(self) -> int:
        """Number of worksheets."""
        return len(self._workbook._worksheets)
    
    def __iter__(self):
        """Iterate over worksheets."""
        return iter(self._workbook._worksheets.values())
    
    def __contains__(self, name: str) -> bool:
        """Check if worksheet exists."""
        return name in self._workbook._worksheets


class Workbook:
    """Excel workbook with unified API and multiple access patterns."""
    
    def __init__(self, filename: Optional[Union[str, Path]] = None):
        self._filename: Optional[Path] = None
        self._worksheets: Dict[str, Worksheet] = {}
        self._active_sheet: Optional[Worksheet] = None
        self._shared_strings: List[str] = []
        self._properties: Dict[str, Union[str, int, float, bool]] = {}
        
        # Initialize with default worksheet
        default_sheet = Worksheet(self, "Sheet1")
        self._worksheets["Sheet1"] = default_sheet
        self._active_sheet = default_sheet
        
        if filename:
            self._load_from_file(filename)
    
    @classmethod
    def load(cls, filename: Union[str, Path]) -> 'Workbook':
        """Load workbook from file."""
        return cls(filename)
    
    @property
    def active(self) -> Worksheet:
        """Get active worksheet."""
        if self._active_sheet is None and self._worksheets:
            self._active_sheet = next(iter(self._worksheets.values()))
        return self._active_sheet
    
    @active.setter
    def active(self, value: Union[Worksheet, str, int]):
        """Set active worksheet by object, name, or index."""
        if isinstance(value, Worksheet):
            if value in self._worksheets.values():
                self._active_sheet = value
            else:
                raise WorksheetNotFoundError("Worksheet not in this workbook")
        elif isinstance(value, str):
            if value in self._worksheets:
                self._active_sheet = self._worksheets[value]
            else:
                raise WorksheetNotFoundError(f"Worksheet '{value}' not found")
        elif isinstance(value, int):
            sheets = list(self._worksheets.values())
            if 0 <= value < len(sheets):
                self._active_sheet = sheets[value]
            else:
                raise WorksheetNotFoundError(f"Worksheet index {value} out of range")
        else:
            raise WorksheetNotFoundError(f"Invalid active sheet value: {value}")
    
    @property
    def worksheets(self) -> WorksheetCollection:
        """Get worksheet collection manager."""
        return WorksheetCollection(self)
    
    @property
    def sheetnames(self) -> List[str]:
        """Get list of worksheet names."""
        return list(self._worksheets.keys())
    
    def create_sheet(self, name: str = None, index: int = None) -> Worksheet:
        """Create new worksheet with optional name and position."""
        if name is None:
            # Generate default name
            counter = len(self._worksheets) + 1
            while f"Sheet{counter}" in self._worksheets:
                counter += 1
            name = f"Sheet{counter}"
        
        worksheet = self.worksheets.add(name)
        
        # Handle index positioning if specified
        if index is not None and 0 <= index < len(self._worksheets):
            # Re-order worksheets to insert at specific position
            sheet_items = list(self._worksheets.items())
            # Remove the newly added sheet from its current position (last)
            new_sheet_item = sheet_items.pop()
            # Insert it at the specified index
            sheet_items.insert(index, new_sheet_item)
            # Rebuild the ordered dictionary
            self._worksheets.clear()
            for sheet_name, sheet_obj in sheet_items:
                self._worksheets[sheet_name] = sheet_obj
        
        return worksheet
    
    def _load_from_file(self, filename: Union[str, Path]):
        """Load workbook from Excel file."""
        self._filename = Path(filename)
        
        if not self._filename.exists():
            raise FileFormatError(f"File not found: {filename}")
        
        if self._filename.suffix.lower() not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            raise FileFormatError(f"Unsupported file format: {self._filename.suffix}")
        
        # Import reader here to avoid circular imports
        from .io.reader import ExcelReader
        
        reader = ExcelReader()
        reader.load_workbook(self, str(filename))
    
    def save(self, filename: Optional[Union[str, Path]] = None, 
             format: Optional[Union[str, FileFormat]] = None, **kwargs):
        """Save workbook to file with specified format."""
        if filename is None:
            if self._filename is None:
                raise FileFormatError("No filename specified and no previous filename available")
            filename = self._filename
        else:
            filename = Path(filename)
        
        if format is None:
            format = FileFormat.from_extension(filename)
        elif isinstance(format, str):
            # Convert string to FileFormat enum
            try:
                format = FileFormat(format)
            except ValueError:
                raise FileFormatError(f"Unsupported format: {format}")
        
        # Import writer here to avoid circular imports
        from .io.writer import ExcelWriter
        
        writer = ExcelWriter()
        writer.save_workbook(self, str(filename), format.value, **kwargs)
        self._filename = Path(filename)
    
    def exportAs(self, format: Union[str, FileFormat], **kwargs) -> str:
        """Export workbook as string in specified format."""
        # Convert string to FileFormat enum if needed
        if isinstance(format, str):
            try:
                format_enum = FileFormat(format)
            except ValueError:
                raise ExportError(f"Unsupported export format: {format}")
        else:
            format_enum = format
        
        if format_enum == FileFormat.JSON:
            from .converters.json_converter import JsonConverter
            converter = JsonConverter()
            return converter.convert_workbook(self, **kwargs)
        elif format_enum == FileFormat.CSV:
            from .converters.csv_converter import CsvConverter
            converter = CsvConverter()
            return converter.convert_workbook(self, **kwargs)
        elif format_enum == FileFormat.MARKDOWN:
            from .converters.markdown_converter import MarkdownConverter
            converter = MarkdownConverter()
            return converter.convert_workbook(self, **kwargs)
        else:
            raise ExportError(f"Unsupported export format: {format_enum.value}")
    
    def copy_worksheet(self, from_worksheet: Union[Worksheet, str]) -> Worksheet:
        """Create a copy of existing worksheet."""
        if isinstance(from_worksheet, str):
            if from_worksheet not in self._worksheets:
                raise WorksheetNotFoundError(f"Source worksheet '{from_worksheet}' not found")
            source = self._worksheets[from_worksheet]
        else:
            source = from_worksheet
        
        # Generate new name
        base_name = f"Copy of {source.name}"
        new_name = base_name
        counter = 1
        while new_name in self._worksheets:
            new_name = f"{base_name} ({counter})"
            counter += 1
        
        # Create new worksheet
        new_worksheet = self.create_sheet(new_name)
        
        # Copy all cell data and formatting
        for coord, cell in source._cells.items():
            row, col = coord
            new_cell = new_worksheet.cell(row, col, cell.value)
            if cell._style:
                new_cell._style = cell._style.copy()
            new_cell._number_format = cell._number_format
            new_cell._hyperlink = cell._hyperlink
            new_cell._comment = cell._comment
        
        # Copy other properties
        new_worksheet._merged_ranges = source._merged_ranges.copy()
        new_worksheet._row_heights = source._row_heights.copy()
        new_worksheet._column_widths = source._column_widths.copy()
        new_worksheet._freeze_panes = source._freeze_panes
        
        return new_worksheet
    
    def close(self):
        """Close workbook and release resources."""
        self._worksheets.clear()
        self._active_sheet = None
        self._shared_strings.clear()
        self._properties.clear()
    
    @property
    def properties(self) -> Dict[str, Union[str, int, float, bool]]:
        """Get workbook properties."""
        return self._properties
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
    
    def __str__(self) -> str:
        """String representation."""
        return f"Workbook({len(self._worksheets)} sheets)"
    
    def __repr__(self) -> str:
        """Debug representation."""
        return f"Workbook(sheets={list(self._worksheets.keys())}, active='{self.active.name if self.active else None}')"