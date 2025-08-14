"""
Worksheet implementation with Pythonic cell access and data operations.
"""

from typing import Dict, Iterator, List, Optional, Tuple, Union, TYPE_CHECKING
from .cell import Cell
from .range import Range
from .formats import CellValue
from .utils import (
    coordinate_to_tuple,
    sanitize_sheet_name,
    InvalidCoordinateError,
    WorksheetNotFoundError
)

if TYPE_CHECKING:
    from .workbook import Workbook


class Worksheet:
    """Excel worksheet with multiple access patterns and batch operations."""
    
    def __init__(self, parent: 'Workbook', name: str):
        self._parent = parent
        self._name = sanitize_sheet_name(name)
        self._cells: Dict[Tuple[int, int], Cell] = {}
        self._max_row = 0
        self._max_column = 0
        self._merged_ranges: set = set()
        self._row_heights: Dict[int, float] = {}
        self._column_widths: Dict[int, float] = {}
        self._hidden_rows: set = set()
        self._hidden_columns: set = set()
        self._freeze_panes: Optional[str] = None
    
    @property
    def name(self) -> str:
        """Worksheet name."""
        return self._name
    
    @property
    def workbook(self) -> 'Workbook':
        """Parent workbook."""
        return self._parent
    
    @name.setter
    def name(self, value: str):
        """Set worksheet name with validation."""
        new_name = sanitize_sheet_name(value)
        if new_name in self._parent._worksheets and new_name != self._name:
            raise WorksheetNotFoundError(f"Worksheet '{new_name}' already exists")
        
        # Update parent's worksheet mapping
        if self._name in self._parent._worksheets:
            del self._parent._worksheets[self._name]
        self._parent._worksheets[new_name] = self
        self._name = new_name
    
    @property
    def max_row(self) -> int:
        """Maximum row with data."""
        return self._max_row
    
    @property
    def max_column(self) -> int:
        """Maximum column with data."""
        return self._max_column
    
    @property
    def dimensions(self) -> str:
        """Range representing used area."""
        if self._max_row == 0 or self._max_column == 0:
            return "A1:A1"
        
        from .utils import tuple_to_coordinate
        start = tuple_to_coordinate(1, 1)
        end = tuple_to_coordinate(self._max_row, self._max_column)
        return f"{start}:{end}"
    
    def _update_bounds(self, row: int, column: int):
        """Update worksheet bounds when cell is modified."""
        self._max_row = max(self._max_row, row)
        self._max_column = max(self._max_column, column)
    
    def __getitem__(self, key: Union[str, Tuple[int, int]]) -> Union[Cell, Range]:
        """Access cell or range using Excel coordinates or 0-based tuples."""
        if isinstance(key, str):
            if ':' in key:
                # Range access: ws['A1:C3']
                return Range(self, key)
            else:
                # Single cell: ws['A1'] (Excel-style)
                row, col = coordinate_to_tuple(key)
                return self.cell(row, col)
        elif isinstance(key, tuple) and len(key) == 2:
            # 0-based tuple access: ws[0, 0] -> A1
            row, col = key
            return self.cell(row + 1, col + 1)  # Convert to 1-based internally
        else:
            raise InvalidCoordinateError(f"Invalid cell access pattern: {key}")
    
    def __setitem__(self, key: Union[str, Tuple[int, int]], value: CellValue):
        """Set cell value using Excel coordinates or 0-based tuples."""
        if isinstance(key, str) and ':' in key:
            # Range assignment: ws['A1:C3'] = data
            range_obj = Range(self, key)
            range_obj.values = value
        elif isinstance(key, tuple) and len(key) == 2:
            # 0-based tuple: ws[0, 0] = value -> A1
            row, col = key
            self.cell(row + 1, col + 1).value = value
        elif isinstance(key, str):
            # Excel coordinate: ws['A1'] = value
            cell = self[key]
            cell.value = value
        else:
            raise InvalidCoordinateError(f"Invalid cell assignment pattern: {key}")
    
    def cell(self, row: int, column: int, value: CellValue = None) -> Cell:
        """Get or create cell at specified position (1-based)."""
        if row < 1 or column < 1:
            raise InvalidCoordinateError(f"Row and column must be >= 1, got ({row}, {column})")
        
        coord = (row, column)
        
        if coord not in self._cells:
            self._cells[coord] = Cell(self, row, column)
            self._update_bounds(row, column)
        
        if value is not None:
            self._cells[coord].value = value
        
        return self._cells[coord]
    
    def append(self, iterable: List[CellValue]):
        """Add row of data to end of worksheet (like list.append)."""
        if not iterable:
            return
        
        row = self._max_row + 1
        for col, value in enumerate(iterable, 1):
            if value is not None:  # Skip None values to save memory
                self.cell(row, col, value)
    
    def extend(self, data: List[List[CellValue]]):
        """Add multiple rows of data (like list.extend)."""
        for row_data in data:
            self.append(row_data)
    
    def insert(self, index: int, iterable: List[CellValue]):
        """Insert row at specified position (like list.insert)."""
        if index < 1:
            index = 1
        
        # Shift existing data down
        cells_to_move = []
        for coord, cell in self._cells.items():
            row, col = coord
            if row >= index:
                cells_to_move.append((coord, cell))
        
        # Remove old cells and create new ones
        for old_coord, cell in cells_to_move:
            old_row, col = old_coord
            del self._cells[old_coord]
            new_cell = Cell(self, old_row + 1, col, cell.value)
            if cell._style:
                new_cell._style = cell._style.copy()
            new_cell._number_format = cell._number_format
            self._cells[(old_row + 1, col)] = new_cell
        
        # Insert new row
        for col, value in enumerate(iterable, 1):
            if value is not None:
                self.cell(index, col, value)
    
    def from_records(self, records: List[Dict[str, CellValue]], include_headers: bool = True):
        """Import data from list of dictionaries (like pandas.from_records)."""
        if not records:
            return
        
        # Get column names
        keys = list(records[0].keys())
        
        start_row = self._max_row + 1
        
        # Add headers if requested
        if include_headers:
            self.append(keys)
        
        # Add data rows
        for record in records:
            row_data = [record.get(key) for key in keys]
            self.append(row_data)
    
    def rows(self) -> Iterator[List[Cell]]:
        """Iterate over all rows with data."""
        for row_idx in range(1, self.max_row + 1):
            row_cells = []
            for col_idx in range(1, self.max_column + 1):
                coord = (row_idx, col_idx)
                if coord in self._cells:
                    row_cells.append(self._cells[coord])
                else:
                    # Create empty cell for iteration
                    row_cells.append(Cell(self, row_idx, col_idx))
            yield row_cells
    
    def columns(self) -> Iterator[List[Cell]]:
        """Iterate over all columns with data."""
        for col_idx in range(1, self.max_column + 1):
            col_cells = []
            for row_idx in range(1, self.max_row + 1):
                coord = (row_idx, col_idx)
                if coord in self._cells:
                    col_cells.append(self._cells[coord])
                else:
                    col_cells.append(Cell(self, row_idx, col_idx))
            yield col_cells
    
    def iter_rows(self, min_row: int = 1, max_row: Optional[int] = None,
                  min_col: int = 1, max_col: Optional[int] = None) -> Iterator[List[Cell]]:
        """Iterate over specified range of rows."""
        if max_row is None:
            max_row = self.max_row or 1
        if max_col is None:
            max_col = self.max_column or 1
        
        for row_idx in range(min_row, max_row + 1):
            row_cells = []
            for col_idx in range(min_col, max_col + 1):
                coord = (row_idx, col_idx)
                if coord in self._cells:
                    row_cells.append(self._cells[coord])
                else:
                    row_cells.append(Cell(self, row_idx, col_idx))
            yield row_cells
    
    def iter_cols(self, min_row: int = 1, max_row: Optional[int] = None,
                  min_col: int = 1, max_col: Optional[int] = None) -> Iterator[List[Cell]]:
        """Iterate over specified range of columns."""
        if max_row is None:
            max_row = self.max_row or 1
        if max_col is None:
            max_col = self.max_column or 1
        
        for col_idx in range(min_col, max_col + 1):
            col_cells = []
            for row_idx in range(min_row, max_row + 1):
                coord = (row_idx, col_idx)
                if coord in self._cells:
                    col_cells.append(self._cells[coord])
                else:
                    col_cells.append(Cell(self, row_idx, col_idx))
            yield col_cells
    
    def merge_cells(self, range_string: str):
        """Merge cells in specified range."""
        self._merged_ranges.add(range_string.upper())
    
    def unmerge_cells(self, range_string: str):
        """Unmerge previously merged cells."""
        self._merged_ranges.discard(range_string.upper())
    
    def delete_rows(self, idx: int, amount: int = 1):
        """Delete specified number of rows."""
        for _ in range(amount):
            # Remove cells in row
            cells_to_remove = [(row, col) for row, col in self._cells.keys() if row == idx]
            for coord in cells_to_remove:
                del self._cells[coord]
            
            # Shift rows up
            cells_to_move = [(row, col) for row, col in self._cells.keys() if row > idx]
            for old_row, col in cells_to_move:
                cell = self._cells.pop((old_row, col))
                cell._row = old_row - 1
                self._cells[(old_row - 1, col)] = cell
            
            # Update max_row
            if self._cells:
                self._max_row = max(row for row, col in self._cells.keys())
            else:
                self._max_row = 0
    
    def delete_cols(self, idx: int, amount: int = 1):
        """Delete specified number of columns."""
        for _ in range(amount):
            # Remove cells in column
            cells_to_remove = [(row, col) for row, col in self._cells.keys() if col == idx]
            for coord in cells_to_remove:
                del self._cells[coord]
            
            # Shift columns left
            cells_to_move = [(row, col) for row, col in self._cells.keys() if col > idx]
            for row, old_col in cells_to_move:
                cell = self._cells.pop((row, old_col))
                cell._column = old_col - 1
                self._cells[(row, old_col - 1)] = cell
            
            # Update max_column
            if self._cells:
                self._max_column = max(col for row, col in self._cells.keys())
            else:
                self._max_column = 0
    
    def freeze_panes(self, cell: Union[str, Cell, None] = None):
        """Freeze panes at specified cell."""
        if cell is None:
            self._freeze_panes = None
        elif isinstance(cell, str):
            self._freeze_panes = cell.upper()
        elif isinstance(cell, Cell):
            self._freeze_panes = cell.coordinate
        else:
            raise InvalidCoordinateError(f"Invalid freeze panes cell: {cell}")
    
    def __str__(self) -> str:
        """String representation."""
        return f"Worksheet('{self._name}')"
    
    def __repr__(self) -> str:
        """Debug representation."""
        return f"Worksheet(name='{self._name}', max_row={self._max_row}, max_col={self._max_column})"
    
    # Column width and row height functionality
    def set_column_width(self, column: Union[int, str], width: float):
        """Set width for a specific column (0-based int or Excel letter)."""
        if isinstance(column, str):
            # Convert letter to 0-based number (A=0, B=1, etc.)
            col_num = ord(column.upper()) - ord('A')
        else:
            # 0-based integer column index
            col_num = column
        
        if col_num < 0:
            raise InvalidCoordinateError(f"Column must be >= 0, got {col_num}")
        
        # Store internally as 1-based for compatibility
        self._column_widths[col_num + 1] = width
    
    def get_column_width(self, column: Union[int, str]) -> float:
        """Get width for a specific column (0-based int or Excel letter)."""
        if isinstance(column, str):
            # Convert letter to 0-based number (A=0, B=1, etc.)
            col_num = ord(column.upper()) - ord('A')
        else:
            # 0-based integer column index
            col_num = column
        
        # Retrieve using 1-based internal storage
        return self._column_widths.get(col_num + 1, 10.0)  # Default width
    
    def set_row_height(self, row: int, height: float):
        """Set height for a specific row (0-based)."""
        if row < 0:
            raise InvalidCoordinateError(f"Row must be >= 0, got {row}")
        
        # Store internally as 1-based for compatibility
        self._row_heights[row + 1] = height
    
    def get_row_height(self, row: int) -> float:
        """Get height for a specific row (0-based)."""
        # Retrieve using 1-based internal storage
        return self._row_heights.get(row + 1, 15.0)  # Default height
    
    def auto_size_column(self, column: Union[int, str]):
        """Auto-size column based on content (0-based int or Excel letter)."""
        if isinstance(column, str):
            # Convert letter to 0-based number
            col_num = ord(column.upper()) - ord('A')
        else:
            # 0-based integer column index
            col_num = column
        
        max_width = 10.0  # Default minimum
        
        # Calculate based on cell content (internal storage is 1-based)
        internal_col = col_num + 1
        for (row, col), cell in self._cells.items():
            if col == internal_col and cell.value is not None:
                content_length = len(str(cell.value))
                estimated_width = min(content_length * 1.2, 50.0)  # Max width 50
                max_width = max(max_width, estimated_width)
        
        self._column_widths[internal_col] = max_width
    
    def set_cell_style(self, coordinate: Union[str, Tuple[int, int]], **style_kwargs):
        """Set cell style with convenient keyword arguments."""
        cell = self[coordinate]
        
        # Font properties
        if 'font_name' in style_kwargs:
            cell.font.name = style_kwargs['font_name']
        if 'font_size' in style_kwargs:
            cell.font.size = style_kwargs['font_size']
        if 'bold' in style_kwargs:
            cell.font.bold = style_kwargs['bold']
        if 'italic' in style_kwargs:
            cell.font.italic = style_kwargs['italic']
        if 'font_color' in style_kwargs:
            cell.font.color = style_kwargs['font_color']
        
        # Fill properties
        if 'fill_color' in style_kwargs:
            cell.fill.color = style_kwargs['fill_color']
        
        # Number format
        if 'number_format' in style_kwargs:
            cell.number_format = style_kwargs['number_format']
        
        # Alignment
        if 'horizontal' in style_kwargs:
            cell.alignment.horizontal = style_kwargs['horizontal']
        if 'vertical' in style_kwargs:
            cell.alignment.vertical = style_kwargs['vertical']
    
    def set_range_style(self, range_str: str, **style_kwargs):
        """Set style for entire range with convenient keyword arguments."""
        range_obj = self[range_str]
        for cell in range_obj:
            self.set_cell_style(cell.coordinate, **style_kwargs)
    
    def populate_data(self, start_cell: Union[str, Tuple[int, int]], data: List[List], 
                     column_styles: Dict[int, Dict] = None, conditional_styles: Dict = None):
        """Populate data with automatic styling based on column and conditions.
        
        Args:
            start_cell: Starting cell coordinate
            data: 2D list of data to populate
            column_styles: Dict mapping column index (0-based) to style kwargs
            conditional_styles: Dict with condition functions and styles
        """
        if isinstance(start_cell, str):
            from .utils import coordinate_to_tuple
            start_row, start_col = coordinate_to_tuple(start_cell)
        else:
            start_row, start_col = start_cell[0] + 1, start_cell[1] + 1  # Convert to 1-based
        
        for row_offset, row_data in enumerate(data):
            for col_offset, value in enumerate(row_data):
                current_row = start_row + row_offset
                current_col = start_col + col_offset
                
                # Set the value
                cell = self.cell(current_row, current_col, value)
                
                # Apply column-specific styles
                if column_styles and col_offset in column_styles:
                    coord = (current_row - 1, current_col - 1)  # Convert to 0-based for style method
                    self.set_cell_style(coord, **column_styles[col_offset])
                
                # Apply conditional styles
                if conditional_styles:
                    for condition_name, condition_config in conditional_styles.items():
                        condition_func = condition_config['condition']
                        if condition_func(value, row_offset, col_offset):
                            coord = (current_row - 1, current_col - 1)
                            style_dict = condition_config['style']
                            # Handle both static dict and function that returns dict
                            if callable(style_dict):
                                style_dict = style_dict(value)
                            self.set_cell_style(coord, **style_dict)
    
    def apply_column_formats(self, start_col: int, formats: List[str]):
        """Apply number formats to consecutive columns.
        
        Args:
            start_col: Starting column index (0-based)
            formats: List of number format strings
        """
        for i, fmt in enumerate(formats):
            col_idx = start_col + i
            # Only iterate over rows that have data in this column
            for coord, cell in self._cells.items():
                row_idx, cell_col_idx = coord
                if cell_col_idx == col_idx:
                    try:
                        self.set_cell_style(coord, number_format=fmt)
                    except (KeyError, ValueError, TypeError):
                        # Skip cells that have invalid coordinates
                        continue
    
    def create_table(self, start_cell: Union[str, Tuple[int, int]], 
                    headers: List[str], data: List[List],
                    header_style: Dict = None, 
                    column_styles: Dict[int, Dict] = None,
                    conditional_styles: Dict = None,
                    auto_width: bool = True):
        """Create a complete table with headers, data, and styling in one call.
        
        Args:
            start_cell: Starting cell coordinate  
            headers: List of header names
            data: 2D list of table data
            header_style: Style dict for header row
            column_styles: Dict mapping column index to style kwargs
            conditional_styles: Dict with conditional formatting rules
            auto_width: Whether to auto-size columns
        """
        if isinstance(start_cell, str):
            from .utils import coordinate_to_tuple
            start_row, start_col = coordinate_to_tuple(start_cell)
        else:
            start_row, start_col = start_cell[0] + 1, start_cell[1] + 1
        
        # Add headers
        for col_offset, header in enumerate(headers):
            cell = self.cell(start_row, start_col + col_offset, header)
            if header_style:
                coord = (start_row - 1, start_col + col_offset - 1)
                self.set_cell_style(coord, **header_style)
        
        # Add data with styles
        if data:
            data_start = (start_row, start_col - 1)  # Adjust for populate_data
            self.populate_data(data_start, data, column_styles, conditional_styles)
        
        # Auto-size columns if requested
        if auto_width:
            for i in range(len(headers)):
                try:
                    self.auto_size_column(start_col - 1 + i)  # Convert to 0-based
                except (IndexError, ValueError):
                    # Skip invalid column indices
                    continue
    
