"""
Range implementation for operating on multiple cells as a group.
"""

from typing import List, Iterator, Tuple, Union, TYPE_CHECKING
from .formats import CellValue
from .utils import parse_range, InvalidCoordinateError
from .style import Style, Font, Fill

if TYPE_CHECKING:
    from .worksheet import Worksheet
    from .cell import Cell


class Range:
    """Excel range representing a rectangular area of cells."""
    
    def __init__(self, worksheet: 'Worksheet', range_string: str):
        self._worksheet = worksheet
        self._range_string = range_string.upper()
        
        try:
            (self._start_row, self._start_col), (self._end_row, self._end_col) = parse_range(range_string)
        except Exception as e:
            raise InvalidCoordinateError(f"Invalid range format: {range_string}") from e
        
        # Ensure start <= end
        if self._start_row > self._end_row:
            self._start_row, self._end_row = self._end_row, self._start_row
        if self._start_col > self._end_col:
            self._start_col, self._end_col = self._end_col, self._start_col
    
    @property
    def coordinate(self) -> str:
        """Range coordinate string."""
        return self._range_string
    
    @property
    def min_row(self) -> int:
        """Minimum row in range."""
        return self._start_row
    
    @property
    def max_row(self) -> int:
        """Maximum row in range."""
        return self._end_row
    
    @property
    def min_column(self) -> int:
        """Minimum column in range."""
        return self._start_col
    
    @property
    def max_column(self) -> int:
        """Maximum column in range."""
        return self._end_col
    
    @property
    def row_count(self) -> int:
        """Number of rows in range."""
        return self._end_row - self._start_row + 1
    
    @property
    def column_count(self) -> int:
        """Number of columns in range."""
        return self._end_col - self._start_col + 1
    
    @property
    def size(self) -> Tuple[int, int]:
        """Range size as (rows, columns)."""
        return self.row_count, self.column_count
    
    def cells(self) -> Iterator['Cell']:
        """Iterate over all cells in range."""
        for row in range(self._start_row, self._end_row + 1):
            for col in range(self._start_col, self._end_col + 1):
                yield self._worksheet.cell(row, col)
    
    def rows_iter(self) -> Iterator[List['Cell']]:
        """Iterate over rows of cells."""
        for row in range(self._start_row, self._end_row + 1):
            row_cells = []
            for col in range(self._start_col, self._end_col + 1):
                row_cells.append(self._worksheet.cell(row, col))
            yield row_cells
    
    def columns_iter(self) -> Iterator[List['Cell']]:
        """Iterate over columns of cells."""
        for col in range(self._start_col, self._end_col + 1):
            col_cells = []
            for row in range(self._start_row, self._end_row + 1):
                col_cells.append(self._worksheet.cell(row, col))
            yield col_cells
    
    def rows(self) -> Iterator[List['Cell']]:
        """Iterate over rows of cells (alias for rows_iter for test compatibility)."""
        return self.rows_iter()
    
    def columns(self) -> Iterator[List['Cell']]:
        """Iterate over columns of cells (alias for columns_iter for test compatibility)."""
        return self.columns_iter()
    
    @property
    def values(self) -> List[List[CellValue]]:
        """Get all values as nested list."""
        result = []
        for row in range(self._start_row, self._end_row + 1):
            row_values = []
            for col in range(self._start_col, self._end_col + 1):
                cell = self._worksheet.cell(row, col)
                row_values.append(cell.value)
            result.append(row_values)
        return result
    
    @values.setter
    def values(self, data: Union[List[List[CellValue]], List[CellValue], CellValue]):
        """Set values from nested list or single value."""
        if not isinstance(data, list):
            # Single value - apply to all cells
            for cell in self.cells():
                cell.value = data
            return
        
        if not data:
            return
        
        # Check if it's a list of lists (2D) or single list (1D)
        if isinstance(data[0], list):
            # 2D data
            for row_idx, row_data in enumerate(data):
                if row_idx >= self.row_count:
                    break
                for col_idx, value in enumerate(row_data):
                    if col_idx >= self.column_count:
                        break
                    cell = self._worksheet.cell(
                        self._start_row + row_idx,
                        self._start_col + col_idx
                    )
                    cell.value = value
        else:
            # 1D data - fill row by row
            flat_index = 0
            for row in range(self._start_row, self._end_row + 1):
                for col in range(self._start_col, self._end_col + 1):
                    if flat_index >= len(data):
                        return
                    cell = self._worksheet.cell(row, col)
                    cell.value = data[flat_index]
                    flat_index += 1
    
    @property
    def font(self) -> Font:
        """Get font of first cell (for styling entire range)."""
        first_cell = self._worksheet.cell(self._start_row, self._start_col)
        return first_cell.font
    
    @font.setter
    def font(self, value: Font):
        """Apply font to entire range."""
        for cell in self.cells():
            cell.style.font = value.copy()
    
    @property
    def fill(self) -> Fill:
        """Get fill of first cell (for styling entire range)."""
        first_cell = self._worksheet.cell(self._start_row, self._start_col)
        return first_cell.fill
    
    @fill.setter
    def fill(self, value: Fill):
        """Apply fill to entire range."""
        for cell in self.cells():
            cell.style.fill = value.copy()
    
    def apply_style(self, style: Style):
        """Apply complete style to entire range."""
        for cell in self.cells():
            cell.style = style.copy()
    
    def clear(self):
        """Clear all values and formatting in range."""
        for cell in self.cells():
            cell.clear()
    
    def merge(self):
        """Mark range for merging (implementation depends on writer)."""
        if hasattr(self._worksheet, '_merged_ranges'):
            self._worksheet._merged_ranges.add(self._range_string)
    
    def unmerge(self):
        """Unmerge previously merged range."""
        if hasattr(self._worksheet, '_merged_ranges'):
            self._worksheet._merged_ranges.discard(self._range_string)
    
    def __str__(self) -> str:
        """String representation."""
        return f"Range({self._range_string})"
    
    def __repr__(self) -> str:
        """Debug representation."""
        return f"Range('{self._range_string}', {self.row_count}x{self.column_count})"
    
    def __iter__(self) -> Iterator['Cell']:
        """Iterate over cells."""
        return self.cells()
    
    def __len__(self) -> int:
        """Number of cells in range."""
        return self.row_count * self.column_count