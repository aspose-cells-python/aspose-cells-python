"""
Anchor system for positioning images within Excel worksheets.
"""

from typing import Tuple, Optional
from enum import Enum


class AnchorType(Enum):
    """Image anchoring behavior types."""
    
    ONE_CELL = "oneCell"      # Move but don't size with cells
    TWO_CELL = "twoCell"      # Move and size with cells  
    ABSOLUTE = "absolute"     # Fixed position, independent of cells


class Anchor:
    """Image positioning and anchoring information."""
    
    def __init__(self, anchor_type: AnchorType = AnchorType.TWO_CELL):
        self._type: AnchorType = anchor_type
        self._from_row: int = 0
        self._from_col: int = 0
        self._from_row_offset: int = 0
        self._from_col_offset: int = 0
        self._to_row: Optional[int] = None
        self._to_col: Optional[int] = None
        self._to_row_offset: int = 0
        self._to_col_offset: int = 0
        self._x: Optional[int] = None  # Absolute positioning
        self._y: Optional[int] = None  # Absolute positioning
    
    @property
    def type(self) -> AnchorType:
        """Get anchor type."""
        return self._type
    
    @type.setter 
    def type(self, value: AnchorType):
        """Set anchor type."""
        self._type = value
    
    @property
    def from_position(self) -> Tuple[int, int]:
        """Get starting position (row, col)."""
        return (self._from_row, self._from_col)
    
    @from_position.setter
    def from_position(self, value: Tuple[int, int]):
        """Set starting position (row, col)."""
        self._from_row, self._from_col = value
    
    @property
    def from_offset(self) -> Tuple[int, int]:
        """Get starting offset (row_offset, col_offset) in pixels."""
        return (self._from_row_offset, self._from_col_offset)
    
    @from_offset.setter
    def from_offset(self, value: Tuple[int, int]):
        """Set starting offset (row_offset, col_offset) in pixels."""
        self._from_row_offset, self._from_col_offset = value
    
    @property
    def to_position(self) -> Optional[Tuple[int, int]]:
        """Get ending position (row, col) for TWO_CELL anchor."""
        if self._type == AnchorType.TWO_CELL and self._to_row is not None:
            return (self._to_row, self._to_col)
        return None
    
    @to_position.setter
    def to_position(self, value: Optional[Tuple[int, int]]):
        """Set ending position (row, col) for TWO_CELL anchor."""
        if value is None:
            self._to_row = None
            self._to_col = None
        else:
            self._to_row, self._to_col = value
            if self._type == AnchorType.ONE_CELL:
                self._type = AnchorType.TWO_CELL
    
    @property
    def to_offset(self) -> Tuple[int, int]:
        """Get ending offset (row_offset, col_offset) in pixels."""
        return (self._to_row_offset, self._to_col_offset)
    
    @to_offset.setter
    def to_offset(self, value: Tuple[int, int]):
        """Set ending offset (row_offset, col_offset) in pixels."""
        self._to_row_offset, self._to_col_offset = value
    
    @property
    def absolute_position(self) -> Optional[Tuple[int, int]]:
        """Get absolute position (x, y) in pixels."""
        if self._type == AnchorType.ABSOLUTE and self._x is not None:
            return (self._x, self._y)
        return None
    
    @absolute_position.setter
    def absolute_position(self, value: Optional[Tuple[int, int]]):
        """Set absolute position (x, y) in pixels."""
        if value is None:
            self._x = None
            self._y = None
        else:
            self._x, self._y = value
            self._type = AnchorType.ABSOLUTE
    
    @classmethod
    def from_cell(cls, cell_ref: str, offset: Tuple[int, int] = (0, 0)) -> 'Anchor':
        """Create anchor from cell reference (e.g., 'A1', 'B2')."""
        from ..utils.coordinates import coordinate_to_tuple
        
        row, col = coordinate_to_tuple(cell_ref)
        anchor = cls(AnchorType.TWO_CELL)
        anchor.from_position = (row - 1, col - 1)  # Convert to 0-based
        anchor.from_offset = offset
        # Set default to_position for TWO_CELL anchor (Excel standard)
        anchor.to_position = (row + 4, col + 2)  # Default span
        anchor.to_offset = (0, 0)
        return anchor
    
    @classmethod
    def from_range(cls, start_cell: str, end_cell: str, 
                   start_offset: Tuple[int, int] = (0, 0),
                   end_offset: Tuple[int, int] = (0, 0)) -> 'Anchor':
        """Create TWO_CELL anchor from cell range."""
        from ..utils.coordinates import coordinate_to_tuple
        
        start_row, start_col = coordinate_to_tuple(start_cell)
        end_row, end_col = coordinate_to_tuple(end_cell)
        
        anchor = cls(AnchorType.TWO_CELL)
        anchor.from_position = (start_row - 1, start_col - 1)  # Convert to 0-based
        anchor.from_offset = start_offset
        anchor.to_position = (end_row - 1, end_col - 1)  # Convert to 0-based
        anchor.to_offset = end_offset
        return anchor
    
    @classmethod
    def absolute(cls, x: int, y: int) -> 'Anchor':
        """Create absolute positioned anchor."""
        anchor = cls(AnchorType.ABSOLUTE)
        anchor.absolute_position = (x, y)
        return anchor
    
    def copy(self) -> 'Anchor':
        """Create a copy of this anchor."""
        new_anchor = Anchor(self._type)
        new_anchor._from_row = self._from_row
        new_anchor._from_col = self._from_col
        new_anchor._from_row_offset = self._from_row_offset
        new_anchor._from_col_offset = self._from_col_offset
        new_anchor._to_row = self._to_row
        new_anchor._to_col = self._to_col
        new_anchor._to_row_offset = self._to_row_offset
        new_anchor._to_col_offset = self._to_col_offset
        new_anchor._x = self._x
        new_anchor._y = self._y
        return new_anchor
    
    def __str__(self) -> str:
        """String representation."""
        if self._type == AnchorType.ABSOLUTE:
            return f"Anchor(absolute: {self.absolute_position})"
        elif self._type == AnchorType.TWO_CELL:
            return f"Anchor(range: {self.from_position} -> {self.to_position})"
        else:
            return f"Anchor(cell: {self.from_position}, offset: {self.from_offset})"
    
    def __repr__(self) -> str:
        """Debug representation."""
        return f"Anchor(type={self._type.value}, from={self.from_position}, to={self.to_position})"