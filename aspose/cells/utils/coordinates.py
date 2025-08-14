"""
Coordinate conversion utilities for Excel cell addressing.
"""

import re
from typing import Tuple
from .exceptions import InvalidCoordinateError


def column_index_to_letter(index: int) -> str:
    """Convert 1-based column index to Excel letter (1 -> A, 27 -> AA)."""
    if index < 1:
        raise InvalidCoordinateError(f"Column index must be >= 1, got {index}")
    
    result = ""
    while index > 0:
        index -= 1
        result = chr(65 + index % 26) + result
        index //= 26
    return result


def column_letter_to_index(letter: str) -> int:
    """Convert Excel column letter to 1-based index (A -> 1, AA -> 27)."""
    if not letter or not letter.isalpha():
        raise InvalidCoordinateError(f"Invalid column letter: {letter}")
    
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def coordinate_to_tuple(coordinate: str) -> Tuple[int, int]:
    """Convert Excel coordinate to (row, column) tuple (A1 -> (1, 1))."""
    match = re.match(r'^([A-Z]+)(\d+)$', coordinate.upper())
    if not match:
        raise InvalidCoordinateError(f"Invalid coordinate format: {coordinate}")
    
    col_letter, row_str = match.groups()
    row = int(row_str)
    col = column_letter_to_index(col_letter)
    
    if row < 1:
        raise InvalidCoordinateError(f"Row number must be >= 1, got {row}")
    
    return row, col


def tuple_to_coordinate(row: int, column: int) -> str:
    """Convert (row, column) tuple to Excel coordinate ((1, 1) -> A1)."""
    if row < 1 or column < 1:
        raise InvalidCoordinateError(f"Row and column must be >= 1, got ({row}, {column})")
    
    col_letter = column_index_to_letter(column)
    return f"{col_letter}{row}"


def parse_range(range_str: str) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    """Parse Excel range to ((start_row, start_col), (end_row, end_col))."""
    if ':' not in range_str:
        raise InvalidCoordinateError(f"Range must contain ':', got: {range_str}")
    
    start_cell, end_cell = range_str.split(':', 1)
    start_coord = coordinate_to_tuple(start_cell.strip())
    end_coord = coordinate_to_tuple(end_cell.strip())
    
    return start_coord, end_coord