"""
Data validation utilities for Excel operations.
"""

from typing import Union
from datetime import datetime
import re

from ..formats import CellValue
from ..constants import (
    MAX_SHEET_NAME_LENGTH, INVALID_SHEET_NAME_CHARS,
    CELL_REF_PATTERN, DEFAULT_SHEET_NAME
)


def is_numeric_string(value: str) -> bool:
    """Check if string represents a numeric value."""
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False


def is_formula(value: str) -> bool:
    """Check if string represents an Excel formula."""
    return isinstance(value, str) and value.startswith('=')


def is_date_value(value: CellValue) -> bool:
    """Check if value represents a date."""
    return isinstance(value, datetime)


def infer_data_type(value: CellValue) -> str:
    """Infer Excel data type from Python value."""
    if value is None:
        return 'empty'
    elif isinstance(value, bool):
        return 'boolean'
    elif isinstance(value, (int, float)):
        return 'number'
    elif isinstance(value, datetime):
        return 'date'
    elif isinstance(value, str):
        if is_formula(value):
            return 'formula'
        elif is_numeric_string(value):
            return 'number'
        else:
            return 'string'
    else:
        return 'string'


def validate_sheet_name(name: str) -> bool:
    """Validate Excel worksheet name."""
    if not name or len(name) > MAX_SHEET_NAME_LENGTH:
        return False
    
    return not any(char in name for char in INVALID_SHEET_NAME_CHARS)


def sanitize_sheet_name(name: str) -> str:
    """Sanitize worksheet name for Excel compatibility."""
    if not name:
        return DEFAULT_SHEET_NAME + "1"
    
    # Remove invalid characters
    for char in INVALID_SHEET_NAME_CHARS:
        name = name.replace(char, '_')
    
    # Truncate to maximum length
    if len(name) > MAX_SHEET_NAME_LENGTH:
        name = name[:MAX_SHEET_NAME_LENGTH]
    
    return name or DEFAULT_SHEET_NAME + "1"


def validate_cell_reference(ref: str) -> bool:
    """Validate Excel cell reference format (e.g., A1, Z99, AA100)."""
    if not ref or not isinstance(ref, str):
        return False
    
    return bool(re.match(CELL_REF_PATTERN, ref.upper()))


def convert_value(value: CellValue, target_type: str, default: CellValue = None) -> CellValue:
    """Convert value to target type with fallback to default."""
    try:
        if target_type == 'string':
            return str(value) if value is not None else ""
        elif target_type == 'int':
            return int(float(str(value)))
        elif target_type == 'float':
            return float(value)
        elif target_type == 'bool':
            return bool(value)
        else:
            return value
    except (ValueError, TypeError):
        return default