"""
Cell implementation with value management and styling capabilities.
"""

from typing import Optional, TYPE_CHECKING
from datetime import datetime

from .formats import CellValue
from .style import Style, Font, Fill, Border, Alignment
from .utils import (
    infer_data_type, 
    convert_value, 
    tuple_to_coordinate,
    CellValueError
)

if TYPE_CHECKING:
    from .worksheet import Worksheet


class Cell:
    """Individual Excel cell with value, type, and styling management."""
    
    def __init__(self, worksheet: 'Worksheet', row: int, column: int, value: CellValue = None):
        # Validate input parameters
        if not isinstance(row, int) or row < 1:
            raise ValueError(f"Row must be a positive integer, got: {row}")
        if not isinstance(column, int) or column < 1:
            raise ValueError(f"Column must be a positive integer, got: {column}")
        
        self._worksheet = worksheet
        self._row = row
        self._column = column
        self._value = value
        self._data_type: Optional[str] = None
        self._style: Optional[Style] = None
        self._number_format: str = "General"
        self._hyperlink: Optional[str] = None
        self._comment: Optional[str] = None
        self._formula: Optional[str] = None  # Store original formula
        self._calculated_value: Optional[CellValue] = None  # Store calculated result
        
        if value is not None:
            self.value = value
    
    @property
    def row(self) -> int:
        """Row number (1-based)."""
        return self._row
    
    @property
    def column(self) -> int:
        """Column number (1-based)."""
        return self._column
    
    @property
    def coordinate(self) -> str:
        """Excel coordinate (e.g., 'A1')."""
        return tuple_to_coordinate(self._row, self._column)
    
    @property
    def worksheet(self) -> 'Worksheet':
        """Parent worksheet."""
        return self._worksheet
    
    @property
    def value(self) -> CellValue:
        """Cell value."""
        return self._value
    
    @value.setter
    def value(self, val: CellValue):
        """Set cell value with automatic type inference."""
        self._value = val
        self._data_type = infer_data_type(val)
        
        # Update worksheet bounds
        if hasattr(self._worksheet, '_update_bounds'):
            self._worksheet._update_bounds(self._row, self._column)
    
    @property
    def data_type(self) -> Optional[str]:
        """Inferred data type."""
        return self._data_type
    
    @property
    def number_format(self) -> str:
        """Number format string."""
        return self._number_format
    
    @number_format.setter
    def number_format(self, value: str):
        """Set number format."""
        self._number_format = value
    
    @property
    def font(self) -> Font:
        """Font styling (creates style if needed)."""
        if self._style is None:
            self._style = Style()
        return self._style.font
    
    @property
    def fill(self) -> Fill:
        """Fill styling (creates style if needed)."""
        if self._style is None:
            self._style = Style()
        return self._style.fill
    
    @property
    def border(self) -> Border:
        """Border styling (creates style if needed)."""
        if self._style is None:
            self._style = Style()
        return self._style.border
    
    @property
    def alignment(self) -> Alignment:
        """Alignment styling (creates style if needed)."""
        if self._style is None:
            self._style = Style()
        return self._style.alignment
    
    @property
    def style(self) -> Style:
        """Complete style object (creates if needed)."""
        if self._style is None:
            self._style = Style()
        return self._style
    
    @style.setter
    def style(self, value: Style):
        """Set complete style."""
        self._style = value
    
    @property
    def hyperlink(self) -> Optional[str]:
        """Hyperlink URL."""
        return self._hyperlink
    
    @hyperlink.setter
    def hyperlink(self, value: Optional[str]):
        """Set hyperlink URL."""
        self._hyperlink = value
    
    @property
    def comment(self) -> Optional[str]:
        """Cell comment text."""
        return self._comment
    
    @comment.setter
    def comment(self, value: Optional[str]):
        """Set cell comment."""
        self._comment = value
    
    def as_str(self, default: str = "") -> str:
        """Convert value to string."""
        return convert_value(self._value, 'string', default)
    
    def as_int(self, default: int = 0) -> int:
        """Convert value to integer."""
        return convert_value(self._value, 'int', default)
    
    def as_float(self, default: float = 0.0) -> float:
        """Convert value to float."""
        return convert_value(self._value, 'float', default)
    
    def as_bool(self, default: bool = False) -> bool:
        """Convert value to boolean."""
        return convert_value(self._value, 'bool', default)
    
    def is_numeric(self) -> bool:
        """Check if cell contains numeric value."""
        return self._data_type == 'number'
    
    def is_date(self) -> bool:
        """Check if cell contains date value."""
        return self._data_type == 'date' or isinstance(self._value, datetime)
    
    def is_formula(self) -> bool:
        """Check if cell contains formula."""
        return self._data_type == 'formula'
    
    def is_empty(self) -> bool:
        """Check if cell is empty."""
        return self._value is None or self._data_type == 'empty'
    
    def clear(self):
        """Clear cell value and formatting."""
        self._value = None
        self._data_type = 'empty'
        self._style = None
        self._number_format = "General"
        self._hyperlink = None
        self._comment = None
        self._formula = None
        self._calculated_value = None
    
    def set_formula(self, formula: str, calculated_value: CellValue = None):
        """Set cell formula (handles = prefix automatically)."""
        if not formula.startswith('='):
            formula = '=' + formula
        self._formula = formula
        self.value = formula
        self._data_type = 'formula'
        
        # Always ensure calculated value is set
        if calculated_value is not None:
            self._calculated_value = calculated_value
        else:
            # Always recalculate when setting a new formula
            self._calculated_value = self._get_basic_formula_result(formula)
    
    @property
    def formula(self) -> Optional[str]:
        """Get the original formula if this is a formula cell."""
        return self._formula
    
    @property
    def calculated_value(self) -> CellValue:
        """Get the calculated result of a formula, or the cell value if not a formula."""
        if self.is_formula() and self._calculated_value is not None:
            return self._calculated_value
        return self._value
    
    @property
    def display_value(self) -> str:
        """Get the display value (what should be shown to users)."""
        if self.is_formula():
            # For formulas, prefer calculated value over raw formula
            if self._calculated_value is not None:
                return str(self._calculated_value)
            else:
                return str(self._value)  # Fallback to formula text
        return str(self._value) if self._value is not None else ""
    
    def get_value(self, mode: str = 'display') -> CellValue:
        """
        Get cell value in different modes.
        
        Args:
            mode: 'display' (calculated/display value), 'formula' (raw formula), 'raw' (raw value)
        
        Returns:
            Cell value based on the requested mode
        """
        if mode == 'formula' and self.is_formula():
            return self._formula or self._value
        elif mode == 'display':
            return self.calculated_value
        elif mode == 'raw':
            return self._value
        else:
            return self.calculated_value  # Default to display mode
    
    def has_hyperlink(self) -> bool:
        """Check if cell has a hyperlink."""
        return self._hyperlink is not None and self._hyperlink.strip() != ""
    
    def get_markdown_link(self, text: Optional[str] = None) -> str:
        """Get markdown formatted link if cell has hyperlink."""
        if not self.has_hyperlink():
            return text or self.display_value
        
        link_text = text or self.display_value
        if not link_text:
            link_text = self._hyperlink
        
        return f"[{link_text}]({self._hyperlink})"
    
    def set_hyperlink(self, url: str, display_text: Optional[str] = None):
        """Set hyperlink with optional display text."""
        self._hyperlink = url
        if display_text is not None:
            self.value = display_text
    
    def copy_from(self, other: 'Cell'):
        """Copy value and style from another cell."""
        self.value = other.value
        if other._style:
            self._style = other._style.copy()
        self._number_format = other._number_format
        self._hyperlink = other._hyperlink
        self._comment = other._comment
        self._formula = other._formula
        self._calculated_value = other._calculated_value
    
    def __str__(self) -> str:
        """String representation."""
        return f"Cell({self.coordinate}={self._value})"
    
    def _get_basic_formula_result(self, formula: str) -> CellValue:
        """Get calculated result using the formula engine."""
        try:
            from .formula import FormulaEvaluator
            evaluator = FormulaEvaluator(self._worksheet)
            return evaluator.evaluate(formula, self.coordinate)
        except Exception:
            # Fallback to simple evaluation
            return self._simple_formula_fallback(formula)
    
    def _simple_formula_fallback(self, formula: str) -> CellValue:
        """Simple fallback for basic formulas when engine fails."""
        formula_upper = formula.upper().strip()
        
        # Remove = prefix
        if formula_upper.startswith('='):
            formula_upper = formula_upper[1:]
        
        # Handle very basic cases
        if formula_upper.startswith(('SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN')):
            return 0
        elif formula_upper.startswith(('NOW', 'TODAY')):
            return "2024-01-01"
        elif formula_upper.startswith('TRUE'):
            return True
        elif formula_upper.startswith('FALSE'):
            return False
        elif all(c in '0123456789+-*/.() ' for c in formula_upper):
            try:
                # Use safe expression evaluation instead of eval
                import ast
                node = ast.parse(formula_upper, mode='eval')
                if self._is_safe_expression(node):
                    return eval(compile(node, '<string>', 'eval'))
                else:
                    return 0
            except (ValueError, SyntaxError, TypeError):
                return 0
        else:
            return 0
    
    def _is_safe_expression(self, node) -> bool:
        """Check if AST node contains only safe mathematical operations."""
        import ast
        
        allowed_nodes = (
            ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Num,
            ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod, ast.Pow,
            ast.USub, ast.UAdd
        )
        
        for child in ast.walk(node):
            if not isinstance(child, allowed_nodes):
                return False
        return True
    
    def __repr__(self) -> str:
        """Debug representation."""
        return f"Cell({self.coordinate}, row={self._row}, col={self._column}, value={self._value!r}, type={self._data_type})"