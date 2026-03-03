"""
Aspose.Cells for Python - Formula Evaluator Module

This module provides basic formula evaluation functionality for cells
that have formulas but no cached values.

Supports:
- String and numeric literals
- Cell references (A1 style)
- Defined names
- Basic functions: CONCATENATE, TEXT, IF, AND, OR, NOT

Note: This is a simplified evaluator for common cases. Complex formulas
may not evaluate correctly and will return None.
"""

import re
from typing import Any, Optional, Dict, List


class FormulaEvaluator:
    """
    Basic formula evaluator for xlsx cells without cached values.

    Examples:
        >>> evaluator = FormulaEvaluator(workbook)
        >>> result = evaluator.evaluate('=CONCATENATE("Hello ", A1)', worksheet)
    """

    def __init__(self, workbook):
        """
        Initialize the formula evaluator.

        Args:
            workbook: The Workbook object containing defined names and worksheets.
        """
        self._workbook = workbook
        self._defined_names_cache = None

    def evaluate(self, formula: str, worksheet=None, cell_ref: str = None) -> Optional[Any]:
        """
        Evaluate a formula and return the result.

        Args:
            formula: The formula string (with or without leading '=').
            worksheet: The worksheet context for cell references.
            cell_ref: The cell reference being evaluated (to prevent circular refs).

        Returns:
            The evaluated value, or None if evaluation fails.
        """
        if formula is None:
            return None

        # Remove leading '=' if present
        if formula.startswith('='):
            formula = formula[1:]

        try:
            return self._evaluate_expression(formula, worksheet, cell_ref)
        except Exception:
            # Return None for any evaluation error
            return None

    def _get_defined_names(self) -> Dict[str, str]:
        """Get a dictionary of defined names to their refers_to values."""
        if self._defined_names_cache is None:
            self._defined_names_cache = {}
            if hasattr(self._workbook, 'properties') and hasattr(self._workbook.properties, 'defined_names'):
                for dn in self._workbook.properties.defined_names._names:
                    self._defined_names_cache[dn.name] = dn.refers_to
        return self._defined_names_cache

    def _evaluate_expression(self, expr: str, worksheet=None, cell_ref: str = None) -> Optional[Any]:
        """
        Evaluate an expression (formula without leading '=').
        """
        expr = expr.strip()

        # Check for string literal
        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1]

        # Check for numeric literal
        try:
            if '.' in expr:
                return float(expr)
            return int(expr)
        except ValueError:
            pass

        # Check for boolean literal
        if expr.upper() == 'TRUE':
            return True
        if expr.upper() == 'FALSE':
            return False

        # Check for function call
        func_match = re.match(r'^([A-Z_][A-Z0-9_]*)\s*\((.*)\)$', expr, re.IGNORECASE | re.DOTALL)
        if func_match:
            func_name = func_match.group(1).upper()
            args_str = func_match.group(2)
            return self._evaluate_function(func_name, args_str, worksheet, cell_ref)

        # Check for defined name
        defined_names = self._get_defined_names()
        if expr in defined_names:
            refers_to = defined_names[expr]
            return self._resolve_reference(refers_to, worksheet)

        # Check for cell reference (e.g., A1, Sheet1!A1)
        if re.match(r'^[A-Za-z]+[0-9]+$', expr) or '!' in expr:
            return self._resolve_reference(expr, worksheet)

        return None

    def _evaluate_function(self, func_name: str, args_str: str, worksheet=None, cell_ref: str = None) -> Optional[Any]:
        """
        Evaluate a function with its arguments.
        """
        args = self._parse_function_args(args_str)

        if func_name == 'CONCATENATE' or func_name == 'CONCAT':
            return self._func_concatenate(args, worksheet, cell_ref)

        if func_name == 'TEXT':
            return self._func_text(args, worksheet, cell_ref)

        if func_name == 'IF':
            return self._func_if(args, worksheet, cell_ref)

        if func_name == 'AND':
            return self._func_and(args, worksheet, cell_ref)

        if func_name == 'OR':
            return self._func_or(args, worksheet, cell_ref)

        if func_name == 'NOT':
            return self._func_not(args, worksheet, cell_ref)

        if func_name == 'LEN':
            return self._func_len(args, worksheet, cell_ref)

        if func_name == 'TRIM':
            return self._func_trim(args, worksheet, cell_ref)

        if func_name == 'UPPER':
            return self._func_upper(args, worksheet, cell_ref)

        if func_name == 'LOWER':
            return self._func_lower(args, worksheet, cell_ref)

        # Unknown function
        return None

    def _parse_function_args(self, args_str: str) -> List[str]:
        """
        Parse function arguments, handling nested parentheses and quoted strings.
        """
        args = []
        current = []
        depth = 0
        in_string = False

        for i, ch in enumerate(args_str):
            if ch == '"' and (i == 0 or args_str[i-1] != '\\'):
                in_string = not in_string
                current.append(ch)
            elif ch == '(' and not in_string:
                depth += 1
                current.append(ch)
            elif ch == ')' and not in_string:
                depth -= 1
                current.append(ch)
            elif ch == ',' and depth == 0 and not in_string:
                args.append(''.join(current).strip())
                current = []
            else:
                current.append(ch)

        if current:
            args.append(''.join(current).strip())

        return args

    def _resolve_reference(self, reference: str, worksheet=None) -> Optional[Any]:
        """
        Resolve a cell reference or range and return its value.
        """
        # Handle sheet-qualified reference (e.g., Sheet1!A1)
        if '!' in reference:
            sheet_part, cell_part = reference.split('!', 1)
            # Remove quotes if present (e.g., 'Sheet Name'!A1)
            sheet_part = sheet_part.strip("'")
            # Find the sheet by name
            target_sheet = None
            for ws in self._workbook.worksheets:
                if ws.name == sheet_part:
                    target_sheet = ws
                    break
            if target_sheet is None:
                return None
            worksheet = target_sheet
            reference = cell_part

        # Remove $ for absolute references
        reference = reference.replace('$', '')

        # Handle single cell reference
        if re.match(r'^[A-Za-z]+[0-9]+$', reference):
            if worksheet is None:
                return None
            try:
                cell = worksheet.cells[reference]
            except (KeyError, IndexError):
                return None
            if cell is None:
                return None
            # Return value, or evaluate formula if no cached value
            if cell.value is not None:
                return cell.value
            if cell.formula is not None:
                # Recursive evaluation (with depth limit built into exception handling)
                return self.evaluate(cell.formula, worksheet, reference)
            return None

        return None

    # Function implementations

    def _func_concatenate(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[str]:
        """CONCATENATE function - joins text strings."""
        result = []
        for arg in args:
            value = self._evaluate_expression(arg, worksheet, cell_ref)
            if value is None:
                value = ''
            result.append(str(value))
        return ''.join(result)

    def _func_text(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[str]:
        """TEXT function - formats a number as text."""
        if len(args) < 2:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        format_str = self._evaluate_expression(args[1], worksheet, cell_ref)
        if value is None or format_str is None:
            return None
        # Basic format support (just return string representation for now)
        return str(value)

    def _func_if(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[Any]:
        """IF function - conditional."""
        if len(args) < 2:
            return None
        condition = self._evaluate_expression(args[0], worksheet, cell_ref)
        if condition:
            return self._evaluate_expression(args[1], worksheet, cell_ref)
        elif len(args) > 2:
            return self._evaluate_expression(args[2], worksheet, cell_ref)
        return False

    def _func_and(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[bool]:
        """AND function - logical AND."""
        for arg in args:
            value = self._evaluate_expression(arg, worksheet, cell_ref)
            if not value:
                return False
        return True

    def _func_or(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[bool]:
        """OR function - logical OR."""
        for arg in args:
            value = self._evaluate_expression(arg, worksheet, cell_ref)
            if value:
                return True
        return False

    def _func_not(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[bool]:
        """NOT function - logical NOT."""
        if len(args) < 1:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        return not value

    def _func_len(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[int]:
        """LEN function - string length."""
        if len(args) < 1:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        if value is None:
            return 0
        return len(str(value))

    def _func_trim(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[str]:
        """TRIM function - removes extra spaces."""
        if len(args) < 1:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        if value is None:
            return ''
        return ' '.join(str(value).split())

    def _func_upper(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[str]:
        """UPPER function - converts to uppercase."""
        if len(args) < 1:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        if value is None:
            return ''
        return str(value).upper()

    def _func_lower(self, args: List[str], worksheet=None, cell_ref: str = None) -> Optional[str]:
        """LOWER function - converts to lowercase."""
        if len(args) < 1:
            return None
        value = self._evaluate_expression(args[0], worksheet, cell_ref)
        if value is None:
            return ''
        return str(value).lower()
