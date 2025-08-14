"""
Formula Evaluator - Evaluates Excel formulas using tokens and functions.
"""

import re
from typing import Any, Dict, List, Union, Optional, TYPE_CHECKING
from .tokenizer import Tokenizer, Token
from .functions import BUILTIN_FUNCTIONS, ExcelError, ValueErrorExcel, DivisionByZeroError

if TYPE_CHECKING:
    from ..worksheet import Worksheet


class CircularReferenceError(ExcelError):
    """Circular reference detected."""
    def __str__(self):
        return "#CIRCULAR!"


class FormulaEvaluator:
    """Evaluates Excel formulas."""
    
    def __init__(self, worksheet: Optional['Worksheet'] = None):
        self.worksheet = worksheet
        self._evaluation_stack = set()  # Track cells being evaluated to detect circular references
    
    def evaluate(self, formula: str, cell_address: Optional[str] = None) -> Any:
        """
        Evaluate a formula and return the result.
        
        Args:
            formula: The formula to evaluate (with or without = prefix)
            cell_address: Address of the cell containing this formula (for circular ref detection)
        
        Returns:
            The evaluated result
        """
        if not formula:
            return ""
        
        # Remove leading = if present
        if formula.startswith('='):
            formula = formula[1:]
        
        if not formula.strip():
            return ""
        
        # Check for circular references
        if cell_address and cell_address in self._evaluation_stack:
            raise CircularReferenceError()
        
        try:
            if cell_address:
                self._evaluation_stack.add(cell_address)
            
            # Tokenize the formula
            tokenizer = Tokenizer('=' + formula)
            tokens = list(tokenizer)
            
            if not tokens:
                return ""
            
            # Evaluate the token stream
            result = self._evaluate_tokens(tokens)
            return result
            
        except ExcelError:
            raise
        except Exception as e:
            raise ValueErrorExcel() from e
        finally:
            if cell_address:
                self._evaluation_stack.discard(cell_address)
    
    def _evaluate_tokens(self, tokens: List[Token]) -> Any:
        """Evaluate a list of tokens."""
        if not tokens:
            return ""
        
        # Simple expression evaluation using shunting yard algorithm
        output_queue = []
        operator_stack = []
        
        i = 0
        while i < len(tokens):
            token = tokens[i]
            
            if token.type == Token.OPERAND:
                output_queue.append(self._evaluate_operand(token))
            
            elif token.type == Token.FUNCTION:
                # Find matching closing parenthesis
                func_name = token.value
                if i + 1 < len(tokens) and tokens[i + 1].type == Token.SUBEXPR and tokens[i + 1].subtype == "OPEN":
                    args_start = i + 2
                    args_end = self._find_matching_paren(tokens, i + 1)
                    
                    # Extract and evaluate arguments
                    args_tokens = tokens[args_start:args_end]
                    args = self._evaluate_function_args(args_tokens)
                    
                    # Call the function
                    result = self._call_function(func_name, args)
                    output_queue.append(result)
                    
                    i = args_end  # Skip to after closing paren
                else:
                    # Function without parentheses (like PI)
                    result = self._call_function(func_name, [])
                    output_queue.append(result)
            
            elif token.type == Token.OPERATOR:
                # Handle operators
                while (operator_stack and 
                       operator_stack[-1].type == Token.OPERATOR and
                       self._precedence(operator_stack[-1]) >= self._precedence(token)):
                    op = operator_stack.pop()
                    right = output_queue.pop() if output_queue else 0
                    left = output_queue.pop() if output_queue else 0
                    result = self._apply_operator(op, left, right)
                    output_queue.append(result)
                operator_stack.append(token)
            
            elif token.type == Token.SUBEXPR:
                if token.subtype == "OPEN":
                    operator_stack.append(token)
                elif token.subtype == "CLOSE":
                    while (operator_stack and 
                           operator_stack[-1].type != Token.SUBEXPR):
                        op = operator_stack.pop()
                        right = output_queue.pop() if output_queue else 0
                        left = output_queue.pop() if output_queue else 0
                        result = self._apply_operator(op, left, right)
                        output_queue.append(result)
                    if operator_stack:
                        operator_stack.pop()  # Remove opening paren
            
            i += 1
        
        # Process remaining operators
        while operator_stack:
            op = operator_stack.pop()
            if op.type == Token.OPERATOR:
                right = output_queue.pop() if output_queue else 0
                left = output_queue.pop() if output_queue else 0
                result = self._apply_operator(op, left, right)
                output_queue.append(result)
        
        return output_queue[0] if output_queue else ""
    
    def _evaluate_operand(self, token: Token) -> Any:
        """Evaluate a single operand token."""
        if token.subtype == Token.NUMBER:
            try:
                return int(token.value) if '.' not in token.value else float(token.value)
            except ValueError:
                return 0
        
        elif token.subtype == Token.TEXT:
            return token.value
        
        elif token.subtype == Token.REFERENCE:
            # Cell reference like A1, B2
            return self._get_cell_value(token.value)
        
        elif token.subtype == Token.RANGE:
            # Range like A1:B2
            return self._get_range_values(token.value)
        
        elif token.subtype == Token.LOGICAL:
            return token.value.upper() == "TRUE"
        
        elif token.subtype == Token.ERROR:
            raise ValueErrorExcel()
        
        else:
            return token.value
    
    def _get_cell_value(self, cell_ref: str) -> Any:
        """Get value from a cell reference."""
        if not self.worksheet:
            return 0
        
        # Parse cell reference (e.g., A1, $B$2)
        match = re.match(r'(\$?)([A-Z]+)(\$?)(\d+)', cell_ref)
        if not match:
            return 0
        
        col_letters = match.group(2)
        row_num = int(match.group(4))
        
        # Convert column letters to number
        col_num = 0
        for i, letter in enumerate(reversed(col_letters)):
            col_num += (ord(letter) - ord('A') + 1) * (26 ** i)
        
        # Get cell from worksheet
        cell = self.worksheet._cells.get((row_num, col_num))
        if not cell:
            return 0
        
        # If it's a formula, evaluate it recursively
        if cell.is_formula():
            try:
                return self.evaluate(cell.formula, cell_ref)
            except CircularReferenceError:
                return "#CIRCULAR!"
        
        return cell.value if cell.value is not None else 0
    
    def _get_range_values(self, range_ref: str) -> List[Any]:
        """Get values from a range reference."""
        if ':' not in range_ref:
            return [self._get_cell_value(range_ref)]
        
        start_ref, end_ref = range_ref.split(':')
        
        # Parse start and end references
        start_match = re.match(r'(\$?)([A-Z]+)(\$?)(\d+)', start_ref)
        end_match = re.match(r'(\$?)([A-Z]+)(\$?)(\d+)', end_ref)
        
        if not start_match or not end_match:
            return []
        
        # Convert to column/row numbers
        start_col = sum((ord(c) - ord('A') + 1) * (26 ** i) 
                       for i, c in enumerate(reversed(start_match.group(2))))
        start_row = int(start_match.group(4))
        
        end_col = sum((ord(c) - ord('A') + 1) * (26 ** i) 
                     for i, c in enumerate(reversed(end_match.group(2))))
        end_row = int(end_match.group(4))
        
        # Collect values from range
        values = []
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                cell_ref = f"{self._col_num_to_letter(col)}{row}"
                values.append(self._get_cell_value(cell_ref))
        
        return values
    
    def _col_num_to_letter(self, col_num: int) -> str:
        """Convert column number to letter."""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    def _find_matching_paren(self, tokens: List[Token], start_pos: int) -> int:
        """Find the matching closing parenthesis."""
        paren_count = 1
        pos = start_pos + 1
        
        while pos < len(tokens) and paren_count > 0:
            token = tokens[pos]
            if token.type == Token.SUBEXPR:
                if token.subtype == "OPEN":
                    paren_count += 1
                elif token.subtype == "CLOSE":
                    paren_count -= 1
            pos += 1
        
        return pos - 1  # Position of closing paren
    
    def _evaluate_function_args(self, tokens: List[Token]) -> List[Any]:
        """Evaluate function arguments."""
        if not tokens:
            return []
        
        args = []
        current_arg = []
        paren_depth = 0
        
        for token in tokens:
            if token.type == Token.ARGUMENT and paren_depth == 0:
                # End of current argument
                if current_arg:
                    arg_result = self._evaluate_tokens(current_arg)
                    args.append(arg_result)
                current_arg = []
            else:
                if token.type == Token.SUBEXPR:
                    if token.subtype == "OPEN":
                        paren_depth += 1
                    elif token.subtype == "CLOSE":
                        paren_depth -= 1
                current_arg.append(token)
        
        # Add final argument
        if current_arg:
            arg_result = self._evaluate_tokens(current_arg)
            args.append(arg_result)
        
        return args
    
    def _call_function(self, func_name: str, args: List[Any]) -> Any:
        """Call a built-in function."""
        if func_name in BUILTIN_FUNCTIONS:
            func = BUILTIN_FUNCTIONS[func_name]
            try:
                return func(*args)
            except Exception as e:
                if isinstance(e, ExcelError):
                    return str(e)
                else:
                    return "#VALUE!"
        else:
            return f"#NAME?"
    
    def _precedence(self, token: Token) -> int:
        """Get operator precedence."""
        precedences = {
            '^': 4,
            '*': 3, '/': 3,
            '+': 2, '-': 2,
            '&': 2,
            '=': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '<>': 1,
        }
        return precedences.get(token.value, 0)
    
    def _apply_operator(self, op_token: Token, left: Any, right: Any) -> Any:
        """Apply an operator to two operands."""
        op = op_token.value
        
        try:
            if op == '+':
                return float(left) + float(right)
            elif op == '-':
                return float(left) - float(right)
            elif op == '*':
                return float(left) * float(right)
            elif op == '/':
                if float(right) == 0:
                    raise DivisionByZeroError()
                return float(left) / float(right)
            elif op == '^':
                return float(left) ** float(right)
            elif op == '&':
                return str(left) + str(right)
            elif op == '=':
                return left == right
            elif op == '<':
                return float(left) < float(right)
            elif op == '>':
                return float(left) > float(right)
            elif op == '<=':
                return float(left) <= float(right)
            elif op == '>=':
                return float(left) >= float(right)
            elif op == '<>':
                return left != right
            else:
                return 0
        except (ValueError, TypeError):
            raise ValueErrorExcel()
        except ZeroDivisionError:
            raise DivisionByZeroError()