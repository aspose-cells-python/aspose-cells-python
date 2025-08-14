"""
Excel Formula Tokenizer - Based on opencells tokenizer
Converts Excel formulas into token streams for evaluation.
"""

import re
from typing import List, Optional


class Token:
    """Represents a single token in a formula."""
    
    # Token types
    LITERAL = "LITERAL"
    OPERAND = "OPERAND" 
    FUNCTION = "FUNCTION"
    SUBEXPR = "SUBEXPR"
    ARGUMENT = "ARGUMENT"
    OPERATOR = "OPERATOR"
    WHITESPACE = "WHITESPACE"
    ERROR = "ERROR"
    
    # Token subtypes
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    LOGICAL = "LOGICAL"
    RANGE = "RANGE"
    REFERENCE = "REFERENCE"
    NAME = "NAME"
    
    # Operator types
    MATH = "MATH"
    CONCAT = "CONCAT"
    INTERSECT = "INTERSECT"
    UNION = "UNION"
    
    def __init__(self, value: str, type_: str, subtype: str = ""):
        self.value = value
        self.type = type_
        self.subtype = subtype
    
    def __repr__(self):
        return f"Token({self.value!r}, {self.type}, {self.subtype})"


class Tokenizer:
    """Tokenizer for Excel formulas."""
    
    # Regex patterns
    CELL_REF_PATTERN = re.compile(r'^(\$?)([A-Z]+)(\$?)(\d+)$')
    RANGE_PATTERN = re.compile(r'^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$')
    FUNCTION_PATTERN = re.compile(r'^[A-Z_][A-Z0-9_.]*$')
    NUMBER_PATTERN = re.compile(r'^-?\d+(\.\d*)?([Ee][+-]?\d+)?$|^-?\d*\.\d+([Ee][+-]?\d+)?$')
    
    # Excel error codes
    ERROR_CODES = {'#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'}
    
    def __init__(self, formula: str):
        self.formula = formula.strip()
        self.tokens: List[Token] = []
        self.position = 0
        self._tokenize()
    
    def _tokenize(self):
        """Parse the formula into tokens."""
        if not self.formula:
            return
        
        # Skip leading = if present
        if self.formula.startswith('='):
            self.position = 1
        
        while self.position < len(self.formula):
            self._skip_whitespace()
            if self.position >= len(self.formula):
                break
                
            if self._try_string():
                continue
            elif self._try_number():
                continue
            elif self._try_operator():
                continue
            elif self._try_function():
                continue
            elif self._try_reference():
                continue
            elif self._try_error():
                continue
            elif self._try_parenthesis():
                continue
            elif self._try_separator():
                continue
            else:
                # Unknown character, treat as text
                self._consume_text()
    
    def _current_char(self) -> Optional[str]:
        """Get current character."""
        if self.position < len(self.formula):
            return self.formula[self.position]
        return None
    
    def _peek_char(self, offset: int = 1) -> Optional[str]:
        """Peek ahead at character."""
        pos = self.position + offset
        if pos < len(self.formula):
            return self.formula[pos]
        return None
    
    def _skip_whitespace(self):
        """Skip whitespace characters."""
        while self.position < len(self.formula) and self.formula[self.position].isspace():
            self.position += 1
    
    def _try_string(self) -> bool:
        """Try to parse a quoted string."""
        char = self._current_char()
        if char not in ('"', "'"):
            return False
        
        quote_char = char
        start_pos = self.position
        self.position += 1  # Skip opening quote
        value = ''
        
        while self.position < len(self.formula):
            char = self._current_char()
            if char == quote_char:
                # Check for escaped quote (doubled quotes)
                if self._peek_char() == quote_char:
                    value += quote_char
                    self.position += 2
                else:
                    self.position += 1  # Skip closing quote
                    break
            else:
                value += char
                self.position += 1
        
        self.tokens.append(Token(value, Token.OPERAND, Token.TEXT))
        return True
    
    def _try_number(self) -> bool:
        """Try to parse a number."""
        start_pos = self.position
        value = ''
        
        # Only handle negative sign if it's at the start or after an operator/opening paren
        can_be_negative = (
            len(self.tokens) == 0 or  # Start of formula
            self.tokens[-1].type in (Token.OPERATOR, Token.SUBEXPR, Token.ARGUMENT)
        )
        
        # Handle negative sign only in valid contexts
        if self._current_char() == '-' and can_be_negative:
            value += '-'
            self.position += 1
        elif self._current_char() == '-':
            # It's likely a subtraction operator, not a negative number
            return False
        
        # Collect digits and decimal points
        while self.position < len(self.formula):
            char = self._current_char()
            if char.isdigit() or char == '.':
                value += char
                self.position += 1
            elif char in 'Ee' and value and value[-1].isdigit():
                # Scientific notation
                value += char
                self.position += 1
                # Handle optional +/- after E
                if self._current_char() in '+-':
                    value += self._current_char()
                    self.position += 1
            else:
                break
        
        if value and self.NUMBER_PATTERN.match(value):
            self.tokens.append(Token(value, Token.OPERAND, Token.NUMBER))
            return True
        else:
            # Not a valid number, reset position
            self.position = start_pos
            return False
    
    def _try_operator(self) -> bool:
        """Try to parse an operator."""
        char = self._current_char()
        operators = {
            '+': (Token.OPERATOR, Token.MATH),
            '-': (Token.OPERATOR, Token.MATH),
            '*': (Token.OPERATOR, Token.MATH),
            '/': (Token.OPERATOR, Token.MATH),
            '^': (Token.OPERATOR, Token.MATH),
            '&': (Token.OPERATOR, Token.CONCAT),
            '=': (Token.OPERATOR, Token.MATH),
            '<': (Token.OPERATOR, Token.MATH),
            '>': (Token.OPERATOR, Token.MATH),
            '%': (Token.OPERATOR, Token.MATH),
        }
        
        if char in operators:
            # Check for multi-character operators
            next_char = self._peek_char()
            if char == '<' and next_char == '>':
                self.tokens.append(Token('<>', Token.OPERATOR, Token.MATH))
                self.position += 2
            elif char == '<' and next_char == '=':
                self.tokens.append(Token('<=', Token.OPERATOR, Token.MATH))
                self.position += 2
            elif char == '>' and next_char == '=':
                self.tokens.append(Token('>=', Token.OPERATOR, Token.MATH))
                self.position += 2
            else:
                type_, subtype = operators[char]
                self.tokens.append(Token(char, type_, subtype))
                self.position += 1
            return True
        return False
    
    def _try_function(self) -> bool:
        """Try to parse a function name."""
        start_pos = self.position
        value = ''
        
        # Functions start with letter or underscore
        char = self._current_char()
        if not (char.isalpha() or char == '_'):
            return False
        
        # Collect function name
        while self.position < len(self.formula):
            char = self._current_char()
            if char.isalnum() or char in '_.':
                value += char
                self.position += 1
            else:
                break
        
        # Check if followed by opening parenthesis
        self._skip_whitespace()
        if self._current_char() == '(':
            if self.FUNCTION_PATTERN.match(value.upper()):
                self.tokens.append(Token(value.upper(), Token.FUNCTION))
                return True
        
        # Not a function, reset position
        self.position = start_pos
        return False
    
    def _try_reference(self) -> bool:
        """Try to parse a cell reference or range."""
        start_pos = self.position
        value = ''
        
        # Collect potential reference
        while self.position < len(self.formula):
            char = self._current_char()
            if char.isalnum() or char in '$:':
                value += char
                self.position += 1
            else:
                break
        
        if value:
            # Check for range (contains colon)
            if ':' in value and self.RANGE_PATTERN.match(value.upper()):
                self.tokens.append(Token(value.upper(), Token.OPERAND, Token.RANGE))
                return True
            # Check for single cell reference
            elif self.CELL_REF_PATTERN.match(value.upper()):
                self.tokens.append(Token(value.upper(), Token.OPERAND, Token.REFERENCE))
                return True
        
        # Not a reference, reset position
        self.position = start_pos
        return False
    
    def _try_error(self) -> bool:
        """Try to parse an error value."""
        start_pos = self.position
        if self._current_char() != '#':
            return False
        
        value = ''
        while self.position < len(self.formula):
            char = self._current_char()
            if char.isalnum() or char in '#!/?':
                value += char
                self.position += 1
            else:
                break
        
        if value in self.ERROR_CODES:
            self.tokens.append(Token(value, Token.OPERAND, Token.ERROR))
            return True
        else:
            self.position = start_pos
            return False
    
    def _try_parenthesis(self) -> bool:
        """Try to parse parentheses."""
        char = self._current_char()
        if char == '(':
            self.tokens.append(Token('(', Token.SUBEXPR, "OPEN"))
            self.position += 1
            return True
        elif char == ')':
            self.tokens.append(Token(')', Token.SUBEXPR, "CLOSE"))
            self.position += 1
            return True
        return False
    
    def _try_separator(self) -> bool:
        """Try to parse separators (comma, semicolon)."""
        char = self._current_char()
        if char in (',', ';'):
            self.tokens.append(Token(char, Token.ARGUMENT))
            self.position += 1
            return True
        return False
    
    def _consume_text(self):
        """Consume remaining text as literal."""
        value = ''
        while self.position < len(self.formula):
            char = self._current_char()
            if char.isspace() or char in '()+-*/^&=<>%,;':
                break
            value += char
            self.position += 1
        
        if value:
            self.tokens.append(Token(value, Token.OPERAND, Token.TEXT))
    
    def __iter__(self):
        """Iterate over tokens."""
        return iter(self.tokens)