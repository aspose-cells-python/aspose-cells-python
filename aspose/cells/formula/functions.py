"""
Excel Function Library - Built-in functions for formula evaluation.
"""

import math
import datetime
from typing import Union, Any, List, Callable
from decimal import Decimal


# Type aliases
Number = Union[int, float, Decimal]
Value = Union[Number, str, bool, datetime.datetime, None]


class ExcelError(Exception):
    """Base class for Excel errors."""
    pass


class DivisionByZeroError(ExcelError):
    """#DIV/0! error."""
    def __str__(self):
        return "#DIV/0!"


class ValueErrorExcel(ExcelError):
    """#VALUE! error."""
    def __str__(self):
        return "#VALUE!"


class NumError(ExcelError):
    """#NUM! error."""
    def __str__(self):
        return "#NUM!"


class NameError(ExcelError):
    """#NAME? error."""
    def __str__(self):
        return "#NAME?"


def to_number(value: Value) -> Number:
    """Convert value to number, raising #VALUE! if not possible."""
    if isinstance(value, (int, float, Decimal)):
        return value
    elif isinstance(value, bool):
        return 1 if value else 0
    elif isinstance(value, str):
        try:
            return float(value) if '.' in value else int(value)
        except ValueError:
            raise ValueErrorExcel()
    else:
        raise ValueErrorExcel()


def to_text(value: Value) -> str:
    """Convert value to text."""
    if value is None:
        return ""
    elif isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    else:
        return str(value)


def to_boolean(value: Value) -> bool:
    """Convert value to boolean."""
    if isinstance(value, bool):
        return value
    elif isinstance(value, (int, float, Decimal)):
        return value != 0
    elif isinstance(value, str):
        upper = value.upper()
        if upper == "TRUE":
            return True
        elif upper == "FALSE":
            return False
        else:
            raise ValueErrorExcel()
    else:
        raise ValueErrorExcel()


# Mathematical Functions
def func_abs(value: Value) -> Number:
    """ABS function - absolute value."""
    return abs(to_number(value))


def func_sum(*args: Value) -> Number:
    """SUM function - sum of values."""
    total = 0
    for arg in args:
        if isinstance(arg, (list, tuple)):
            total += func_sum(*arg)
        else:
            try:
                total += to_number(arg)
            except (ValueErrorExcel, TypeError):
                continue  # Skip non-numeric values
    return total


def func_average(*args: Value) -> Number:
    """AVERAGE function - average of values."""
    values = []
    for arg in args:
        if isinstance(arg, (list, tuple)):
            values.extend([v for v in arg if isinstance(v, (int, float, Decimal))])
        else:
            try:
                values.append(to_number(arg))
            except (ValueErrorExcel, TypeError):
                continue
    
    if not values:
        raise DivisionByZeroError()
    
    return sum(values) / len(values)


def func_count(*args: Value) -> int:
    """COUNT function - count of numeric values."""
    count = 0
    for arg in args:
        if isinstance(arg, (list, tuple)):
            count += func_count(*arg)
        else:
            try:
                to_number(arg)
                count += 1
            except (ValueErrorExcel, TypeError):
                continue
    return count


def func_counta(*args: Value) -> int:
    """COUNTA function - count of non-empty values."""
    count = 0
    for arg in args:
        if isinstance(arg, (list, tuple)):
            count += func_counta(*arg)
        else:
            if arg is not None and arg != "":
                count += 1
    return count


def func_max(*args: Value) -> Number:
    """MAX function - maximum value."""
    values = []
    for arg in args:
        if isinstance(arg, (list, tuple)):
            values.extend([v for v in arg if isinstance(v, (int, float, Decimal))])
        else:
            try:
                values.append(to_number(arg))
            except (ValueErrorExcel, TypeError):
                continue
    
    if not values:
        return 0
    
    return max(values)


def func_min(*args: Value) -> Number:
    """MIN function - minimum value."""
    values = []
    for arg in args:
        if isinstance(arg, (list, tuple)):
            values.extend([v for v in arg if isinstance(v, (int, float, Decimal))])
        else:
            try:
                values.append(to_number(arg))
            except (ValueErrorExcel, TypeError):
                continue
    
    if not values:
        return 0
    
    return min(values)


def func_round(number: Value, digits: Value = 0) -> Number:
    """ROUND function - round to specified digits."""
    num = to_number(number)
    dig = int(to_number(digits))
    return round(num, dig)


def func_power(number: Value, power: Value) -> Number:
    """POWER function - raise to power."""
    num = to_number(number)
    pow_val = to_number(power)
    return num ** pow_val


def func_sqrt(number: Value) -> Number:
    """SQRT function - square root."""
    num = to_number(number)
    if num < 0:
        raise NumError()
    return math.sqrt(num)


def func_exp(number: Value) -> Number:
    """EXP function - e raised to power."""
    num = to_number(number)
    return math.exp(num)


def func_ln(number: Value) -> Number:
    """LN function - natural logarithm."""
    num = to_number(number)
    if num <= 0:
        raise NumError()
    return math.log(num)


def func_log10(number: Value) -> Number:
    """LOG10 function - base-10 logarithm."""
    num = to_number(number)
    if num <= 0:
        raise NumError()
    return math.log10(num)


# Trigonometric Functions
def func_sin(number: Value) -> Number:
    """SIN function."""
    return math.sin(to_number(number))


def func_cos(number: Value) -> Number:
    """COS function."""
    return math.cos(to_number(number))


def func_tan(number: Value) -> Number:
    """TAN function."""
    return math.tan(to_number(number))


def func_pi() -> Number:
    """PI function."""
    return math.pi


# Logical Functions
def func_if(condition: Value, true_value: Value, false_value: Value = False) -> Value:
    """IF function."""
    try:
        if to_boolean(condition):
            return true_value
        else:
            return false_value
    except ValueErrorExcel:
        return false_value


def func_and(*args: Value) -> bool:
    """AND function."""
    for arg in args:
        if not to_boolean(arg):
            return False
    return True


def func_or(*args: Value) -> bool:
    """OR function."""
    for arg in args:
        if to_boolean(arg):
            return True
    return False


def func_not(value: Value) -> bool:
    """NOT function."""
    return not to_boolean(value)


def func_true() -> bool:
    """TRUE function."""
    return True


def func_false() -> bool:
    """FALSE function."""
    return False


# Text Functions
def func_concatenate(*args: Value) -> str:
    """CONCATENATE function."""
    return "".join(to_text(arg) for arg in args)


def func_len(text: Value) -> int:
    """LEN function - length of text."""
    return len(to_text(text))


def func_left(text: Value, num_chars: Value = 1) -> str:
    """LEFT function - leftmost characters."""
    text_str = to_text(text)
    num = int(to_number(num_chars))
    return text_str[:num]


def func_right(text: Value, num_chars: Value = 1) -> str:
    """RIGHT function - rightmost characters."""
    text_str = to_text(text)
    num = int(to_number(num_chars))
    return text_str[-num:] if num > 0 else ""


def func_mid(text: Value, start_pos: Value, num_chars: Value) -> str:
    """MID function - substring."""
    text_str = to_text(text)
    start = int(to_number(start_pos)) - 1  # Excel is 1-based
    num = int(to_number(num_chars))
    return text_str[start:start+num]


def func_upper(text: Value) -> str:
    """UPPER function - convert to uppercase."""
    return to_text(text).upper()


def func_lower(text: Value) -> str:
    """LOWER function - convert to lowercase."""
    return to_text(text).lower()


def func_trim(text: Value) -> str:
    """TRIM function - remove extra spaces."""
    return " ".join(to_text(text).split())


# Date Functions
def func_today() -> datetime.date:
    """TODAY function - current date."""
    return datetime.date.today()


def func_now() -> datetime.datetime:
    """NOW function - current date and time."""
    return datetime.datetime.now()


def func_year(date_value: Value) -> int:
    """YEAR function - year from date."""
    if isinstance(date_value, datetime.datetime):
        return date_value.year
    elif isinstance(date_value, datetime.date):
        return date_value.year
    else:
        raise ValueErrorExcel()


def func_month(date_value: Value) -> int:
    """MONTH function - month from date."""
    if isinstance(date_value, datetime.datetime):
        return date_value.month
    elif isinstance(date_value, datetime.date):
        return date_value.month
    else:
        raise ValueErrorExcel()


def func_day(date_value: Value) -> int:
    """DAY function - day from date."""
    if isinstance(date_value, datetime.datetime):
        return date_value.day
    elif isinstance(date_value, datetime.date):
        return date_value.day
    else:
        raise ValueErrorExcel()


# Registry of all built-in functions
BUILTIN_FUNCTIONS: dict[str, Callable] = {
    # Math functions
    'ABS': func_abs,
    'SUM': func_sum,
    'AVERAGE': func_average,
    'COUNT': func_count,
    'COUNTA': func_counta,
    'MAX': func_max,
    'MIN': func_min,
    'ROUND': func_round,
    'POWER': func_power,
    'SQRT': func_sqrt,
    'EXP': func_exp,
    'LN': func_ln,
    'LOG10': func_log10,
    
    # Trigonometric
    'SIN': func_sin,
    'COS': func_cos,
    'TAN': func_tan,
    'PI': func_pi,
    
    # Logical
    'IF': func_if,
    'AND': func_and,
    'OR': func_or,
    'NOT': func_not,
    'TRUE': func_true,
    'FALSE': func_false,
    
    # Text
    'CONCATENATE': func_concatenate,
    'LEN': func_len,
    'LEFT': func_left,
    'RIGHT': func_right,
    'MID': func_mid,
    'UPPER': func_upper,
    'LOWER': func_lower,
    'TRIM': func_trim,
    
    # Date
    'TODAY': func_today,
    'NOW': func_now,
    'YEAR': func_year,
    'MONTH': func_month,
    'DAY': func_day,
}