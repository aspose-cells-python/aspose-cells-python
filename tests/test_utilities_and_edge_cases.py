"""
Utilities and Edge Cases Tests - Simplified
Test utility functions and edge cases with minimal code.
"""

import pytest
from aspose.cells import Workbook, FileFormat
from aspose.cells.utils.coordinates import (
    column_index_to_letter, 
    column_letter_to_index,
    coordinate_to_tuple,
    tuple_to_coordinate,
    parse_range
)
from aspose.cells.utils.validation import (
    is_numeric_string,
    is_formula,
    validate_sheet_name,
    sanitize_sheet_name,
    validate_cell_reference,
    infer_data_type
)
from aspose.cells.utils.exceptions import InvalidCoordinateError, FileFormatError


class TestCoordinateUtils:
    """Test coordinate utility functions."""
    
    def test_column_conversions(self):
        """Test column index/letter conversions."""
        # Basic conversions
        assert column_index_to_letter(1) == "A"
        assert column_index_to_letter(26) == "Z"
        assert column_index_to_letter(27) == "AA"
        
        assert column_letter_to_index("A") == 1
        assert column_letter_to_index("Z") == 26
        assert column_letter_to_index("AA") == 27
        
        # Round trip test
        for i in [1, 26, 27, 100, 702]:
            letter = column_index_to_letter(i)
            back = column_letter_to_index(letter)
            assert back == i
    
    def test_coordinate_conversions(self):
        """Test coordinate conversions."""
        assert coordinate_to_tuple("A1") == (1, 1)
        assert coordinate_to_tuple("B2") == (2, 2)
        assert coordinate_to_tuple("AA10") == (10, 27)
        
        assert tuple_to_coordinate(1, 1) == "A1"
        assert tuple_to_coordinate(2, 2) == "B2"
        assert tuple_to_coordinate(10, 27) == "AA10"
    
    def test_range_parsing(self):
        """Test range parsing."""
        start, end = parse_range("A1:B2")
        assert start == (1, 1)
        assert end == (2, 2)
        
        with pytest.raises(InvalidCoordinateError):
            parse_range("INVALID")


class TestValidationUtils:
    """Test validation utility functions."""
    
    def test_numeric_string_detection(self):
        """Test numeric string detection."""
        assert is_numeric_string("123") is True
        assert is_numeric_string("123.45") is True
        assert is_numeric_string("-123") is True
        assert is_numeric_string("abc") is False
        assert is_numeric_string("12abc") is False
        assert is_numeric_string(123) is True  # Numbers are numeric
    
    def test_formula_detection(self):
        """Test formula detection."""
        assert is_formula("=A1+B1") is True
        assert is_formula("=SUM(A1:A10)") is True
        assert is_formula("A1+B1") is False
        assert is_formula("123") is False
    
    def test_sheet_name_validation(self):
        """Test sheet name validation."""
        assert validate_sheet_name("Sheet1") is True
        assert validate_sheet_name("Data_2024") is True
        assert validate_sheet_name("") is False
        assert validate_sheet_name("A" * 32) is False  # Too long
        assert validate_sheet_name("Sheet\\1") is False  # Invalid char
    
    def test_sheet_name_sanitization(self):
        """Test sheet name sanitization."""
        assert sanitize_sheet_name("Sheet1") == "Sheet1"
        assert sanitize_sheet_name("Sheet\\1") == "Sheet_1"
        assert sanitize_sheet_name("") == "Sheet1"
        assert len(sanitize_sheet_name("A" * 40)) == 31  # Truncated
    
    def test_cell_reference_validation(self):
        """Test cell reference validation."""
        assert validate_cell_reference("A1") is True
        assert validate_cell_reference("Z99") is True
        assert validate_cell_reference("AA100") is True
        assert validate_cell_reference("") is False
        assert validate_cell_reference("A0") is False
        assert validate_cell_reference("123") is False
    
    def test_data_type_inference(self):
        """Test data type inference."""
        assert infer_data_type(None) == "empty"
        assert infer_data_type(True) == "boolean"
        assert infer_data_type(123) == "number"
        assert infer_data_type("text") == "string"
        assert infer_data_type("=SUM(A1:A10)") == "formula"
        assert infer_data_type("123") == "number"  # Numeric string


class TestFileFormats:
    """Test FileFormat functionality."""
    
    def test_format_enum(self):
        """Test FileFormat enum."""
        assert FileFormat.XLSX.value == "xlsx"
        assert FileFormat.CSV.value == "csv"
        assert FileFormat.JSON.value == "json"
        assert FileFormat.MARKDOWN.value == "markdown"
    
    def test_format_from_extension(self):
        """Test format inference from extension."""
        assert FileFormat.from_extension("test.xlsx") == FileFormat.XLSX
        assert FileFormat.from_extension("test.csv") == FileFormat.CSV
        assert FileFormat.from_extension("test.json") == FileFormat.JSON
        assert FileFormat.from_extension("test.md") == FileFormat.MARKDOWN
        assert FileFormat.from_extension("test.unknown") == FileFormat.XLSX  # Default
    
    def test_format_properties(self):
        """Test format properties."""
        assert FileFormat.XLSX.extension == ".xlsx"
        assert FileFormat.CSV.extension == ".csv"
        assert FileFormat.JSON.extension == ".json"
        assert FileFormat.MARKDOWN.extension == ".md"
        
        # Test mime types exist
        for fmt in FileFormat.get_supported_formats():
            assert isinstance(fmt.mime_type, str)
            assert len(fmt.mime_type) > 0


class TestExceptions:
    """Test custom exceptions."""
    
    def test_exception_creation(self):
        """Test exception creation and messages."""
        with pytest.raises(InvalidCoordinateError) as exc:
            raise InvalidCoordinateError("Invalid coordinate")
        assert "Invalid coordinate" in str(exc.value)
        
        with pytest.raises(FileFormatError) as exc:
            raise FileFormatError("Unsupported format")
        assert "Unsupported format" in str(exc.value)


class TestEdgeCases:
    """Test edge cases and error conditions."""
    
    def test_empty_workbook(self):
        """Test operations on empty workbook."""
        wb = Workbook()
        assert wb.active is not None
        assert len(wb.worksheets) == 1
        assert wb.active.max_row == 0
        wb.close()
    
    def test_unicode_handling(self):
        """Test Unicode character handling."""
        wb = Workbook()
        ws = wb.active
        
        unicode_texts = ["English", "中文", "العربية", "🚀🌟"]
        
        for i, text in enumerate(unicode_texts, 1):
            ws.cell(i, 1, text)
            assert ws.cell(i, 1).value == text
        
        wb.close()
    
    def test_special_characters(self):
        """Test special characters in values."""
        wb = Workbook()
        ws = wb.active
        
        special_chars = [
            "Text with\nnewline",
            "Text with\"quotes\"",
            "Text with|pipe",
            "Text with,comma"
        ]
        
        for i, text in enumerate(special_chars, 1):
            ws.cell(i, 1, text)
            assert ws.cell(i, 1).value == text
        
        wb.close()
    
    def test_extreme_values(self):
        """Test extreme numeric values."""
        wb = Workbook()
        ws = wb.active
        
        values = [0, -0, 1, -1, 1e10, -1e10, 1e-10]
        
        for i, value in enumerate(values, 1):
            ws.cell(i, 1, value)
            stored = ws.cell(i, 1).value
            assert stored == value or abs(stored - value) < 1e-15
        
        wb.close()
    
    def test_formula_edge_cases(self):
        """Test formula edge cases."""
        wb = Workbook()
        ws = wb.active
        
        formulas = ["=A1", "=A1+B1", "=SUM(A1:A10)", "= A1 + B1 "]
        
        for i, formula in enumerate(formulas, 1):
            ws.cell(i, 1, formula)
            assert ws.cell(i, 1).value == formula
        
        wb.close()
    
    def test_large_data_handling(self):
        """Test handling larger datasets."""
        wb = Workbook()
        ws = wb.active
        
        # Add 50x5 grid of data
        for row in range(1, 51):
            for col in range(1, 6):
                ws.cell(row, col, f"R{row}C{col}")
        
        assert ws.cell(1, 1).value == "R1C1"
        assert ws.cell(50, 5).value == "R50C5"
        assert ws.max_row == 50
        assert ws.max_column == 5
        
        wb.close()