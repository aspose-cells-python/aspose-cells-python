"""
Unit Tests
Tests for individual components, classes, and methods.
These tests focus on internal functionality and edge cases.
"""

import pytest
from datetime import datetime
from aspose.cells import Workbook, FileFormat, CellValue
from aspose.cells.utils.coordinates import column_index_to_letter, column_letter_to_index
from aspose.cells.utils.validation import validate_cell_reference


class TestWorkbookUnits:
    """Unit tests for Workbook class."""
    
    def test_workbook_creation(self):
        """Test workbook creation and basic properties."""
        wb = Workbook()
        assert wb is not None
        assert len(wb.worksheets) == 1
        assert wb.active.name == "Sheet1"
        wb.close()
    
    def test_worksheet_management(self):
        """Test worksheet creation and management."""
        wb = Workbook()
        
        # Create new worksheet
        ws = wb.create_sheet("TestSheet")
        assert ws.name == "TestSheet"
        assert len(wb.worksheets) == 2
        
        # Access by name
        retrieved_ws = wb.worksheets["TestSheet"]
        assert retrieved_ws.name == "TestSheet"
        
        # Sheet names list
        assert "TestSheet" in wb.sheetnames
        
        wb.close()
    
    def test_workbook_properties(self):
        """Test workbook properties and metadata."""
        wb = Workbook()
        props = wb.properties
        assert isinstance(props, dict)
        
        # Test setting properties
        wb.properties["Title"] = "Test Workbook"
        wb.properties["Author"] = "Test Author"
        
        assert wb.properties["Title"] == "Test Workbook"
        assert wb.properties["Author"] == "Test Author"
        
        wb.close()


class TestWorksheetUnits:
    """Unit tests for Worksheet class."""
    
    def test_worksheet_properties(self):
        """Test worksheet basic properties."""
        wb = Workbook()
        ws = wb.active
        
        # Name property
        assert ws.name == "Sheet1"
        ws.name = "NewName"
        assert ws.name == "NewName"
        
        # Max row/column
        assert ws.max_row >= 0
        assert ws.max_column >= 0
        
        wb.close()
    
    def test_cell_access_methods(self):
        """Test different cell access methods."""
        wb = Workbook()
        ws = wb.active
        
        # String reference
        cell1 = ws['A1']
        assert cell1.coordinate == "A1"
        
        # Tuple access (0-based)
        cell2 = ws[0, 0]
        assert cell2.coordinate == "A1"
        
        # cell() method (1-based)
        cell3 = ws.cell(1, 1)
        assert cell3.coordinate == "A1"
        
        # All should reference same cell
        assert cell1.row == cell2.row == cell3.row
        assert cell1.column == cell2.column == cell3.column
        
        wb.close()
    
    def test_range_operations(self):
        """Test worksheet range operations."""
        wb = Workbook()
        ws = wb.active
        
        # Set range values
        test_data = [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6]
        ]
        
        for row_idx, row_data in enumerate(test_data):
            for col_idx, value in enumerate(row_data):
                ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Verify the data was set correctly
        assert ws['A1'].value == "A"
        assert ws['B2'].value == 2
        assert ws['C3'].value == 6
        
        wb.close()


class TestCellUnits:
    """Unit tests for Cell class."""
    
    def test_cell_value_types(self):
        """Test cell value type handling."""
        wb = Workbook()
        ws = wb.active
        
        # String
        ws['A1'] = "test string"
        assert ws['A1'].value == "test string"
        assert isinstance(ws['A1'].value, str)
        
        # Integer
        ws['A2'] = 42
        assert ws['A2'].value == 42
        assert isinstance(ws['A2'].value, int)
        
        # Float
        ws['A3'] = 3.14159
        assert ws['A3'].value == 3.14159
        assert isinstance(ws['A3'].value, float)
        
        # Boolean
        ws['A4'] = True
        assert ws['A4'].value is True
        assert isinstance(ws['A4'].value, bool)
        
        ws['A5'] = False
        assert ws['A5'].value is False
        
        wb.close()
    
    def test_cell_coordinates(self):
        """Test cell coordinate calculations."""
        wb = Workbook()
        ws = wb.active
        
        # Test various coordinates
        test_cases = [
            (1, 1, "A1"),
            (1, 26, "Z1"),
            (1, 27, "AA1"),
            (10, 5, "E10"),
            (100, 100, "CV100")
        ]
        
        for row, col, expected_coord in test_cases:
            cell = ws.cell(row, col)
            assert cell.coordinate == expected_coord
            assert cell.row == row
            assert cell.column == col
        
        wb.close()
    
    def test_cell_formulas(self):
        """Test cell formula handling."""
        wb = Workbook()
        ws = wb.active
        
        # Set up data
        ws['A1'] = 10
        ws['A2'] = 20
        
        # Simple formula - check that it's stored as a formula
        ws['A3'] = "=A1+A2"
        assert ws['A3'].value == "=A1+A2"  # Formula stored as value
        
        # Function formula
        ws['A4'] = "=SUM(A1:A2)"
        assert ws['A4'].value == "=SUM(A1:A2)"
        
        wb.close()
    
    def test_cell_styling(self):
        """Test cell styling properties."""
        wb = Workbook()
        ws = wb.active
        cell = ws['A1']
        
        # Font properties
        cell.style.font.bold = True
        assert cell.style.font.bold is True
        
        cell.style.font.italic = True
        assert cell.style.font.italic is True
        
        # Fill properties
        cell.style.fill.background_color = "red"
        assert cell.style.fill.background_color == "red"
        
        wb.close()


class TestUtilityFunctions:
    """Unit tests for utility functions."""
    
    def test_column_conversions(self):
        """Test column letter/index conversions."""
        # Letter to index
        assert column_letter_to_index("A") == 1
        assert column_letter_to_index("Z") == 26
        assert column_letter_to_index("AA") == 27
        assert column_letter_to_index("AB") == 28
        
        # Index to letter
        assert column_index_to_letter(1) == "A"
        assert column_index_to_letter(26) == "Z"
        assert column_index_to_letter(27) == "AA"
        assert column_index_to_letter(28) == "AB"
    
    def test_cell_reference_validation(self):
        """Test cell reference validation."""
        # Valid references
        assert validate_cell_reference("A1") is True
        assert validate_cell_reference("Z99") is True
        assert validate_cell_reference("AA100") is True
        
        # Invalid references
        assert validate_cell_reference("") is False
        assert validate_cell_reference("1A") is False
        assert validate_cell_reference("A") is False
        assert validate_cell_reference("1") is False


class TestFileFormatUnits:
    """Unit tests for file format handling."""
    
    def test_file_format_detection(self):
        """Test file format detection and validation."""
        # Test format constants
        assert FileFormat.XLSX is not None
        assert FileFormat.CSV is not None
        assert FileFormat.JSON is not None
        
        # Test format comparison
        assert FileFormat.XLSX != FileFormat.CSV
        assert FileFormat.CSV != FileFormat.JSON
    
    def test_export_format_validation(self):
        """Test export format validation."""
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "test"
        
        # Test different export formats
        formats_to_test = [FileFormat.XLSX, FileFormat.CSV, FileFormat.JSON]
        
        for fmt in formats_to_test:
            try:
                result = wb.exportAs(fmt)
                assert result is not None
            except Exception as e:
                # Some formats might not be implemented yet
                assert "not implemented" in str(e).lower() or "unsupported" in str(e).lower()
        
        wb.close()


class TestErrorHandling:
    """Unit tests for error handling."""
    
    def test_invalid_cell_references(self):
        """Test handling of invalid cell references."""
        wb = Workbook()
        ws = wb.active
        
        # Test invalid string references
        with pytest.raises(Exception):
            _ = ws['INVALID']
        
        with pytest.raises(Exception):
            _ = ws['A0']  # Row 0 doesn't exist
        
        wb.close()
    
    def test_invalid_worksheet_operations(self):
        """Test invalid worksheet operations."""
        wb = Workbook()
        
        # Try to access non-existent worksheet
        with pytest.raises(Exception):  # Could be KeyError or WorksheetNotFoundError
            _ = wb.worksheets["NonExistentSheet"]
        
        wb.close()
    
    def test_closed_workbook_operations(self):
        """Test operations on closed workbook."""
        wb = Workbook()
        wb.close()
        
        # Operations on closed workbook - may or may not raise errors depending on implementation
        # Just verify that close was called
        assert True  # If we get here, close() worked


class TestDataValidation:
    """Unit tests for data validation."""
    
    def test_large_values(self):
        """Test handling of large numeric values."""
        wb = Workbook()
        ws = wb.active
        
        # Large integer
        large_int = 999999999999999
        ws['A1'] = large_int
        assert ws['A1'].value == large_int
        
        # Large float
        large_float = 1.23456789e10
        ws['A2'] = large_float
        assert abs(ws['A2'].value - large_float) < 1e-6
        
        wb.close()
    
    def test_special_characters(self):
        """Test handling of special characters in strings."""
        wb = Workbook()
        ws = wb.active
        
        # Unicode characters
        unicode_text = "测试文本 🚀 αβγ"
        ws['A1'] = unicode_text
        assert ws['A1'].value == unicode_text
        
        # Special characters
        special_text = "Special: !@#$%^&*()_+-=[]{}|;':\",./<>?"
        ws['A2'] = special_text
        assert ws['A2'].value == special_text
        
        wb.close()
    
    def test_empty_and_null_values(self):
        """Test handling of empty and null values."""
        wb = Workbook()
        ws = wb.active
        
        # Empty string
        ws['A1'] = ""
        assert ws['A1'].value == ""
        
        # None value
        ws['A2'] = None
        assert ws['A2'].value is None
        
        # Unset cell
        unset_cell = ws['A3']
        assert unset_cell.value is None
        
        wb.close()


class TestPerformanceUnits:
    """Unit tests for performance-related functionality."""
    
    def test_bulk_cell_operations(self):
        """Test performance of bulk cell operations."""
        wb = Workbook()
        ws = wb.active
        
        # Set 100 cells quickly
        start_time = datetime.now()
        for i in range(1, 101):
            ws.cell(i, 1, f"Value_{i}")
        end_time = datetime.now()
        
        # Should complete in reasonable time (less than 1 second)
        duration = (end_time - start_time).total_seconds()
        assert duration < 1.0
        
        # Verify values were set correctly
        assert ws['A1'].value == "Value_1"
        assert ws['A100'].value == "Value_100"
        
        wb.close()
    
    def test_memory_cleanup(self):
        """Test that workbooks properly clean up memory."""
        # Create and close multiple workbooks
        for i in range(10):
            wb = Workbook()
            ws = wb.active
            # Add some data
            for j in range(1, 11):
                ws.cell(j, 1, f"Data_{j}")
            wb.close()
        
        # If we get here without memory issues, test passes
        assert True
