"""
Range and Style Tests - Simplified
Test Range operations and Style system with minimal code.
"""

import pytest
from aspose.cells import Workbook
from aspose.cells.range import Range
from aspose.cells.style import Style, Font, Fill, Border, Alignment
from aspose.cells.utils.exceptions import InvalidCoordinateError


class TestRange:
    """Test Range functionality."""
    
    def test_range_creation(self):
        """Test range creation and properties."""
        wb = Workbook()
        ws = wb.active
        
        range_obj = Range(ws, "A1:B2")
        assert range_obj.coordinate == "A1:B2"
        assert range_obj.min_row == 1
        assert range_obj.max_row == 2
        assert range_obj.row_count == 2
        assert range_obj.column_count == 2
        assert len(range_obj) == 4
        
        wb.close()
    
    def test_range_invalid(self):
        """Test invalid range handling."""
        wb = Workbook()
        ws = wb.active
        
        with pytest.raises(InvalidCoordinateError):
            Range(ws, "INVALID")
        
        wb.close()
    
    def test_range_values(self):
        """Test range value operations."""
        wb = Workbook()
        ws = wb.active
        
        # Add test data
        ws['A1'] = "Header1"
        ws['B1'] = "Header2" 
        ws['A2'] = 100
        ws['B2'] = 200
        
        range_obj = Range(ws, "A1:B2")
        
        # Test getting values
        values = range_obj.values
        assert values[0][0] == "Header1"
        assert values[1][1] == 200
        
        # Test setting values
        range_obj.values = [["A", "B"], [1, 2]]
        assert ws['A1'].value == "A"
        assert ws['B2'].value == 2
        
        wb.close()
    
    def test_range_iteration(self):
        """Test range iteration."""
        wb = Workbook()
        ws = wb.active
        
        ws['A1'] = 1
        ws['B1'] = 2
        ws['A2'] = 3
        ws['B2'] = 4
        
        range_obj = Range(ws, "A1:B2")
        
        # Test cell iteration
        cells = list(range_obj.cells())
        assert len(cells) == 4
        
        # Test row iteration
        rows = list(range_obj.rows_iter())
        assert len(rows) == 2
        assert len(rows[0]) == 2
        
        wb.close()
    
    def test_range_styling(self):
        """Test range styling operations."""
        wb = Workbook()
        ws = wb.active
        
        range_obj = Range(ws, "A1:B2")
        
        # Test font styling
        font = Font()
        font.bold = True
        font.color = "red"
        range_obj.font = font
        
        # Verify styling applied
        for cell in range_obj.cells():
            assert cell.style.font.bold is True
            assert cell.style.font.color == "red"
        
        wb.close()


class TestStyle:
    """Test Style system."""
    
    def test_font(self):
        """Test Font styling."""
        font = Font()
        assert font.name == "Calibri"
        assert font.size == 11
        assert font.bold is False
        
        font.bold = True
        font.color = "blue"
        
        font_copy = font.copy()
        assert font_copy.bold is True
        assert font_copy.color == "blue"
        assert font_copy is not font
    
    def test_fill(self):
        """Test Fill styling."""
        fill = Fill()
        assert fill.color == "white"
        
        fill.color = "yellow"
        fill_copy = fill.copy()
        assert fill_copy.color == "yellow"
    
    def test_border(self):
        """Test Border styling."""
        border = Border()
        
        # Test lazy creation
        left = border.left
        assert left.style == "none"
        
        # Test setting all borders
        border.set_all_borders("thin", "black")
        assert border.left.style == "thin"
        assert border.right.color == "black"
    
    def test_alignment(self):
        """Test Alignment styling."""
        alignment = Alignment()
        assert alignment.horizontal == "general"
        
        alignment.horizontal = "center"
        alignment.wrap_text = True
        
        copy_alignment = alignment.copy()
        assert copy_alignment.horizontal == "center"
        assert copy_alignment.wrap_text is True
    
    def test_complete_style(self):
        """Test complete Style object."""
        style = Style()
        
        # Test lazy property creation
        font = style.font
        assert isinstance(font, Font)
        
        # Test property setting
        style.font.bold = True
        style.fill.color = "lightblue"
        style.number_format = "0.00"
        
        # Test style copy
        style_copy = style.copy()
        assert style_copy.font.bold is True
        assert style_copy.fill.color == "lightblue"
        assert style_copy.number_format == "0.00"
        assert style_copy is not style


class TestStyleIntegration:
    """Test style integration with cells."""
    
    def test_cell_styling(self):
        """Test applying styles to cells."""
        wb = Workbook()
        ws = wb.active
        
        cell = ws['A1']
        cell.value = "Styled Cell"
        
        # Apply various styles
        cell.style.font.bold = True
        cell.style.font.size = 16
        cell.style.fill.color = "yellow"
        cell.style.border.set_outline("thick", "red")
        cell.style.alignment.horizontal = "center"
        
        # Verify styles
        assert cell.style.font.bold is True
        assert cell.style.font.size == 16
        assert cell.style.fill.color == "yellow"
        assert cell.style.border.left.style == "thick"
        assert cell.style.alignment.horizontal == "center"
        
        wb.close()
    
    def test_style_independence(self):
        """Test that cell styles are independent."""
        wb = Workbook()
        ws = wb.active
        
        cell1 = ws['A1']
        cell2 = ws['B1']
        
        # Style cell1
        cell1.style.font.bold = True
        cell1.style.fill.color = "red"
        
        # cell2 should not be affected
        assert cell2.style.font.bold is False
        assert cell2.style.fill.color == "white"
        
        wb.close()