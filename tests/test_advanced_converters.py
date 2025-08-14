"""
Advanced Converters Tests - Simplified
Test MarkItDown, and other converters with minimal code.
"""

import pytest
import json
from aspose.cells import Workbook
from aspose.cells.converters.markdown_converter import MarkdownConverter
from aspose.cells.converters.json_converter import JsonConverter


class TestBasicConverters:
    """Test basic Markdown and JSON converters."""
    
    def test_markdown_converter(self):
        """Test Markdown converter."""
        converter = MarkdownConverter()
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Header"
        ws['A2'] = "Data"
        
        result = converter.convert_workbook(wb)
        
        assert "Header" in result
        assert "Data" in result
        assert "|" in result
        wb.close()
    
    def test_json_converter(self):
        """Test JSON converter."""
        converter = JsonConverter()
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Header"
        ws['A2'] = "Data"
        
        result = converter.convert_workbook(wb)
        data = json.loads(result)
        
        assert isinstance(data, list)
        assert len(data) >= 1
        wb.close()
    
    def test_all_converters_integration(self):
        """Test all converters with same data."""
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Test"
        ws['A2'] = 123
        
        converters = [
            MarkdownConverter(),
            JsonConverter()
        ]
        
        for converter in converters:
            result = converter.convert_workbook(wb)
            assert isinstance(result, str)
            assert len(result) > 0
            assert "Test" in result or "test" in result.lower()
        
        wb.close()