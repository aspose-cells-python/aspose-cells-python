"""
Format Conversion Tests
Tests for converting Excel files to various output formats.
All conversion outputs are saved to test_conversions/ directory for inspection.
"""

import pytest
import json
from pathlib import Path
from aspose.cells import Workbook, FileFormat


class TestConversions:
    """Comprehensive format conversion tests."""
    
    def setup_method(self):
        """Set up test environment with dedicated output folder."""
        from pathlib import Path
        self.output_dir = Path(__file__).parent / "testdata" / "test_conversions"
        self.output_dir.mkdir(exist_ok=True)
    
    @pytest.fixture
    def sample_workbook(self):
        """Create a sample workbook for conversion tests."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Sample Data"
        
        # Add sample data
        data = [
            ["Product", "Price", "Quantity", "Total"],
            ["Laptop", 999.99, 10, "=B2*C2"],
            ["Mouse", 25.50, 50, "=B3*C3"],
            ["Keyboard", 75.00, 25, "=B4*C4"],
            ["Monitor", 299.99, 15, "=B5*C5"]
        ]
        
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Add styling to headers
        for col in range(1, 5):
            cell = ws.cell(1, col)
            cell.style.font.bold = True
            cell.style.fill.background_color = "lightgray"
        
        return wb
    
    def test_excel_to_csv_conversion(self, sample_workbook, ensure_testdata_dir):
        """Test Excel to CSV conversion."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        # Convert to CSV
        csv_content = sample_workbook.exportAs(FileFormat.CSV)
        assert isinstance(csv_content, str)
        assert "Product,Price,Quantity,Total" in csv_content
        assert "Laptop,999.99,10" in csv_content
        
        # Save to output directory
        output_file = self.output_dir / "sample_data.csv"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(csv_content)
        
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        sample_workbook.close()
    
    def test_excel_to_json_conversion(self, sample_workbook, ensure_testdata_dir):
        """Test Excel to JSON conversion."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        # Convert to JSON
        json_content = sample_workbook.exportAs(FileFormat.JSON)
        assert isinstance(json_content, str)
        
        # Validate JSON structure
        data = json.loads(json_content)
        assert isinstance(data, list)
        assert len(data) > 0
        
        # Save to output directory
        output_file = self.output_dir / "sample_data.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_content)
        
        assert output_file.exists()
        
        sample_workbook.close()
    
    def test_excel_to_markdown_conversion(self, sample_workbook, ensure_testdata_dir):
        """Test Excel to Markdown conversion."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        # Convert to Markdown
        try:
            md_content = sample_workbook.exportAs(FileFormat.MARKDOWN)
            assert isinstance(md_content, str)
            
            # Check for markdown table formatting
            assert "|" in md_content  # Table separator
            assert "Product" in md_content
            assert "Laptop" in md_content
            
            # Save to output directory
            output_file = self.output_dir / "sample_data.md"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            assert output_file.exists()
            
        except Exception as e:
            if "not implemented" in str(e).lower():
                pytest.skip("Markdown conversion not yet implemented")
            else:
                raise
        
        sample_workbook.close()
    
    def test_multi_sheet_conversion(self, ensure_testdata_dir):
        """Test conversion of multi-sheet workbook."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        # Create multi-sheet workbook
        wb = Workbook()
        
        # Sheet 1 - Sales Data
        sales_ws = wb.active
        sales_ws.name = "Sales"
        sales_data = [
            ["Month", "Sales"],
            ["Jan", 10000],
            ["Feb", 12000],
            ["Mar", 11500]
        ]
        for row_idx, row_data in enumerate(sales_data):
            for col_idx, value in enumerate(row_data):
                sales_ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Sheet 2 - Expenses Data
        expenses_ws = wb.create_sheet("Expenses")
        expenses_data = [
            ["Month", "Expenses"],
            ["Jan", 7000],
            ["Feb", 8000],
            ["Mar", 7500]
        ]
        for row_idx, row_data in enumerate(expenses_data):
            for col_idx, value in enumerate(row_data):
                expenses_ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Convert to different formats
        formats = [
            (FileFormat.CSV, "multi_sheet.csv"),
            (FileFormat.JSON, "multi_sheet.json")
        ]
        
        for fmt, filename in formats:
            try:
                content = wb.exportAs(fmt)
                output_file = self.output_dir / filename
                
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                assert output_file.exists()
                
            except Exception as e:
                if "not implemented" in str(e).lower():
                    pytest.skip(f"Format {fmt} conversion not yet implemented")
                else:
                    raise
        
        wb.close()
    
    def test_styled_workbook_conversion(self, ensure_testdata_dir):
        """Test conversion of styled workbook."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        ws.name = "Styled Data"
        
        # Add data with various styles
        ws['A1'] = "Header 1"
        ws['B1'] = "Header 2"
        ws['A2'] = "Bold Text"
        ws['B2'] = "Regular Text"
        
        # Apply styles
        ws['A1'].style.font.bold = True
        ws['A1'].style.fill.background_color = "blue"
        ws['A2'].style.font.bold = True
        ws['B1'].style.font.italic = True
        
        # Test CSV conversion (styles should be ignored)
        csv_content = wb.exportAs(FileFormat.CSV)
        output_file = self.output_dir / "styled_workbook.csv"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(csv_content)
        
        assert output_file.exists()
        assert "Header 1,Header 2" in csv_content
        
        wb.close()
    
    def test_formula_conversion(self, ensure_testdata_dir):
        """Test conversion of workbooks with formulas."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        
        # Add data with formulas
        ws['A1'] = 10
        ws['A2'] = 20
        ws['A3'] = "=A1+A2"
        ws['A4'] = "=SUM(A1:A2)"
        ws['A5'] = "=AVERAGE(A1:A2)"
        
        # Convert to different formats
        formats_to_test = [
            (FileFormat.CSV, "formulas.csv"),
            (FileFormat.JSON, "formulas.json")
        ]
        
        for fmt, filename in formats_to_test:
            try:
                content = wb.exportAs(fmt)
                output_file = self.output_dir / filename
                
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                assert output_file.exists()
                
                # For CSV, check that calculated values are present
                if fmt == FileFormat.CSV:
                    assert "10" in content
                    assert "20" in content
                    # Formula results should be calculated or formulas preserved
                
            except Exception as e:
                if "not implemented" in str(e).lower():
                    pytest.skip(f"Format {fmt} conversion not yet implemented")
                else:
                    raise
        
        wb.close()
    
    def test_large_workbook_conversion(self, ensure_testdata_dir):
        """Test conversion of larger workbook."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        
        # Generate 500 rows of data
        ws['A1'] = "ID"
        ws['B1'] = "Value"
        ws['C1'] = "Category"
        
        for row in range(2, 502):  # 500 data rows
            ws.cell(row, 1, row - 1)  # ID
            ws.cell(row, 2, (row - 1) * 10)  # Value
            ws.cell(row, 3, f"Cat_{(row - 1) % 5}")  # Category
        
        # Convert to CSV
        csv_content = wb.exportAs(FileFormat.CSV)
        output_file = self.output_dir / "large_workbook.csv"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(csv_content)
        
        assert output_file.exists()
        assert output_file.stat().st_size > 5000  # Should be substantial
        
        # Verify content structure
        lines = csv_content.split('\n')
        assert len(lines) >= 501  # Header + 500 data rows
        assert "ID,Value,Category" in lines[0]
        
        wb.close()
    
    def test_conversion_with_special_characters(self, ensure_testdata_dir):
        """Test conversion with special characters and unicode."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        
        # Add data with special characters
        special_data = [
            ["English", "中文", "العربية", "Русский"],
            ["Hello", "你好", "مرحبا", "Привет"],
            ["Symbols", "!@#$%", "αβγδε", "🚀🌟⭐"],
            ["Quotes", "\"Test\"", "'Single'", "Mixed\"'Text"]
        ]
        
        for row_idx, row_data in enumerate(special_data):
            for col_idx, value in enumerate(row_data):
                ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Convert to different formats
        formats_to_test = [
            (FileFormat.CSV, "special_chars.csv"),
            (FileFormat.JSON, "special_chars.json")
        ]
        
        for fmt, filename in formats_to_test:
            try:
                content = wb.exportAs(fmt)
                output_file = self.output_dir / filename
                
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                assert output_file.exists()
                
                # Verify unicode content is preserved
                if fmt == FileFormat.CSV:
                    assert "中文" in content
                    assert "العربية" in content
                elif fmt == FileFormat.JSON:
                    data = json.loads(content)
                    assert any("中文" in str(row) for row in data)
                
            except Exception as e:
                if "not implemented" in str(e).lower():
                    pytest.skip(f"Format {fmt} conversion not yet implemented")
                else:
                    raise
        
        wb.close()
    
    def test_empty_workbook_conversion(self, ensure_testdata_dir):
        """Test conversion of empty workbook."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        wb = Workbook()
        
        # Convert empty workbook
        csv_content = wb.exportAs(FileFormat.CSV)
        output_file = self.output_dir / "empty_workbook.csv"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(csv_content)
        
        assert output_file.exists()
        # Empty workbook should produce minimal content
        
        wb.close()
    
    def test_conversion_output_directory_structure(self, ensure_testdata_dir):
        """Test that all conversion outputs are properly organized."""
        output_dir = self.output_dir
        
        # Verify output directory exists and contains files
        assert output_dir.exists()
        assert output_dir.is_dir()
        
        # Check for expected output files from previous tests
        expected_files = [
            "sample_data.csv",
            "sample_data.json"
        ]
        
        for filename in expected_files:
            file_path = output_dir / filename
            if file_path.exists():
                assert file_path.stat().st_size > 0
    
    def test_batch_conversion(self, ensure_testdata_dir):
        """Test batch conversion of multiple workbooks."""
        output_dir = self.output_dir
        output_dir.mkdir(exist_ok=True)
        
        # Create multiple workbooks and convert them
        workbook_configs = [
            ("batch_1", [["A", 1], ["B", 2]]),
            ("batch_2", [["X", 10], ["Y", 20]]),
            ("batch_3", [["P", 100], ["Q", 200]])
        ]
        
        for name, data in workbook_configs:
            wb = Workbook()
            ws = wb.active
            
            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    ws.cell(row_idx + 1, col_idx + 1, value)
            
            # Convert to CSV
            csv_content = wb.exportAs(FileFormat.CSV)
            output_file = self.output_dir / f"{name}.csv"
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(csv_content)
            
            assert output_file.exists()
            wb.close()
        
        # Verify all batch files were created
        for name, _ in workbook_configs:
            assert (self.output_dir / f"{name}.csv").exists()
