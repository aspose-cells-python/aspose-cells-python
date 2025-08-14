"""
Excel Generation Tests
Tests for creating various Excel files with different features and data.
All generated files are saved to testdata/ directory for inspection and further testing.
"""

import pytest
from pathlib import Path
from aspose.cells import Workbook, FileFormat


class TestExcelGeneration:
    """Comprehensive Excel file generation tests."""
    
    def test_basic_workbook_generation(self, ensure_testdata_dir, sample_data):
        """Generate basic Excel workbook with sample data."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Sales Data"
        
        # Add sample data
        for row_idx, row_data in enumerate(sample_data):
            for col_idx, value in enumerate(row_data):
                ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Add formulas for calculations
        for row in range(2, 7):  # Rows 2-6 (data rows)
            # Total = sum of Q1-Q4
            ws.cell(row, 6, f"=SUM(B{row}:E{row})")
            # Average = Total/4
            ws.cell(row, 7, f"=F{row}/4")
            # Growth % = (Q4-Q1)/Q1*100
            ws.cell(row, 8, f"=(E{row}-B{row})/B{row}*100")
        
        # Add totals row
        ws.cell(7, 1, "TOTALS")
        for col in range(2, 9):  # Columns B-H
            col_letter = chr(ord('A') + col - 1)
            ws.cell(7, col, f"=SUM({col_letter}2:{col_letter}6)")
        
        output_file = ensure_testdata_dir / "basic_workbook.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()
        assert output_file.stat().st_size > 0

    def test_financial_report_generation(self, ensure_testdata_dir, financial_data):
        """Generate financial report with multiple worksheets."""
        wb = Workbook()
        
        # Summary sheet
        summary_ws = wb.active
        summary_ws.name = "Summary"
        
        # Company info
        info = financial_data["company_info"]
        summary_ws['A1'] = info["name"]
        summary_ws['A2'] = info["period"]
        summary_ws['A3'] = f"Prepared by: {info['prepared_by']}"
        
        # Revenue details sheet
        revenue_ws = wb.create_sheet("Revenue Details")
        revenue_data = financial_data["revenue_data"]
        
        for row_idx, row_data in enumerate(revenue_data):
            for col_idx, value in enumerate(row_data):
                revenue_ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Add profit and margin calculations
        for row in range(2, 8):  # Data rows
            # Profit = Revenue - Expenses
            revenue_ws.cell(row, 4, f"=B{row}-C{row}")
            # Margin % = Profit/Revenue*100
            revenue_ws.cell(row, 5, f"=D{row}/B{row}*100")
        
        output_file = ensure_testdata_dir / "financial_report.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()

    def test_employee_data_generation(self, ensure_testdata_dir, employee_data):
        """Generate employee data workbook with calculations."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Employee Data"
        
        # Add employee data
        for row_idx, row_data in enumerate(employee_data):
            for col_idx, value in enumerate(row_data):
                ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Calculate total compensation
        for row in range(2, 8):  # Employee data rows
            # Total Comp = Salary + (Salary * Bonus% / 100)
            ws.cell(row, 6, f"=D{row}+(D{row}*E{row}/100)")
        
        output_file = ensure_testdata_dir / "employee_data.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()

    def test_styled_workbook_generation(self, ensure_testdata_dir):
        """Generate workbook with various styling features."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Styled Data"
        
        # Headers with styling
        headers = ["Product", "Price", "Quantity", "Total", "Status"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(1, col, header)
            # Apply basic styling through cell properties
            cell.style.font.bold = True
            cell.style.fill.background_color = "lightblue"
        
        # Sample data with conditional formatting
        data = [
            ["Laptop", 999.99, 50, "=B2*C2", "In Stock"],
            ["Mouse", 25.50, 200, "=B3*C3", "In Stock"],
            ["Keyboard", 75.00, 0, "=B4*C4", "Out of Stock"],
            ["Monitor", 299.99, 25, "=B5*C5", "Low Stock"],
        ]
        
        for row_idx, row_data in enumerate(data, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row_idx, col_idx, value)
                
                # Style status column based on value
                if col_idx == 5:  # Status column
                    if value == "Out of Stock":
                        cell.style.font.color = "red"
                    elif value == "Low Stock":
                        cell.style.font.color = "orange"
                    else:
                        cell.style.font.color = "green"
        
        output_file = ensure_testdata_dir / "styled_workbook.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()

    def test_multi_sheet_workbook_generation(self, ensure_testdata_dir):
        """Generate workbook with multiple related worksheets."""
        wb = Workbook()
        
        # Products sheet
        products_ws = wb.active
        products_ws.name = "Products"
        
        products_data = [
            ["ID", "Name", "Category", "Price"],
            [1, "Laptop Pro", "Electronics", 1299.99],
            [2, "Office Chair", "Furniture", 299.99],
            [3, "Coffee Maker", "Appliances", 89.99],
        ]
        
        for row_idx, row_data in enumerate(products_data):
            for col_idx, value in enumerate(row_data):
                products_ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Orders sheet
        orders_ws = wb.create_sheet("Orders")
        
        orders_data = [
            ["Order ID", "Product ID", "Quantity", "Order Date"],
            [1001, 1, 2, "2024-01-15"],
            [1002, 3, 1, "2024-01-16"],
            [1003, 2, 3, "2024-01-17"],
        ]
        
        for row_idx, row_data in enumerate(orders_data):
            for col_idx, value in enumerate(row_data):
                orders_ws.cell(row_idx + 1, col_idx + 1, value)
        
        # Summary sheet with cross-sheet references
        summary_ws = wb.create_sheet("Summary")
        summary_ws['A1'] = "Total Products"
        summary_ws['B1'] = "=COUNTA(Products!A:A)-1"  # Count excluding header
        summary_ws['A2'] = "Total Orders"
        summary_ws['B2'] = "=COUNTA(Orders!A:A)-1"
        
        output_file = ensure_testdata_dir / "multi_sheet_workbook.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()

    def test_large_dataset_generation(self, ensure_testdata_dir):
        """Generate workbook with larger dataset for performance testing."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Large Dataset"
        
        # Headers
        headers = ["ID", "Name", "Value", "Category", "Date", "Status"]
        for col, header in enumerate(headers, 1):
            ws.cell(1, col, header)
        
        # Generate 1000 rows of data
        import random
        from datetime import datetime, timedelta
        
        categories = ["A", "B", "C", "D"]
        statuses = ["Active", "Inactive", "Pending"]
        base_date = datetime(2024, 1, 1)
        
        for row in range(2, 1002):  # 1000 data rows
            ws.cell(row, 1, row - 1)  # ID
            ws.cell(row, 2, f"Item_{row-1:04d}")  # Name
            ws.cell(row, 3, round(random.uniform(10, 1000), 2))  # Value
            ws.cell(row, 4, random.choice(categories))  # Category
            ws.cell(row, 5, (base_date + timedelta(days=random.randint(0, 365))).strftime("%Y-%m-%d"))  # Date
            ws.cell(row, 6, random.choice(statuses))  # Status
        
        output_file = ensure_testdata_dir / "large_dataset.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()
        # Verify file size is reasonable for 1000+ rows
        assert output_file.stat().st_size > 30000  # At least 30KB

    def test_template_workbook_generation(self, ensure_testdata_dir):
        """Generate template workbook for reuse."""
        wb = Workbook()
        ws = wb.active
        ws.name = "Template"
        
        # Create a template structure
        ws['A1'] = "COMPANY REPORT TEMPLATE"
        ws['A3'] = "Company Name:"
        ws['A4'] = "Report Period:"
        ws['A5'] = "Prepared By:"
        ws['A7'] = "Data Section:"
        
        # Headers for data section
        data_headers = ["Item", "Q1", "Q2", "Q3", "Q4", "Total", "Average"]
        for col, header in enumerate(data_headers, 1):
            ws.cell(8, col, header)
        
        # Add some placeholder formulas
        for row in range(9, 15):  # 6 data rows
            ws.cell(row, 6, f"=SUM(B{row}:E{row})")  # Total
            ws.cell(row, 7, f"=F{row}/4")  # Average
        
        output_file = ensure_testdata_dir / "report_template.xlsx"
        wb.save(str(output_file), FileFormat.XLSX)
        wb.close()
        
        assert output_file.exists()
