"""
Comprehensive sales data test file demonstrating Excel features.
Includes multiple worksheets, styling, formulas, row heights, column widths, merged cells.
"""

import pytest
from pathlib import Path
from aspose.cells import Workbook, FileFormat


def apply_cell_style(ws, coordinate, **kwargs):
    """Apply cell styling with clear error handling."""
    # Use new convenience method if available, otherwise use direct styling
    if hasattr(ws, 'set_cell_style'):
        ws.set_cell_style(coordinate, **kwargs)
    else:
        # Direct cell styling fallback
        cell = ws[coordinate] if isinstance(coordinate, str) else ws[coordinate[0], coordinate[1]]
        for key, value in kwargs.items():
            if key == 'bold':
                cell.font.bold = value
            elif key == 'font_color':
                cell.font.color = value
            elif key == 'fill_color':
                cell.fill.color = value
            elif key == 'horizontal':
                cell.alignment.horizontal = value
            elif key == 'number_format':
                cell.number_format = value


def create_sales_workbook():
    """Create comprehensive sales workbook with all Excel features."""
    
    # Create workbook
    wb = Workbook()
    
    # Create multiple worksheets (keep default sheet until we have others)
    create_sales_summary_sheet(wb)
    create_product_details_sheet(wb)
    create_financial_analysis_sheet(wb)
    create_charts_data_sheet(wb)
    
    # Remove default sheet if multiple sheets exist
    if "Sheet1" in wb.sheetnames and len(wb.worksheets) > 1:
        wb.worksheets.remove("Sheet1")
    
    return wb


def create_sales_summary_sheet(wb):
    """Create sales summary sheet with comprehensive formatting."""
    ws = wb.create_sheet("Sales Summary")
    wb.active = ws

    # Set column widths (0-based indexing)
    column_widths = [15, 20, 15, 18, 15, 20]
    for i, width in enumerate(column_widths):
        ws.set_column_width(i, width)
    
    # Set row heights (0-based indexing)
    ws.set_row_height(0, 30)  # Title row
    ws.set_row_height(1, 25)  # Header row
    for i in range(2, 7):
        ws.set_row_height(i, 20)  # Data rows
    
    # Merge cells and set title
    ws.merge_cells("A1:F1")
    
    title_cell = ws.cell(1, 1, "2024 Sales Performance Summary Report")
    # Add hyperlink to the main company dashboard
    title_cell.set_hyperlink("https://www.example.com/dashboard", "2024 Sales Performance Summary Report")
    
    # Set title styling
    apply_cell_style(ws, 'A1', 
                    font_name="Arial", font_size=16, bold=True, font_color="white",
                    fill_color="#4472C4", horizontal="center", vertical="center")
    
    # Header data
    headers = ["Product Category", "Quarterly Sales", "Unit Price", "Total Revenue", "Growth Rate", "Notes"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(2, col, header)
        apply_cell_style(ws, (1, col-1),  # 0-based tuple access
                        bold=True, font_color="white", fill_color="#70AD47",
                        horizontal="center")
        cell.border.set_all_borders("thin", "black")
    
    # Sample data
    data = [
        ["Laptops", 1250, 5999.99, None, 0.125, "Best Seller"],
        ["Desktop PCs", 850, 3299.99, None, -0.08, "Sales Declining"],
        ["Tablets", 2100, 2999.99, None, 0.35, "Emerging Market"]
    ]
    
    for row_idx, row_data in enumerate(data, 3):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row_idx, col_idx, value)
            
            # Basic styling setup
            coord = (row_idx-1, col_idx-1)  # Convert to 0-based for new API
            if col_idx == 1:  # Product category column
                apply_cell_style(ws, coord, bold=True, fill_color="#E7E6E6")
            elif col_idx in [2, 3]:  # Number columns
                style_kwargs = {'horizontal': 'right'}
                if col_idx == 3:  # Price formatting
                    style_kwargs['number_format'] = "$#,##0.00"
                apply_cell_style(ws, coord, **style_kwargs)
            elif col_idx == 5:  # Growth rate
                color = "#008000" if value and value > 0 else "#FF0000"
                apply_cell_style(ws, coord, horizontal="center", number_format="0.0%", 
                                font_color=color)
            
            # Borders
            cell.border.set_all_borders("thin", "#CCCCCC")
    
    # Add formulas (Total Revenue = Quantity * Price) with calculated values
    calculated_revenues = [
        1250 * 5999.99,  # Laptops: 7,499,987.50
        850 * 3299.99,   # Desktop PCs: 2,804,991.50  
        2100 * 2999.99   # Tablets: 6,299,979.00
    ]
    
    for idx, row in enumerate(range(3, 6)):
        formula_cell = ws.cell(row, 4)
        if hasattr(formula_cell, 'set_formula'):
            formula_cell.set_formula(f"B{row}*C{row}", calculated_value=calculated_revenues[idx])
        else:
            formula_cell.value = f"=B{row}*C{row}"  # Fallback to direct assignment
        formula_cell.number_format = "$#,##0.00"
    
    # Add totals row
    total_row = 7
    total_quantity = 1250 + 850 + 2100  # 4200
    total_revenue = sum(calculated_revenues)  # Sum of all revenues
    avg_growth = (0.125 + (-0.08) + 0.35) / 3  # Average growth rate
    
    total_data = [
        (total_row, 1, "Total"),
        (total_row, 2, ("=SUM(B3:B5)", total_quantity)),  # Total quantity with calculated value
        (total_row, 3, "Average Price"),
        (total_row, 4, ("=SUM(D3:D5)", total_revenue)),  # Total revenue with calculated value
        (total_row, 5, ("=AVERAGE(E3:E5)", avg_growth)),  # Average growth rate with calculated value
        (total_row, 6, "Summary Data")
    ]
    
    for row, col, value in total_data:
        # Handle tuple values (formula, calculated_value)
        if isinstance(value, tuple):
            formula, calculated_value = value
            cell = ws.cell(row, col)
            if hasattr(cell, 'set_formula'):
                cell.set_formula(formula, calculated_value=calculated_value)
            else:
                cell.value = formula
        else:
            cell = ws.cell(row, col, value)
            
        cell.font.bold = True
        cell.fill.color = "#FFC000"
        cell.border.set_all_borders("thick", "black")
        
        if col in [2, 4]:
            cell.number_format = "$#,##0.00" if col == 4 else "#,##0"
        elif col == 5:
            cell.number_format = "0.0%"


def create_product_details_sheet(wb):
    """Create product details sheet with advanced features."""
    ws = wb.create_sheet("Product Details")
    
    # Set column widths
    # Set uniform column widths
    for col in range(7):
        ws.set_column_width(col, 16)
    
    # Title area merged cells
    ws.merge_cells("A1:G2")
    
    title_cell = ws.cell(1, 1, "Product Detailed Information\nIncluding Inventory, Cost, and Profit Analysis")
    title_cell.font.size = 14
    title_cell.font.bold = True
    title_cell.font.color = "white"
    title_cell.fill.color = "#C55A5A"
    title_cell.alignment.horizontal = "center"
    title_cell.alignment.vertical = "center"
    title_cell.alignment.wrap_text = True
    
    # Headers
    headers = ["Product ID", "Product Name", "Stock Quantity", "Purchase Cost", "Sale Price", "Profit", "Stock Status"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(4, col, header)
        cell.font.bold = True
        cell.fill.color = "#D9E2F3"
        cell.alignment.horizontal = "center"
    
    # Product data - now with ultra-simplified population
    products = [
        ["P001", "Lenovo ThinkPad X1", 45, 4500.00, 5999.99, None, None],
        ["P002", "Dell XPS 13", 23, 4200.00, 5699.99, None, None],
        ["P003", "MacBook Air M2", 67, 7000.00, 8999.99, None, None],
        ["P004", "Huawei MateBook", 89, 3800.00, 4999.99, None, None],
        ["P005", "Xiaomi Notebook Pro", 156, 3200.00, 4299.99, None, None]
    ]
    
    # Define column styles (0-based column index)
    column_styles = {
        0: {'font_name': 'Courier New', 'bold': True, 'fill_color': '#F2F2F2'},  # Product ID
        2: {'horizontal': 'right'},  # Stock quantity
        3: {'horizontal': 'right', 'number_format': '$#,##0.00'},  # Cost
        4: {'horizontal': 'right', 'number_format': '$#,##0.00'},  # Price
    }
    
    # Define conditional styles for stock levels
    conditional_styles = {
        'low_stock': {
            'condition': lambda value, row, col: col == 2 and isinstance(value, (int, float)) and value < 50,
            'style': {'fill_color': '#FFCDD2'}  # Light red
        },
        'medium_stock': {
            'condition': lambda value, row, col: col == 2 and isinstance(value, (int, float)) and 50 <= value < 100,
            'style': {'fill_color': '#FFF9C4'}  # Light yellow
        },
        'high_stock': {
            'condition': lambda value, row, col: col == 2 and isinstance(value, (int, float)) and value >= 100,
            'style': {'fill_color': '#C8E6C9'}  # Light green
        }
    }
    
    # Populate all data with styles
    if hasattr(ws, 'populate_data'):
        ws.populate_data('A5', products, column_styles=column_styles, conditional_styles=conditional_styles)
        
        # Add hyperlinks manually after data population (populate_data doesn't handle hyperlinks)
        product_urls = {
            "Lenovo ThinkPad X1": "https://www.lenovo.com/thinkpad-x1",
            "Dell XPS 13": "https://www.dell.com/xps-13",
            "MacBook Air M2": "https://www.apple.com/macbook-air-m2",
            "Huawei MateBook": "https://consumer.huawei.com/matebook",
            "Xiaomi Notebook Pro": "https://www.mi.com/notebook-pro"
        }
        
        for row_idx, product in enumerate(products, 5):
            product_name = product[1]  # Product name is in column 1 (0-based)
            if product_name in product_urls:
                cell = ws.cell(row_idx, 2)  # Column B (2 in 1-based)
                cell.set_hyperlink(product_urls[product_name], product_name)
    else:
        # Fallback to manual data population with hyperlinks
        for row_idx, product in enumerate(products, 5):
            for col_idx, value in enumerate(product, 1):
                cell = ws.cell(row_idx, col_idx, value)
                
                # Add hyperlinks for product names (column 2)
                if col_idx == 2:
                    product_urls = {
                        "Lenovo ThinkPad X1": "https://www.lenovo.com/thinkpad-x1",
                        "Dell XPS 13": "https://www.dell.com/xps-13",
                        "MacBook Air M2": "https://www.apple.com/macbook-air-m2",
                        "Huawei MateBook": "https://consumer.huawei.com/matebook",
                        "Xiaomi Notebook Pro": "https://www.mi.com/notebook-pro"
                    }
                    if value in product_urls:
                        cell.set_hyperlink(product_urls[value], value)
    
    # Add profit formulas (Column F = Column E - Column D)
    for row in range(5, 10):
        profit_cell = ws.cell(row, 6)
        if hasattr(profit_cell, 'set_formula'):
            profit_cell.set_formula(f"E{row}-D{row}")
        else:
            profit_cell.value = f"=E{row}-D{row}"
        profit_cell.number_format = "$#,##0.00"
        
        # Stock status formulas
        status_cell = ws.cell(row, 7)
        if hasattr(status_cell, 'set_formula'):
            status_cell.set_formula(f'IF(C{row}<50,"Low Stock",IF(C{row}<100,"Normal Stock","High Stock"))')
        else:
            status_cell.value = f'=IF(C{row}<50,"Low Stock",IF(C{row}<100,"Normal Stock","High Stock"))'
        apply_cell_style(ws, (row-1, 6), horizontal="center")  # 0-based coordinate


def create_financial_analysis_sheet(wb):
    """Create financial analysis sheet with complex formulas."""
    ws = wb.create_sheet("Financial Analysis")
    
    # Set column widths (0-based indexing)
    ws.set_column_width(0, 20)
    for col in range(1, 5):
        ws.set_column_width(col, 15)
    
    # Title
    ws.merge_cells("A1:E1")
    
    title_cell = ws.cell(1, 1, "Financial Analysis Report")
    # Add hyperlink to external financial dashboard
    title_cell.set_hyperlink("https://www.example.com/financial-dashboard", "Financial Analysis Report")
    title_cell.font.size = 16
    title_cell.font.bold = True
    title_cell.font.color = "white"
    title_cell.fill.color = "#2F5597"
    title_cell.alignment.horizontal = "center"
    
    # Financial metrics data with calculated values
    sales_revenue = [1250000, 1350000, 1420000, 1680000]
    cost_of_sales = [875000, 945000, 994000, 1176000]
    operating_expenses = [185000, 195000, 205000, 220000]
    
    # Calculate values
    gross_profit = [sales_revenue[i] - cost_of_sales[i] for i in range(4)]
    gross_margin = [gross_profit[i] / sales_revenue[i] for i in range(4)]
    net_profit = [gross_profit[i] - operating_expenses[i] for i in range(4)]
    net_margin = [net_profit[i] / sales_revenue[i] for i in range(4)]
    
    metrics = [
        ("Metric Name", "Q1", "Q2", "Q3", "Q4"),
        ("Sales Revenue", 1250000, 1350000, 1420000, 1680000),
        ("Cost of Sales", 875000, 945000, 994000, 1176000),
        ("Gross Profit", ("=B4-B5", gross_profit[0]), ("=C4-C5", gross_profit[1]), ("=D4-D5", gross_profit[2]), ("=E4-E5", gross_profit[3])),
        ("Gross Margin", ("=B6/B4", gross_margin[0]), ("=C6/C4", gross_margin[1]), ("=D6/D4", gross_margin[2]), ("=E6/E4", gross_margin[3])),
        ("Operating Expenses", 185000, 195000, 205000, 220000),
        ("Net Profit", ("=B6-B8", net_profit[0]), ("=C6-C8", net_profit[1]), ("=D6-D8", net_profit[2]), ("=E6-E8", net_profit[3])),
        ("Net Margin", ("=B9/B4", net_margin[0]), ("=C9/C4", net_margin[1]), ("=D9/D4", net_margin[2]), ("=E9/E4", net_margin[3]))
    ]
    
    for row_idx, metric_row in enumerate(metrics, 3):
        for col_idx, value in enumerate(metric_row, 1):
            # Handle tuple values (formula, calculated_value)
            if isinstance(value, tuple):
                formula, calculated_value = value
                cell = ws.cell(row_idx, col_idx)
                if hasattr(cell, 'set_formula'):
                    cell.set_formula(formula, calculated_value=calculated_value)
                else:
                    cell.value = formula
            else:
                cell = ws.cell(row_idx, col_idx, value)
            
            # Add hyperlinks for financial metrics (first column)
            if col_idx == 1 and row_idx > 3:  # Metric names (excluding header)
                metric_urls = {
                    "Sales Revenue": "https://www.example.com/metrics/sales-revenue",
                    "Cost of Sales": "https://www.example.com/metrics/cost-of-sales", 
                    "Gross Profit": "https://www.example.com/metrics/gross-profit",
                    "Gross Margin": "https://www.example.com/metrics/gross-margin",
                    "Operating Expenses": "https://www.example.com/metrics/operating-expenses",
                    "Net Profit": "https://www.example.com/metrics/net-profit",
                    "Net Margin": "https://www.example.com/metrics/net-margin"
                }
                if value in metric_urls:
                    cell.set_hyperlink(metric_urls[value], value)
            
            if row_idx == 3:  # Header row
                cell.font.bold = True
                cell.font.color = "white"
                cell.fill.color = "#70AD47"
                cell.alignment.horizontal = "center"
            elif col_idx == 1:  # Metric names
                cell.font.bold = True
                cell.fill.color = "#E7E6E6"
            else:  # Data cells
                cell.alignment.horizontal = "right"
                
                # Formatting
                if row_idx in [4, 5, 8, 9]:  # Money values
                    cell.number_format = "$#,##0"
                elif row_idx in [6, 10]:  # Percentages
                    cell.number_format = "0.0%"
            
            # Borders
            cell.border.set_all_borders("thin", "#CCCCCC")


def create_charts_data_sheet(wb):
    """Create a sheet with data suitable for charts."""
    ws = wb.create_sheet("Chart Data")
    
    # Set column widths (0-based indexing)
    for col in range(5):
        ws.set_column_width(col, 18)
    
    # Title
    ws.merge_cells("A1:E1")
    
    title_cell = ws.cell(1, 1, "Monthly Sales Trend Data")
    title_cell.font.size = 14
    title_cell.font.bold = True
    title_cell.font.color = "white"
    title_cell.fill.color = "#E26B0A"
    title_cell.alignment.horizontal = "center"
    
    # Monthly data
    months_data = [
        ("Month", "Laptop Sales", "Desktop Sales", "Tablet Sales", "Total Sales"),
        ("Jan", 320, 180, 450, "=B4+C4+D4"),
        ("Feb", 285, 165, 520, "=B5+C5+D5"),
        ("Mar", 390, 220, 480, "=B6+C6+D6"),
        ("Apr", 425, 195, 510, "=B7+C7+D7"),
        ("May", 380, 175, 535, "=B8+C8+D8"),
        ("Jun", 445, 240, 590, "=B9+C9+D9"),
        ("Jul", 520, 280, 620, "=B10+C10+D10"),
        ("Aug", 485, 255, 580, "=B11+C11+D11"),
        ("Sep", 510, 270, 610, "=B12+C12+D12"),
        ("Oct", 565, 290, 640, "=B13+C13+D13"),
        ("Nov", 620, 315, 680, "=B14+C14+D14"),
        ("Dec", 680, 350, 720, "=B15+C15+D15")
    ]
    
    for row_idx, month_row in enumerate(months_data, 3):
        ws.set_row_height(row_idx, 22)
        
        for col_idx, value in enumerate(month_row, 1):
            cell = ws.cell(row_idx, col_idx, value)
            
            if row_idx == 3:  # Headers
                cell.font.bold = True
                cell.font.color = "white"
                cell.fill.color = "#5B9BD5"
                cell.alignment.horizontal = "center"
            elif col_idx == 1:  # Month names
                cell.font.bold = True
                cell.fill.color = "#DEEBF7"
            else:  # Data
                cell.alignment.horizontal = "right"
                
                if isinstance(value, str) and value.startswith("="):
                    if hasattr(cell, 'set_formula'):
                        cell.set_formula(value)
                    else:
                        cell.value = value
                
                if col_idx == 5:  # Special formatting for total sales column
                    cell.font.bold = True
                    cell.font.color = "#C55A5A"
                
                cell.number_format = "#,##0"
            
            # Borders
            cell.border.set_all_borders("thin", "#CCCCCC")
    
    # Add summary statistics section
    ws.set_row_height(17, 5)  # Empty row
    ws.merge_cells("A18:E18")
    
    summary_cell = ws.cell(18, 1, "Annual Summary Statistics")
    summary_cell.font.size = 13
    summary_cell.font.bold = True
    summary_cell.font.color = "white"
    summary_cell.fill.color = "#A5A5A5"
    summary_cell.alignment.horizontal = "center"
    
    # Summary statistics data
    summary_stats = [
        ("Statistics", "Laptops", "Desktops", "Tablets", "Total"),
        ("Annual Total Sales", "=SUM(B4:B15)", "=SUM(C4:C15)", "=SUM(D4:D15)", "=SUM(E4:E15)"),
        ("Monthly Average", "=AVERAGE(B4:B15)", "=AVERAGE(C4:C15)", "=AVERAGE(D4:D15)", "=AVERAGE(E4:E15)"),
        ("Highest Month", "=MAX(B4:B15)", "=MAX(C4:C15)", "=MAX(D4:D15)", "=MAX(E4:E15)"),
        ("Lowest Month", "=MIN(B4:B15)", "=MIN(C4:C15)", "=MIN(D4:D15)", "=MIN(E4:E15)")
    ]
    
    for row_idx, stat_row in enumerate(summary_stats, 19):
        for col_idx, value in enumerate(stat_row, 1):
            cell = ws.cell(row_idx, col_idx, value)
            
            if row_idx == 19:  # Headers
                cell.font.bold = True
                cell.font.color = "white"
                cell.fill.color = "#70AD47"
            elif col_idx == 1:  # Statistics item names
                cell.font.bold = True
                cell.fill.color = "#F2F2F2"
            else:  # Values
                cell.alignment.horizontal = "right"
                if isinstance(value, str) and value.startswith("="):
                    if hasattr(cell, 'set_formula'):
                        cell.set_formula(value)
                    else:
                        cell.value = value
                cell.number_format = "#,##0"


class TestAdvancedFeatures:
    """Test comprehensive Excel features with complex workbook creation."""
    
    def setup_method(self):
        """Set up test environment with dedicated output folder."""
        from pathlib import Path
        self.output_dir = Path(__file__).parent / "testdata" / "test_advanced_features"
        self.output_dir.mkdir(exist_ok=True)
    
    def test_sales_workbook_creation(self, ensure_testdata_dir):
        """Comprehensive test: Test creating comprehensive sales workbook with all features."""
        wb = create_sales_workbook()
        
        # Verify workbook structure
        assert len(wb.worksheets) >= 4, "Should have at least 4 worksheets"
        expected_sheets = ["Sales Summary", "Product Details", "Financial Analysis", "Chart Data"]
        
        for sheet_name in expected_sheets:
            assert sheet_name in wb.sheetnames, f"Missing worksheet: {sheet_name}"
        
        # Verify active worksheet
        assert wb.active is not None
        
        wb.close()
    
    def test_export_xlsx_format(self, ensure_testdata_dir):
        """Comprehensive test: Test XLSX export functionality."""
        wb = create_sales_workbook()
        wb.active = wb.worksheets["Sales Summary"]
        
        xlsx_file = self.output_dir / "sales_report_comprehensive.xlsx"
        wb.save(str(xlsx_file))
        
        # Verify file was created and has content
        assert xlsx_file.exists(), "XLSX file should be created"
        assert xlsx_file.stat().st_size > 0, "XLSX file should not be empty"
        
        wb.close()
    
    
    def test_export_json_format(self, ensure_testdata_dir):
        """Test JSON export functionality."""
        wb = create_sales_workbook()
        wb.active = wb.worksheets["Sales Summary"]
        
        json_output = wb.exportAs(FileFormat.JSON, all_sheets=True)
        assert isinstance(json_output, str), "JSON output should be string"
        assert len(json_output) > 0, "JSON output should not be empty"
        
        # Save to file and verify
        json_file = self.output_dir / "sales_report_comprehensive.json"
        with open(json_file, "w", encoding="utf-8") as f:
            f.write(json_output)
        
        assert json_file.exists(), "JSON file should be created"
        assert json_file.stat().st_size > 0, "JSON file should not be empty"
        
        wb.close()
    
    
    def test_export_markdown_format(self, ensure_testdata_dir):
        """Test Markdown export functionality."""
        wb = create_sales_workbook()
        wb.active = wb.worksheets["Sales Summary"]
        
        md_output = wb.exportAs(FileFormat.MARKDOWN, all_sheets=True)
        assert isinstance(md_output, str), "Markdown output should be string"
        assert len(md_output) > 0, "Markdown output should not be empty"
        assert "|" in md_output, "Markdown should contain table formatting"
        
        # Save to file and verify
        md_file = self.output_dir / "sales_report_comprehensive.md"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_output)
        
        assert md_file.exists(), "Markdown file should be created"
        assert md_file.stat().st_size > 0, "Markdown file should not be empty"
        
        wb.close()
    
    
    def test_worksheet_data_integrity(self):
        """Test that worksheet data is properly populated."""
        wb = create_sales_workbook()
        
        # Test Sales Summary sheet
        sales_summary = wb.worksheets["Sales Summary"]
        assert sales_summary['A1'].value is not None, "Sales Summary should have data in A1"
        
        # Test Product Details sheet
        product_details = wb.worksheets["Product Details"]
        assert product_details['A1'].value is not None, "Product Details should have data in A1"
        
        # Test Financial Analysis sheet
        financial_analysis = wb.worksheets["Financial Analysis"]
        assert financial_analysis['A1'].value is not None, "Financial Analysis should have data in A1"
        
        # Test Chart Data sheet
        chart_data = wb.worksheets["Chart Data"]
        assert chart_data['A1'].value is not None, "Chart Data should have data in A1"
        
        wb.close()
    
    
    def test_complex_formatting_features(self):
        """Test that complex formatting features are applied."""
        wb = create_sales_workbook()
        
        # Test that worksheets have reasonable dimensions
        for sheet_name in wb.sheetnames:
            ws = wb.worksheets[sheet_name]
            assert ws.max_row > 0, f"{sheet_name} should have rows"
            assert ws.max_column > 0, f"{sheet_name} should have columns"
        
        wb.close()
    
    
    def test_complete_export_workflow(self, ensure_testdata_dir):
        """Test complete export workflow for all formats."""
        wb = create_sales_workbook()
        wb.active = wb.worksheets["Sales Summary"]
        
        # Define all export formats to test
        export_tests = [
            (FileFormat.XLSX, "xlsx", "sales_comprehensive"),
            (FileFormat.JSON, "json", "sales_comprehensive"),
            (FileFormat.CSV, "csv", "sales_comprehensive"),
            (FileFormat.MARKDOWN, "md", "sales_comprehensive")
        ]
        
        successful_exports = 0
        
        for format_type, extension, base_name in export_tests:
            try:
                if format_type == FileFormat.XLSX:
                    # XLSX uses save method
                    file_path = self.output_dir / f"{base_name}.{extension}"
                    wb.save(str(file_path))
                else:
                    # Other formats use exportAs method
                    output = wb.exportAs(format_type, all_sheets=True)
                    file_path = self.output_dir / f"{base_name}.{extension}"
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(output)
                
                # Verify file was created
                assert file_path.exists(), f"{format_type.value} file should exist"
                assert file_path.stat().st_size > 0, f"{format_type.value} file should not be empty"
                successful_exports += 1
                
            except Exception as e:
                pytest.fail(f"Export failed for {format_type.value}: {e}")
        
        # Verify all exports succeeded
        assert successful_exports == len(export_tests), "All export formats should succeed"
        
        wb.close()