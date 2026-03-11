"""
Test case to create all supported chart types from scratch and save to outputfiles/progcharts folder.

This test creates the following chart types:
- Line Chart
- Bar Chart (Column)
- Pie Chart
- Area Chart
- Box and Whisker Chart
- Waterfall Chart
- Combo Chart
- Scatter (XY) Chart
- Stock Chart
- Surface Chart
- Radar Chart
- Treemap Chart
- Sunburst Chart
- Histogram Chart
- Funnel Chart
- Map Chart
"""

import os
import sys
import pytest

# Add parent directory to path to import aspose.cells_foss
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose.cells_foss import Workbook, ChartType


def test_create_line_chart():
    """Create a line chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Month"
    ws.cells["B1"].value = "Sales"
    ws.cells["C1"].value = "Expenses"
    
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    sales = [100, 150, 120, 180, 200, 170]
    expenses = [80, 90, 85, 100, 110, 95]
    
    for i, (month, sale, expense) in enumerate(zip(months, sales, expenses), 2):
        ws.cells[f"A{i}"].value = month
        ws.cells[f"B{i}"].value = sale
        ws.cells[f"C{i}"].value = expense
    
    # Create line chart
    chart = ws.charts.add_line(0, 4, 20, 12)
    chart.title = "Monthly Sales and Expenses"
    chart.category_data = "A2:A7"
    chart.show_legend = True
    chart.legend_position = "right"
    
    # Add series
    chart.n_series.add("B2:B7", category_data="A2:A7", name="Sales")
    chart.n_series.add("C2:C7", category_data="A2:A7", name="Expenses")
    
    # Save
    output_path = "outputfiles/progcharts/line_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_bar_chart():
    """Create a bar (column) chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Product"
    ws.cells["B1"].value = "Q1"
    ws.cells["C1"].value = "Q2"
    ws.cells["D1"].value = "Q3"
    ws.cells["E1"].value = "Q4"
    
    products = ["Product A", "Product B", "Product C", "Product D"]
    data = [
        [120, 150, 180, 200],
        [90, 110, 130, 160],
        [200, 220, 190, 210],
        [150, 170, 160, 180]
    ]
    
    for i, product in enumerate(products, 2):
        ws.cells[f"A{i}"].value = product
        for j, value in enumerate(data[i-2], 2):
            ws.cells[f"{chr(64+j)}{i}"].value = value
    
    # Create bar chart
    chart = ws.charts.add_bar(0, 6, 20, 14)
    chart.title = "Quarterly Sales by Product"
    chart.category_data = "A2:A5"
    chart.bar_direction = "col"
    chart.grouping = "clustered"
    chart.gap_width = 150
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("B2:B5", category_data="A2:A5", name="Q1")
    chart.n_series.add("C2:C5", category_data="A2:A5", name="Q2")
    chart.n_series.add("D2:D5", category_data="A2:A5", name="Q3")
    chart.n_series.add("E2:E5", category_data="A2:A5", name="Q4")
    
    # Save
    output_path = "outputfiles/progcharts/bar_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_pie_chart():
    """Create a pie chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Category"
    ws.cells["B1"].value = "Value"
    
    categories = ["Electronics", "Clothing", "Food", "Books", "Others"]
    values = [35, 25, 20, 10, 10]
    
    for i, (cat, val) in enumerate(zip(categories, values), 2):
        ws.cells[f"A{i}"].value = cat
        ws.cells[f"B{i}"].value = val
    
    # Create pie chart
    chart = ws.charts.add_pie(0, 4, 20, 12)
    chart.title = "Sales Distribution by Category"
    chart.category_data = "A2:A6"
    chart.show_legend = True
    chart.legend_position = "right"
    chart.vary_colors = True
    chart.first_slice_angle = 0
    
    # Add series
    chart.n_series.add("B2:B6", category_data="A2:A6", name="Sales")
    
    # Save
    output_path = "outputfiles/progcharts/pie_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_area_chart():
    """Create an area chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Year"
    ws.cells["B1"].value = "Revenue"
    ws.cells["C1"].value = "Profit"
    
    years = ["2019", "2020", "2021", "2022", "2023"]
    revenue = [1000, 1200, 1500, 1800, 2000]
    profit = [200, 250, 350, 400, 500]
    
    for i, (year, rev, prof) in enumerate(zip(years, revenue, profit), 2):
        ws.cells[f"A{i}"].value = year
        ws.cells[f"B{i}"].value = rev
        ws.cells[f"C{i}"].value = prof
    
    # Create area chart
    chart = ws.charts.add_area(0, 4, 20, 12)
    chart.title = "Revenue and Profit Trend"
    chart.category_data = "A2:A6"
    chart.grouping = "standard"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("B2:B6", category_data="A2:A6", name="Revenue")
    chart.n_series.add("C2:C6", category_data="A2:A6", name="Profit")
    
    # Save
    output_path = "outputfiles/progcharts/area_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_box_whisker_chart():
    """Create a box and whisker chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data - sample statistics for different groups
    ws.cells["A1"].value = "Group"
    ws.cells["B1"].value = "Q1"
    ws.cells["C1"].value = "Q2"
    ws.cells["D1"].value = "Q3"
    ws.cells["E1"].value = "Q4"
    ws.cells["F1"].value = "Q5"
    
    groups = ["Group A", "Group B", "Group C"]
    data = [
        [10, 15, 20, 25, 30],
        [12, 18, 22, 28, 35],
        [8, 14, 19, 24, 32]
    ]
    
    for i, group in enumerate(groups, 2):
        ws.cells[f"A{i}"].value = group
        for j, value in enumerate(data[i-2], 2):
            ws.cells[f"{chr(64+j)}{i}"].value = value
    
    # Create box and whisker chart
    chart = ws.charts.add_box_whisker(0, 4, 20, 12)
    chart.title = "Statistical Distribution by Group"
    chart.category_data = "B1:F1"
    chart.show_legend = True
    chart.quartile_method = "exclusive"
    chart.box_show_mean_line = False
    chart.box_show_mean_marker = True
    chart.box_show_inner_points = False
    chart.box_show_outlier_points = True
    chart.box_gap_width = 1

    # Add series
    chart.n_series.add("B2:F2", category_data="B1:F1", name="Group A")
    chart.n_series.add("B3:F3", category_data="B1:F1", name="Group B")
    chart.n_series.add("B4:F4", category_data="B1:F1", name="Group C")
    
    # Save
    output_path = "outputfiles/progcharts/box_whisker_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")

    # Ensure box-whisker chartEx settings are serialized
    import zipfile
    with zipfile.ZipFile(output_path) as zf:
        chart_xml = zf.read("xl/charts/chartEx1.xml").decode("utf-8")
        assert '<cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1" />' in chart_xml
        assert '<cx:catScaling gapWidth="1" />' in chart_xml


def test_create_waterfall_chart():
    """Create a waterfall chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Item"
    ws.cells["B1"].value = "Value"
    
    items = ["Starting", "Sales", "Costs", "Expenses", "Taxes", "Ending"]
    values = [1000, 500, -200, -150, -100, 1050]
    
    for i, (item, val) in enumerate(zip(items, values), 2):
        ws.cells[f"A{i}"].value = item
        ws.cells[f"B{i}"].value = val
    
    # Create waterfall chart
    chart = ws.charts.add_waterfall(0, 4, 20, 12)
    chart.title = "Cash Flow Waterfall"
    chart.category_data = "A2:A7"
    chart.show_legend = False
    
    # Add series
    chart.n_series.add("B2:B7", category_data="A2:A7", name="Cash Flow")
    
    # Save
    output_path = "outputfiles/progcharts/waterfall_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_scatter_chart():
    """Create a scatter (XY) chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "X"
    ws.cells["B1"].value = "Y1"
    ws.cells["C1"].value = "Y2"
    
    x_values = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    y1_values = [2, 4, 5, 4, 5, 7, 8, 9, 10, 12]
    y2_values = [1, 3, 2, 5, 4, 6, 5, 8, 7, 9]
    
    for i, (x, y1, y2) in enumerate(zip(x_values, y1_values, y2_values), 2):
        ws.cells[f"A{i}"].value = x
        ws.cells[f"B{i}"].value = y1
        ws.cells[f"C{i}"].value = y2
    
    # Create scatter chart
    chart = ws.charts.add_scatter(0, 4, 20, 12)
    chart.title = "Scatter Plot Example"
    chart.scatter_style = "lineMarker"
    chart.show_legend = True
    
    # Add series with x_values
    chart.n_series.add("B2:B11", category_data="A2:A11", name="Series 1", x_values="A2:A11")
    chart.n_series.add("C2:C11", category_data="A2:A11", name="Series 2", x_values="A2:A11")
    
    # Save
    output_path = "outputfiles/progcharts/scatter_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_combo_chart():
    """Create a combo chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Month"
    ws.cells["B1"].value = "Sales"
    ws.cells["C1"].value = "Profit %"
    
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    sales = [100, 150, 120, 180, 200, 170]
    profit_pct = [20, 25, 18, 22, 28, 24]
    
    for i, (month, sale, profit) in enumerate(zip(months, sales, profit_pct), 2):
        ws.cells[f"A{i}"].value = month
        ws.cells[f"B{i}"].value = sale
        ws.cells[f"C{i}"].value = profit
    
    # Create combo chart
    chart = ws.charts.add_combo(0, 4, 20, 12)
    chart.title = "Sales and Profit Margin"
    chart.category_data = "A2:A7"
    chart.show_legend = True
    
    # Add series with different chart types
    chart.n_series.add("B2:B7", category_data="A2:A7", name="Sales", chart_type=ChartType.BAR)
    chart.n_series.add("C2:C7", category_data="A2:A7", name="Profit %", chart_type=ChartType.LINE)
    
    # Configure sub-charts
    chart.sub_charts.append({
        'type': ChartType.BAR,
        'series': [0],
        'bar_direction': 'col',
        'grouping': 'clustered',
        'gap_width': 150,
        'ax_ids': [70000000, 70000001]
    })
    chart.sub_charts.append({
        'type': ChartType.LINE,
        'series': [1],
        'ax_ids': [70000000, 70000002]
    })
    
    # Add axes
    chart.add_axis(axis_type="cat", axis_id=70000000, position="b")
    chart.add_axis(axis_type="val", axis_id=70000001, position="l")
    chart.add_axis(axis_type="val", axis_id=70000002, position="r")
    
    # Save
    output_path = "outputfiles/progcharts/combo_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_stock_chart():
    """Create a stock chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data for High-Low-Close stock chart
    ws.cells["A1"].value = "Date"
    ws.cells["B1"].value = "High"
    ws.cells["C1"].value = "Low"
    ws.cells["D1"].value = "Close"
    
    dates = ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2024-01-05"]
    highs = [105, 108, 110, 107, 112]
    lows = [100, 103, 105, 102, 108]
    closes = [103, 106, 108, 105, 110]
    
    for i, (date, high, low, close) in enumerate(zip(dates, highs, lows, closes), 2):
        ws.cells[f"A{i}"].value = date
        ws.cells[f"B{i}"].value = high
        ws.cells[f"C{i}"].value = low
        ws.cells[f"D{i}"].value = close
    
    # Create stock chart
    chart = ws.charts.add_stock(0, 4, 20, 12)
    chart.title = "Stock Price Movement"
    chart.category_data = "A2:A6"
    chart.stock_style = "high_low_close"
    chart.show_legend = False
    
    # Add series
    chart.n_series.add("B2:B6", category_data="A2:A6", name="High")
    chart.n_series.add("C2:C6", category_data="A2:A6", name="Low")
    chart.n_series.add("D2:D6", category_data="A2:A6", name="Close")
    
    # Save
    output_path = "outputfiles/progcharts/stock_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_surface_chart():
    """Create a surface chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data for surface chart (matrix)
    ws.cells["A1"].value = "X\\Y"
    ws.cells["B1"].value = 1
    ws.cells["C1"].value = 2
    ws.cells["D1"].value = 3
    ws.cells["E1"].value = 4
    
    ws.cells["A2"].value = 1
    ws.cells["A3"].value = 2
    ws.cells["A4"].value = 3
    ws.cells["A5"].value = 4
    
    # Create a surface function z = x^2 + y^2
    for i in range(2, 6):
        for j in range(2, 6):
            x = i - 1
            y = j - 1
            z = x**2 + y**2
            ws.cells[f"{chr(64+j)}{i}"].value = z
    
    # Create 3D surface chart
    chart = ws.charts.add_surface(0, 4, 20, 12, is_3d=True, wireframe=False)
    chart.title = "3D Surface Chart"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("B2:E5", category_data="A2:A5", name="Surface")
    
    # Save
    output_path = "outputfiles/progcharts/surface_chart_3d.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")
    
    # Create wireframe surface chart
    wb2 = Workbook()
    ws2 = wb2.worksheets[0]
    
    # Copy data
    for i in range(1, 6):
        for j in range(1, 6):
            if i == 1 and j == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = "X\\Y"
            elif i == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = j - 1
            elif j == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = i - 1
            else:
                x = i - 1
                y = j - 1
                z = x**2 + y**2
                ws2.cells[f"{chr(64+j)}{i}"].value = z
    
    chart2 = ws2.charts.add_surface(0, 4, 20, 12, is_3d=True, wireframe=True)
    chart2.title = "3D Wireframe Surface Chart"
    
    chart2.n_series.add("B2:E5", category_data="A2:A5", name="Surface")
    
    output_path2 = "outputfiles/progcharts/surface_chart_wireframe.xlsx"
    wb2.save(output_path2)
    assert os.path.exists(output_path2)
    print(f"Created: {output_path2}")


def test_create_radar_chart():
    """Create radar charts from scratch."""
    # Test standard radar chart
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Metric"
    ws.cells["B1"].value = "Product A"
    ws.cells["C1"].value = "Product B"
    ws.cells["D1"].value = "Product C"
    
    metrics = ["Quality", "Price", "Service", "Features", "Reliability"]
    data = [
        [8, 6, 9],
        [7, 8, 6],
        [9, 7, 8],
        [6, 9, 7],
        [8, 8, 9]
    ]
    
    for i, metric in enumerate(metrics, 2):
        ws.cells[f"A{i}"].value = metric
        for j, value in enumerate(data[i-2], 2):
            ws.cells[f"{chr(64+j)}{i}"].value = value
    
    # Create standard radar chart
    chart = ws.charts.add_radar(0, 4, 20, 12, radar_style="standard")
    chart.title = "Product Comparison - Standard Radar"
    chart.category_data = "A2:A6"
    chart.show_legend = True
    
    chart.n_series.add("B2:B6", category_data="A2:A6", name="Product A")
    chart.n_series.add("C2:C6", category_data="A2:A6", name="Product B")
    chart.n_series.add("D2:D6", category_data="A2:A6", name="Product C")
    
    output_path = "outputfiles/progcharts/radar_chart_standard.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")
    
    # Test filled radar chart
    wb2 = Workbook()
    ws2 = wb2.worksheets[0]
    
    # Copy data
    for i in range(1, 7):
        for j in range(1, 5):
            if i == 1 and j == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = "Metric"
            elif i == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = f"Product {chr(64+j)}"
            elif j == 1:
                ws2.cells[f"{chr(64+j)}{i}"].value = metrics[i-2]
            else:
                ws2.cells[f"{chr(64+j)}{i}"].value = data[i-2][j-2]
    
    chart2 = ws2.charts.add_radar(0, 4, 20, 12, radar_style="filled")
    chart2.title = "Product Comparison - Filled Radar"
    chart2.category_data = "A2:A6"
    chart2.show_legend = True
    
    chart2.n_series.add("B2:B6", category_data="A2:A6", name="Product A")
    chart2.n_series.add("C2:C6", category_data="A2:A6", name="Product B")
    chart2.n_series.add("D2:D6", category_data="A2:A6", name="Product C")
    
    output_path2 = "outputfiles/progcharts/radar_chart_filled.xlsx"
    wb2.save(output_path2)
    assert os.path.exists(output_path2)
    print(f"Created: {output_path2}")


def test_create_treemap_chart():
    """Create a treemap chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data
    ws.cells["A1"].value = "Category"
    ws.cells["B1"].value = "Subcategory"
    ws.cells["C1"].value = "Value"
    
    data = [
        ["Electronics", "Phones", 500],
        ["Electronics", "Laptops", 400],
        ["Electronics", "Tablets", 200],
        ["Clothing", "Shirts", 300],
        ["Clothing", "Pants", 250],
        ["Clothing", "Shoes", 200],
        ["Food", "Fruits", 150],
        ["Food", "Vegetables", 120],
        ["Food", "Dairy", 100]
    ]
    
    for i, (cat, sub, val) in enumerate(data, 2):
        ws.cells[f"A{i}"].value = cat
        ws.cells[f"B{i}"].value = sub
        ws.cells[f"C{i}"].value = val
    
    # Create treemap chart
    chart = ws.charts.add_treemap(0, 4, 20, 12)
    chart.title = "Sales by Category and Subcategory"
    chart.category_data = "A2:A10"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("C2:C10", category_data="A2:A10", name="Sales")
    
    # Save
    output_path = "outputfiles/progcharts/treemap_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_sunburst_chart():
    """Create a sunburst chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data - hierarchical structure with 3 levels
    ws.cells["A1"].value = "Category"
    ws.cells["B1"].value = "Subcategory"
    ws.cells["C1"].value = "Item"
    ws.cells["D1"].value = "Value"
    
    data = [
        ["Electronics", "Phones", "iPhone", 200],
        ["Electronics", "Phones", "Samsung", 150],
        ["Electronics", "Laptops", "MacBook", 180],
        ["Electronics", "Laptops", "Dell", 120],
        ["Clothing", "Shirts", "T-Shirt", 100],
        ["Clothing", "Shirts", "Polo", 80],
        ["Clothing", "Pants", "Jeans", 90],
        ["Clothing", "Pants", "Chinos", 70],
        ["Food", "Fruits", "Apples", 60],
        ["Food", "Fruits", "Oranges", 50],
        ["Food", "Vegetables", "Carrots", 40],
        ["Food", "Vegetables", "Broccoli", 35]
    ]
    
    for i, (cat, sub, item, val) in enumerate(data, 2):
        ws.cells[f"A{i}"].value = cat
        ws.cells[f"B{i}"].value = sub
        ws.cells[f"C{i}"].value = item
        ws.cells[f"D{i}"].value = val
    
    # Create sunburst chart
    chart = ws.charts.add_sunburst(0, 4, 20, 12)
    chart.title = "Sales Hierarchy - Sunburst"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("D2:D13", category_data="A2:A13", name="Sales")
    
    # Save
    output_path = "outputfiles/progcharts/sunburst_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_histogram_chart():
    """Create a histogram chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data - distribution of test scores
    ws.cells["A1"].value = "Score"
    
    scores = [65, 72, 78, 82, 85, 88, 90, 92, 95, 98, 100, 102, 105, 108, 110, 112, 115, 118, 120, 125,
              68, 75, 80, 84, 87, 89, 91, 93, 96, 99, 101, 104, 106, 109, 111, 113, 116, 119, 122, 128]
    
    for i, score in enumerate(scores, 2):
        ws.cells[f"A{i}"].value = score
    
    # Create histogram chart
    chart = ws.charts.add_histogram(0, 4, 20, 12)
    chart.title = "Score Distribution - Histogram"
    chart.show_legend = False
    
    # Configure histogram bins
    chart.histogram_bin_count = 10  # Divide into 10 bins
    chart.histogram_interval_closed = "r"  # Right-closed intervals
    
    # Add series
    chart.n_series.add("A2:A41", name="Scores")
    
    # Save
    output_path = "outputfiles/progcharts/histogram_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")
    
    # Create histogram with bin size instead of count
    wb2 = Workbook()
    ws2 = wb2.worksheets[0]
    
    # Copy data
    for i, score in enumerate(scores, 2):
        ws2.cells[f"A{i}"].value = score
    
    chart2 = ws2.charts.add_histogram(0, 4, 20, 12)
    chart2.title = "Score Distribution - Histogram (Bin Size)"
    chart2.show_legend = False
    
    # Configure histogram with bin size
    chart2.histogram_bin_size = 10  # Each bin is 10 units wide
    chart2.histogram_interval_closed = "r"
    
    chart2.n_series.add("A2:A41", name="Scores")
    
    output_path2 = "outputfiles/progcharts/histogram_chart_binsize.xlsx"
    wb2.save(output_path2)
    assert os.path.exists(output_path2)
    print(f"Created: {output_path2}")


def test_create_funnel_chart():
    """Create a funnel chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data - sales funnel stages
    ws.cells["A1"].value = "Stage"
    ws.cells["B1"].value = "Count"
    
    stages = ["Website Visitors", "Product Page Views", "Add to Cart", "Checkout", "Purchase"]
    counts = [10000, 5000, 2000, 1000, 500]
    
    for i, (stage, count) in enumerate(zip(stages, counts), 2):
        ws.cells[f"A{i}"].value = stage
        ws.cells[f"B{i}"].value = count
    
    # Create funnel chart
    chart = ws.charts.add_funnel(0, 4, 20, 12)
    chart.title = "Sales Funnel"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("B2:B6", category_data="A2:A6", name="Funnel")
    
    # Save
    output_path = "outputfiles/progcharts/funnel_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_map_chart():
    """Create a map (region map) chart from scratch."""
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Add data - sales by region
    ws.cells["A1"].value = "Region"
    ws.cells["B1"].value = "Sales"
    
    regions = ["California", "Texas", "New York", "Florida", "Illinois", "Pennsylvania", "Ohio", "Georgia"]
    sales = [500000, 450000, 400000, 350000, 300000, 280000, 250000, 220000]
    
    for i, (region, sale) in enumerate(zip(regions, sales), 2):
        ws.cells[f"A{i}"].value = region
        ws.cells[f"B{i}"].value = sale
    
    # Create map chart
    chart = ws.charts.add_map(0, 4, 20, 12)
    chart.title = "Sales by Region"
    chart.show_legend = True
    
    # Add series
    chart.n_series.add("B2:B9", category_data="A2:A9", name="Sales")
    
    # Save
    output_path = "outputfiles/progcharts/map_chart.xlsx"
    wb.save(output_path)
    assert os.path.exists(output_path)
    print(f"Created: {output_path}")


def test_create_all_charts():
    """Create all supported chart types in one test."""
    # Ensure output directory exists
    os.makedirs("outputfiles/progcharts", exist_ok=True)
    
    print("\n=== Creating All Supported Chart Types ===\n")
    print("Note: Box and Whisker chart creation is not yet implemented.\n")
    
    test_create_line_chart()
    test_create_bar_chart()
    test_create_pie_chart()
    test_create_area_chart()
    # test_create_box_whisker_chart()
    test_create_waterfall_chart()
    test_create_scatter_chart()
    test_create_combo_chart()
    test_create_stock_chart()
    test_create_surface_chart()
    test_create_radar_chart()
    test_create_treemap_chart()
    test_create_sunburst_chart()
    test_create_histogram_chart()
    test_create_funnel_chart()
    test_create_map_chart()
    
    print("\n=== All Supported Charts Created Successfully ===\n")
    print("Output files saved to: outputfiles/progcharts/")


if __name__ == "__main__":
    test_create_all_charts()
