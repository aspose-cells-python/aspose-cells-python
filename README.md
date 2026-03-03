# Aspose.Cells for Python

A lightweight Python library for creating, reading, and modifying Excel files (.xlsx format) without requiring Microsoft Excel.

[![PyPI version](https://img.shields.io/pypi/v/aspose-cells-foss)](https://pypi.org/project/aspose-cells-foss/)
[![Python](https://img.shields.io/pypi/pyversions/aspose-cells-foss)](https://pypi.org/project/aspose-cells-foss/)


## Features

- **Create & Edit Excel Files**: Create new workbooks or modify existing .xlsx files
- **Cell Operations**: Read/write cell values, formulas, and apply formatting
- **Styling**: Apply fonts, colors, borders, number formats, and alignment
- **Multiple Worksheets**: Add, remove, and manage worksheets
- **Data Validation**: Add dropdown lists, number ranges, and custom validation rules
- **Comments**: Add and manage cell comments with author and rich text
- **Hyperlinks**: Create links to URLs, emails, files, and internal references
- **Auto-Filters**: Apply filtering to data ranges
- **Conditional Formatting**: Apply rules-based formatting
- **CSV/JSON/Markdown Export**: Save workbooks in multiple formats
- **Encryption**: Password-protect Excel files with AES encryption
- **Workbook Protection**: Protect workbook structure and worksheets
- **Charts**: Create 16 chart types — Line, Bar, Pie, Area, Scatter, Waterfall, Combo, Stock, Surface, Radar, Treemap, Sunburst, Histogram, Funnel, Box & Whisker, and Map
- **Pictures**: Embed and anchor images (JPEG, PNG, etc.) to worksheet cells
- **Drawing Shapes**: Add 30+ preset shapes — rectangles, ovals, arrows, stars, text boxes, callouts, and more
- **Sparklines**: Embed mini-charts (Line, Column, Win-Loss) directly inside cells
- **Excel Tables**: Create and style structured tables (ListObject) with auto-filter and column headers
- **Manual Page Breaks**: Add and remove horizontal and vertical page breaks
- **Formula Evaluator**: Evaluate basic formulas and cell references at read time
- **Merge Cells**: Merge and unmerge cell ranges
- **Print Area**: Define the print area for a worksheet

## Installation

```bash
pip install aspose-cells-foss
```

## Quick Start

### Create a new Excel file

```python
from aspose_cells import Workbook

# Create a new workbook
workbook = Workbook()

# Get the first worksheet
worksheet = workbook.worksheets[0]

# Set cell values
worksheet.cells["A1"].put_value("Hello")
worksheet.cells["B1"].put_value("World")
worksheet.cells["A2"].put_value(42)
worksheet.cells["B2"].put_value(3.14)

# Save the workbook
workbook.save("output.xlsx")
```

### Read an existing Excel file

```python
from aspose_cells import Workbook

# Open an existing workbook
workbook = Workbook("input.xlsx")

# Access a worksheet
worksheet = workbook.worksheets[0]

# Read cell values
value = worksheet.cells["A1"].value
print(f"Cell A1 contains: {value}")
```

### Apply styling

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
cell = worksheet.cells["A1"]

cell.put_value("Styled Text")

# Get and modify the cell style
style = cell.get_style()
style.font.is_bold = True
style.font.color = "FF0000"  # Red
style.font.size = 14
cell.set_style(style)

workbook.save("styled.xlsx")
```

### Add data validation (dropdown list)

```python
from aspose_cells import Workbook, DataValidationType

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add a dropdown list validation
validation = worksheet.data_validations.add()
validation.type = DataValidationType.LIST
validation.formula1 = '"Option1,Option2,Option3"'
validation.add_area("A1:A10")

workbook.save("validation.xlsx")
```

### Export to CSV

```python
from aspose_cells import Workbook, SaveFormat

workbook = Workbook("input.xlsx")
workbook.save("output.csv", SaveFormat.CSV)
```

### Password protection

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
worksheet.cells["A1"].put_value("Confidential Data")

# Save with password protection
workbook.save("protected.xlsx", password="mypassword")

# Open a password-protected file
workbook2 = Workbook("protected.xlsx", password="mypassword")
```

### Create a chart

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add data
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
sales  = [100, 150, 120, 180, 200, 170]
for i, (m, s) in enumerate(zip(months, sales), 2):
    worksheet.cells[f"A{i}"].value = m
    worksheet.cells[f"B{i}"].value = s

# Create a line chart anchored to rows 0-20, columns 4-12
chart = worksheet.charts.add_line(0, 4, 20, 12)
chart.title = "Monthly Sales"
chart.n_series.add("B2:B7", category_data="A2:A7", name="Sales")

workbook.save("chart.xlsx")
```

### Add a picture

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Embed an image anchored between two cells (0-based row/column indices)
worksheet.pictures.add("logo.png",
    upper_left_row=1, upper_left_column=1,
    lower_right_row=8, lower_right_column=5)

workbook.save("with_picture.xlsx")
```

### Add drawing shapes

```python
from aspose_cells import Workbook, MsoDrawingType, FillType, TextAlignmentType, TextAnchorType

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add a rounded rectangle
shape = worksheet.shapes.add(
    MsoDrawingType.ROUNDED_RECTANGLE,
    upper_left_row=1, upper_left_column=1,
    lower_right_row=5, lower_right_column=5
)
shape.text = "Hello!"
shape.fill.fore_color = "90EE90"          # light green fill
shape.font.bold = True
shape.text_horizontal_alignment = TextAlignmentType.CENTER
shape.text_vertical_alignment = TextAnchorType.MIDDLE

# Add a text box
textbox = worksheet.shapes.add_text_box(7, 1, 11, 8)
textbox.text = "Notes go here"

workbook.save("shapes.xlsx")
```

### Add sparklines

```python
from aspose_cells import Workbook, SparklineType

workbook = Workbook()
worksheet = workbook.worksheets[0]
worksheet.name = "Sales"

# Data in B2:F6, sparklines displayed in G2:G6
group = worksheet.sparkline_groups.add(
    sparkline_type=SparklineType.LINE,
    data_range="Sales!B2:F6",
    is_vertical=False,
    location_range="G2:G6"
)
group.color_series = "0070C0"   # blue sparkline line

workbook.save("sparklines.xlsx")
```

### Create an Excel table

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Write headers and data
headers = ["Product", "Qty", "Price"]
worksheet.cells["A1"].value, worksheet.cells["B1"].value, worksheet.cells["C1"].value = headers
worksheet.cells["A2"].value, worksheet.cells["B2"].value, worksheet.cells["C2"].value = "Widget", 10, 9.99

# Create a table over A1:C2 (0-based indices)
table = worksheet.tables.add(start_row=0, start_col=0,
                              end_row=1, end_col=2,
                              has_headers=True, name="ProductTable")
table.table_style_info.name = "TableStyleMedium9"

workbook.save("table.xlsx")
```

### Add manual page breaks

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add a horizontal page break before row 20 (0-based)
worksheet.horizontal_page_breaks.add(19)

# Add a vertical page break before column D (0-based index 3)
worksheet.vertical_page_breaks.add(3)

workbook.save("page_breaks.xlsx")
```

## Requirements

- Python 3.7 or higher
- pycryptodome >= 3.15.0
- olefile >= 0.46

## Documentation

For more examples and detailed API documentation, see the [examples](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Python/tree/main/examples) directory.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [license.txt](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Python/blob/main/License/LICENSE.txt) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Python/issues)
