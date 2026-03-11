# Aspose.Cells FOSS

A Python library for creating, reading, and modifying Excel files (.xlsx format) without requiring Microsoft Excel.

[![PyPI version](https://badge.fury.io/py/aspose-cells-foss.svg)](https://badge.fury.io/py/aspose-cells-foss)
[![Python](https://img.shields.io/pypi/pyversions/aspose-cells-foss.svg)](https://pypi.org/project/aspose-cells-foss/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Create & Edit Excel Files**: Create new workbooks or modify existing .xlsx files
- **Cell Operations**: Read/write cell values and formulas
- **Styling**: Apply fonts, colors, borders, number formats, and alignment
- **Multiple Worksheets**: Add, remove, rename, and manage worksheets
- **Charts**: Create and modify charts (line, bar, pie, scatter, combo, waterfall, treemap, and more)
- **Tables**: Add and manage ListObject tables with styles and auto-filters
- **Shapes & TextBoxes**: Add shapes, text boxes, and pictures with hyperlinks
- **Data Validation**: Add dropdown lists, number ranges, and custom validation rules
- **Comments**: Add and manage cell comments with author and rich text
- **Hyperlinks**: Create links to URLs, emails, files, and internal references
- **Auto-Filters**: Apply filtering to data ranges
- **Conditional Formatting**: Apply rules-based formatting
- **CSV / JSON / Markdown Export**: Save workbooks in multiple text formats
- **Encryption**: Password-protect Excel files with AES encryption
- **Workbook & Worksheet Protection**: Protect workbook structure and individual sheets

## Installation

```bash
pip install aspose-cells-foss
```

## Quick Start

### Create a new Excel file

```python
from aspose.cells_foss import Workbook

# Create a new workbook
workbook = Workbook()

# Get the first worksheet
worksheet = workbook.worksheets[0]

# Set cell values
worksheet.cells["A1"].value = "Hello"
worksheet.cells["B1"].value = "World"
worksheet.cells["A2"].value = 42
worksheet.cells["B2"].value = 3.14

# Save the workbook
workbook.save("output.xlsx")
```

### Read an existing Excel file

```python
from aspose.cells_foss import Workbook

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
from aspose.cells_foss import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
cell = worksheet.cells["A1"]

cell.value = "Styled Text"

# Get and modify the cell style
style = cell.get_style()
style.font.bold = True
style.font.color = "#FF0000"  # Red
style.font.size = 14
cell.apply_style(style)

workbook.save("styled.xlsx")
```

### Add data validation (dropdown list)

```python
from aspose.cells_foss import Workbook, DataValidationType

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add a dropdown list validation to A1:A10
validation = worksheet.data_validations.add("A1:A10")
validation.type = DataValidationType.LIST
validation.formula1 = '"Option1,Option2,Option3"'

workbook.save("validation.xlsx")
```

### Export to CSV

```python
from aspose.cells_foss import Workbook

workbook = Workbook("input.xlsx")
workbook.save_as_csv("output.csv")
```

### Password protection

```python
from aspose.cells_foss import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
worksheet.cells["A1"].value = "Confidential Data"

# Save with password protection
workbook.save("protected.xlsx", password="mypassword")

# Open a password-protected file
workbook2 = Workbook("protected.xlsx", password="mypassword")
```

## Requirements

- Python 3.7 or higher
- pycryptodome >= 3.15.0
- olefile >= 0.46

## Documentation

For more examples and detailed API documentation, see the [examples](https://github.com/aspose-cells-foss/aspose-cells-python/tree/main/examples) directory.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/aspose-cells-foss/aspose-cells-python/blob/main/License/license.txt) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/aspose-cells-foss/aspose-cells-python/issues)
