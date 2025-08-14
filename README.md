# Aspose.Cells.Python

High-performance Python Excel processing library with advanced conversion capabilities. Can be used standalone or as a MarkItDown plugin for superior Excel-to-Markdown conversion.

[![License: Split](https://img.shields.io/badge/License-Split-blue.svg)](https://aspose.org/pricing)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

## Overview

Aspose.Cells.Python is a high-performance Excel processing library that provides a clean, Pythonic API for working with Excel files. It can be used as a standalone library or integrated with MarkItDown for enhanced Excel-to-Markdown conversion featuring hyperlink preservation, improved formatting, and 2-3x faster processing speeds.


## Quick Start

### Installation

**Note: Package not yet published to PyPI. Install from source:**

```bash
# Clone and install
git clone https://github.com/aspose-cells/aspose-cells-python.git
cd aspose-cells-python
pip install -e .

# With MarkItDown plugin support
pip install -e .[markitdown]
```



### Basic Usage

```python
from aspose.cells import Workbook, FileFormat

# Create and populate workbook
wb = Workbook()
ws = wb.active

# Multiple ways to set data
ws['A1'] = "Product"        # Excel-style
ws[0, 1] = "Price"          # Python-style (0-based)
ws.cell(1, 3).value = "Qty" # Traditional method

# Batch data operations
ws.append(["Laptop", 999.99, 5])
ws.append(["Mouse", 29.95, 50])

# Save in multiple formats
wb.save("products.xlsx", FileFormat.XLSX)
wb.save("products.csv", FileFormat.CSV)

# Export as strings
json_str = wb.exportAs(FileFormat.JSON)
markdown_str = wb.exportAs(FileFormat.MARKDOWN)
```

## Performance Comparison

### MarkItDown Default vs Aspose Plugin

Our plugin provides significant advantages over MarkItDown's default Excel converter:

**Key Advantages:**
- ⚡ Faster processing performance (2-3x speed improvement in typical cases)
- ✅ Hyperlink preservation and formatting
- ✅ Better handling of merged cells and complex layouts
- ✅ Cleaner table structure without empty columns/rows
- ✅ Enhanced metadata support

#### Example Comparison

**Original Excel File:**
![Excel Test File](images/test.png)

**Default MarkItDown Output:**
```markdown
## Product Details
| Product Detailed Information Including Inventory, Cost, and Profit Analysis | Unnamed: 1 | Unnamed: 2 | Unnamed: 3 | Unnamed: 4 | Unnamed: 5 | Unnamed: 6 |
| --- | --- | --- | --- | --- | --- | --- |
| NaN | NaN | NaN | NaN | NaN | NaN | NaN |
| NaN | NaN | NaN | NaN | NaN | NaN | NaN |
| Product ID | Product Name | Stock Quantity | Purchase Cost | Sale Price | Profit | Stock Status |
| P001 | Lenovo ThinkPad X1 | 45 | 4500 | 5999.99 | 1499.99 | Low Stock |
| P002 | Dell XPS 13 | 23 | 4200 | 5699.99 | 1499.99 | Low Stock |
```

**Aspose Plugin Output:**
```markdown
<!-- Generator: Aspose.Cells.Python MarkItDown Plugin -->

## Product Details

| Product Detailed Information Including Inventory, Cost, and Profit Analysis | B | C | D | E | F | G |
| --- | --- | --- | --- | --- | --- | --- |
| Product ID | Product Name | Stock Quantity | Purchase Cost | Sale Price | Profit | Stock Status |
| P001 | [Lenovo ThinkPad X1](https://www.lenovo.com/thinkpad-x1) | 45 | 4500 | 5,999.99 | 1,499.99 | Low Stock |
| P002 | [Dell XPS 13](https://www.dell.com/xps-13) | 23 | 4200 | 5,699.99 | 1,499.99 | Low Stock |
| P003 | [MacBook Air M2](https://www.apple.com/macbook-air-m2) | 67 | 7000 | 8,999.99 | 1,999.99 | Normal Stock |
```

*Notice the preserved hyperlinks, better formatting, and cleaner output.*

## MarkItDown Plugin Integration

The library includes a plugin for Microsoft MarkItDown that provides enhanced Excel-to-Markdown conversion:

### Command Line Usage

```bash
# Convert Excel file using our plugin
markitdown --use-plugins test.xlsx -o test.md

# List available plugins
markitdown --list-plugins
```

### Python API Usage

```python
# Install with plugin support
pip install aspose-cells-python[markitdown]

OR

# Install plugin from source
pip install -e .[markitdown]

# The plugin is automatically registered
from markitdown import MarkItDown

md = MarkItDown(enable_plugins=True)
result = md.convert("spreadsheet.xlsx")
print(result.text_content)
```

**Plugin Features:**
- Document metadata and conversion info
- Multi-sheet processing with headers
- Professional table formatting



## Advanced Features

### Styling and Formatting

```python
# Cell styling
ws['A1'].font.bold = True
ws['A1'].font.size = 14
ws['A1'].font.color = "blue"
ws['A1'].fill.color = "lightgray"

# Range styling
ws['A1:C1'].font.bold = True
ws['A1:C1'].fill.color = "lightblue"
```

### Data Import from Records

```python
# From dictionary list
records = [
    {"name": "Alice", "age": 25, "city": "New York"},
    {"name": "Bob", "age": 30, "city": "London"}
]
ws.from_records(records)

# From lists
data = [
    ["Charlie", 35, "Tokyo"],
    ["Diana", 28, "Paris"]
]
ws.extend(data)
```

## License

This project is licensed under the Aspose Split License Agreement - see the [LICENSE](license/Aspose_Split-License-Agreement_2025-07-08_WIP.txt) file for details.

Part of the [Aspose.org](https://aspose.org) open source ecosystem.

## Requirements

- Python 3.8+
- Optional: markitdown>=0.1.0 (for MarkItDown plugin)