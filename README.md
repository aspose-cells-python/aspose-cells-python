# Aspose.Cells for Python

Modern Pythonic Excel processing library with advanced plugin ecosystem for enhanced document conversion.

[![License: Split](https://img.shields.io/badge/License-Split-blue.svg)](https://aspose.org/)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

## Overview

Aspose.Cells for Python is a high-performance Excel processing library that provides a clean, Pythonic API for working with Excel files. It offers both standalone functionality and seamless integration with popular document conversion frameworks through its plugin ecosystem.

## Installation

**Note: Package not yet published to PyPI. Install from source:**

```bash
# Core library
git clone https://github.com/aspose-cells/aspose-cells-python.git
cd aspose-cells-python-org
pip install -e .

# With plugin support
pip install -e .[markitdown]  # MarkItDown plugin
pip install -e .[docling]     # Docling backend
```

## Quick Start

```python
from aspose.cells import Workbook, FileFormat

# Create and populate workbook
wb = Workbook()
ws = wb.active

# Multiple ways to set data
ws['A1'] = "Product"        # Excel-style
ws[0, 1] = "Price"          # Python-style (0-based)
ws.cell(1, 3).value = "Qty" # Traditional method

# Batch operations
ws.append(["Laptop", 999.99, 5])
ws.append(["Mouse", 29.95, 50])

# Export to multiple formats
wb.save("products.xlsx", FileFormat.XLSX)
wb.save("products.csv", FileFormat.CSV)
json_str = wb.exportAs(FileFormat.JSON)
markdown_str = wb.exportAs(FileFormat.MARKDOWN)
```

## Core Features

### Multi-Access Cell Interface
Three flexible ways to work with cells:
- **Excel-style**: `ws['A1'] = value`
- **Python-style**: `ws[0, 1] = value` (0-based indexing)  
- **Traditional**: `ws.cell(1, 3).value = value`

### Styling and Formatting
```python
# Cell styling
ws['A1'].font.bold = True
ws['A1'].font.size = 14
ws['A1'].fill.color = "lightgray"

# Range styling
ws['A1:C1'].font.bold = True
```

### Data Import/Export
```python
# From records
records = [{"name": "Alice", "age": 25}, {"name": "Bob", "age": 30}]
ws.from_records(records)

# Export formats
formats = [FileFormat.XLSX, FileFormat.CSV, FileFormat.JSON, FileFormat.MARKDOWN]
```

## Plugin Ecosystem

Aspose.Cells for Python extends functionality through specialized plugins for popular document conversion frameworks:

### MarkItDown Plugin
Enhanced Excel-to-Markdown conversion with improved formatting capabilities.

**Key Benefits:**
- Hyperlink preservation and clean table formatting  
- Configurable conversion parameters
- Enhanced output quality

```bash
# Install and use
pip install -e .[markitdown]
markitdown --use-plugins spreadsheet.xlsx -o output.md
```

**→ [View MarkItDown Plugin Documentation](aspose/cells/plugins/markitdown_plugin/README.md)**

### Docling Backend
Document backend for IBM Docling framework.

**Key Benefits:**
- Enhanced Excel processing for complex spreadsheets
- Multi-format export (Markdown, JSON)
- Advanced document structure support

```bash
# Install and use
pip install -e .[docling]
```

```python
from docling.datamodel.base_models import InputFormat
from docling.datamodel.document import InputDocument
from aspose.cells.plugins.docling_backend import CellsDocumentBackend

# Create input document with Aspose backend
input_doc = InputDocument(
    path_or_stream="spreadsheet.xlsx",
    format=InputFormat.XLSX,
    backend=CellsDocumentBackend,
    filename="spreadsheet.xlsx"
)

# Initialize backend and convert
backend = CellsDocumentBackend(in_doc=input_doc, path_or_stream="spreadsheet.xlsx")
document = backend.convert()
markdown_content = document.export_to_markdown()
```

**→ [View Docling Backend Documentation](aspose/cells/plugins/docling_backend/README.md)**

## License

This project is licensed under the Aspose Split License Agreement - see the [LICENSE](license/Aspose_Split-License-Agreement_2025-07-08_WIP.txt) file for details.

Part of the [Aspose.org](https://aspose.org) open source ecosystem.

## Requirements

- Python 3.8+
- Optional: markitdown>=0.1.0 (for MarkItDown plugin)
- Optional: docling (for Docling backend)