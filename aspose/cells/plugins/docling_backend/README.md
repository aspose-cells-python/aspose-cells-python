# Aspose.Cells Docling Backend

Document conversion backend for IBM Docling, providing Excel processing capabilities through the Aspose.Cells.Python engine.

## Overview

This backend integrates Aspose.Cells.Python with IBM Docling to deliver:

- **Excel Processing**: Handles Excel files with formulas and basic formatting
- **Multi-format Export**: Convert Excel to Markdown, JSON, and other formats supported by Docling
- **Document Structure**: Pagination and document model integration
- **Seamless Integration**: Drop-in replacement for default Docling Excel backend

## Installation

```bash
# Install from source with Docling backend support
pip install -e .[docling]
```

## Quick Start

### Basic Document Conversion

```python
from docling.datamodel.base_models import InputFormat
from docling.datamodel.document import InputDocument
from docling.document_converter import DocumentConverter
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

# Export to markdown
markdown_content = document.export_to_markdown()
print(markdown_content)
```

## Features

- **Excel Support**: Handles spreadsheets with formulas and images
- **Multi-page Documents**: Proper pagination support with sheet-based page structure  
- **Export Flexibility**: Support for Markdown, JSON, and other Docling export formats
- **Document Model Integration**: Full compatibility with Docling's document structure


## Supported Formats

The backend currently supports:

- **Input**: Excel files (.xlsx only)
- **Output**: Markdown, JSON (via Docling document model)

## Requirements

- Python 3.8+
- docling
- aspose-cells-python

## License

Part of [Aspose.Cells.Python](../../) under the Aspose Split License Agreement.