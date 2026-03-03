# Aspose.Cells for Python — Development Guide for AI Agents

## Do
- Mirror the [Aspose.Cells for .NET](https://reference.aspose.com/cells/net/) public API as closely as possible — class names, property names, and method signatures should match their .NET counterparts
- Use `snake_case` for the primary Python method names and add a `PascalCase` alias for every public `add`, `remove`, `clear`, and `get` method (e.g., `add()` + `Add()`)
- Represent colors as 6-digit RRGGBB hex strings — **no `#` prefix** (e.g., `"FF0000"` not `"#FF0000"`)
- Use 0-based row and column indices throughout the library
- Preserve unrecognised XML by storing it in `_source_*` attributes on model objects and writing it back verbatim during save (round-trip fidelity)
- Call `worksheet._mark_drawing_dirty()` whenever a picture or shape is added or modified, so drawing XML is regenerated on save
- Use `_add_loaded()` (not `add()`) inside XML loaders — loaders must never trigger side effects like marking the workbook dirty
- Add every new public class and enum to `__all__` in `aspose_cells/__init__.py`
- Bump the version string in **both** `pyproject.toml` and `aspose_cells/__init__.py` when releasing
- Write at least one integration test in `examples/` for every new feature, covering create → save → reload → verify
- Follow ECMA-376 / SpreadsheetML element names and attribute names in XML loader/saver comments so the code is auditable against the spec
- Use `xml.etree.ElementTree` (stdlib) for XML reading and writing — no third-party XML libraries

## Don't
- Never add a runtime dependency that is not already in `pyproject.toml` without a discussion — the library must stay lightweight
- Never use mutable default arguments in function signatures (`def foo(items=[])` is a bug)
- Never silently swallow exceptions in loaders — log or re-raise with context so failures are diagnosable
- Never store absolute file paths in saved XLSX part relationships — always use relative paths (e.g., `../media/image1.png`)
- Never break round-trip fidelity: loading a file and saving it without changes must not corrupt or drop any XML parts
- Never expose internal `_source_*` XML fields in the public API
- Never write colors with a `#` prefix — Excel's OOXML format does not use it
- Never use 1-based row/column indices internally — all internal coordinates are 0-based; convert only at the user-facing boundary if needed
- Never add `print()` statements to library code — use them only in `examples/` scripts
- Never modify auto-generated files — rebuild from source instead

## Commands

```bash
# Install the library in editable mode (recommended during development)
pip install -e .

# Run all example/integration tests
pytest examples/ -v

# Run a single test file
pytest examples/test_create_all_charts.py -v

# Build a distribution package
python -m build

# Upload to TestPyPI (verify before real release)
upload_to_testpypi.bat

# Upload to PyPI
upload_to_pypi.bat
```

## Project Structure

```
aspose_cells/                  # Main library package
├── __init__.py                # Public API surface — exports and __version__
├── workbook.py                # Workbook class, SaveFormat enum
├── worksheet.py               # Worksheet class
├── cell.py                    # Cell class
├── cells.py                   # Cells collection and coordinate helpers
├── style.py                   # Style, Font, NumberFormat
├── chart.py                   # Chart, ChartCollection, ChartType, NSeries, …
├── picture.py                 # Picture, PictureCollection
├── shape.py                   # Shape, ShapeCollection, MsoDrawingType, …
├── sparkline.py               # SparklineGroup, SparklineGroupCollection, …
├── table.py                   # Table, TableCollection, TableStyleInfo, …
├── page_break.py              # HorizontalPageBreakCollection, VerticalPageBreakCollection
├── formula_evaluator.py       # FormulaEvaluator (basic formula evaluation)
├── data_validation.py         # DataValidation, DataValidationType, …
├── xml_loader.py              # Top-level XLSX ZIP reader (orchestrates sub-loaders)
├── xml_saver.py               # Top-level XLSX ZIP writer (orchestrates sub-savers)
├── xml_chart_loader.py        # Chart XML ↔ Chart objects
├── xml_chart_saver.py
├── xml_picture_loader.py      # Drawing XML ↔ Picture objects
├── xml_picture_saver.py
├── xml_shape_loader.py        # Drawing XML ↔ Shape objects
├── xml_shape_saver.py
├── xml_sparkline_loader.py    # Worksheet extLst ↔ SparklineGroup objects
├── xml_sparkline_saver.py
├── xml_table_loader.py        # table*.xml parts ↔ Table objects
├── xml_table_saver.py
├── xml_hyperlink_handler.py
├── xml_properties_loader.py
├── xml_properties_saver.py
├── xml_autofilter_loader.py
├── xml_autofilter_saver.py
├── xml_conditional_format_loader.py
├── xml_conditional_format_saver.py
├── xlsx_encryptor.py          # AES encryption/decryption of .xlsx files
├── encryption_params.py
├── cfb_handler.py             # Compound File Binary (OLE) detection
├── csv_handler.py
├── markdown_handler.py
├── json_handler.py
└── workbook_properties.py

examples/                      # Integration tests (pytest-compatible)
│   test_create_all_charts.py
│   test_create_picture.py
│   test_create_shape.py
│   test_create_sparkline.py
│   test_create_exceltable.py
│   test_manual_page_breaks.py
│   test_merge_cells.py
│   test_print_area.py
│   test_encryption.py
│   test_data_validation.py
│   … (one file per feature)

License/                       # License text
pyproject.toml                 # Build metadata and dependencies
README.md                      # User-facing documentation
AGENTS.md                      # This file
```

## Architecture

### XLSX as a ZIP

An `.xlsx` file is a ZIP archive of XML parts. The library reads and writes the archive directly using Python's `zipfile` module — no Excel or COM automation required.

### Loader / Saver pattern

Every feature follows this split:

| File | Responsibility |
|------|---------------|
| `aspose_cells/<feature>.py` | Model classes and enums (pure Python, no I/O) |
| `aspose_cells/xml_<feature>_loader.py` | Deserialise XML part(s) → populate model objects |
| `aspose_cells/xml_<feature>_saver.py` | Serialise model objects → produce XML part(s) |

The top-level `xml_loader.py` and `xml_saver.py` orchestrate all sub-loaders and sub-savers and manage the ZIP archive.

### Round-trip preservation

When a loader encounters XML it doesn't fully model (unknown attributes, extension elements, etc.), it stores the raw XML string in a `_source_*` attribute (e.g., `_source_xml`, `_source_part_path`, `_source_blip_extLst_xml`). The corresponding saver writes it back verbatim. This keeps files valid in Excel even when features are only partially implemented.

### Drawing parts

Pictures and shapes share a single `xl/drawings/drawingN.xml` part per worksheet. Both `xml_picture_saver.py` and `xml_shape_saver.py` write into that part. When either collection is mutated, `worksheet._mark_drawing_dirty()` is called so the combined drawing XML is regenerated on save.

### Sparklines

Sparklines are stored as `<x14:sparklineGroup>` elements inside the worksheet's `<extLst>` — they do **not** use a separate part file. The loader/saver appends to or parses the worksheet XML's extension list directly.

### Tables

Each table is a separate XML part at `xl/tables/tableN.xml`. The saver writes one file per `Table` object and registers the relationship in `xl/worksheets/_rels/sheetN.xml.rels`.

## Code Conventions

### Dual naming (Python + .NET style)
```python
# Primary: snake_case
def add(self, row, col): ...

# Alias: PascalCase (for .NET-style callers)
def Add(self, row, col):
    return self.add(row, col)
```

### Color strings
```python
# Good — 6-digit RRGGBB, no hash
shape.fill.fore_color = "FF0000"
style.font.color = "0070C0"

# Bad — hash prefix breaks XML serialisation
shape.fill.fore_color = "#FF0000"
```

### 0-based coordinates
```python
# Good
worksheet.horizontal_page_breaks.add(19)   # break before row 20
worksheet.pictures.add("img.png", upper_left_row=0, upper_left_column=0,
                        lower_right_row=5, lower_right_column=3)

# Bad — do not convert to 1-based inside the library
worksheet.cells["A1"]  # A1 notation is fine at the user boundary
```

### Loader vs public add
```python
# In xml_*_loader.py — does NOT mark dirty, does NOT trigger side-effects
collection._add_loaded(obj)

# In user code / public API — triggers dirty flag, validation, etc.
collection.add(...)
```

### Adding a new feature (checklist)

1. Create `aspose_cells/<feature>.py` — model classes and enums
2. Create `aspose_cells/xml_<feature>_loader.py` — reads XML → model
3. Create `aspose_cells/xml_<feature>_saver.py` — model → XML
4. Wire the loader into `xml_loader.py` and the saver into `xml_saver.py`
5. Export all public classes from `aspose_cells/__init__.py` and add to `__all__`
6. Add a `worksheet.<feature_collection>` property to `worksheet.py` if applicable
7. Write an integration test in `examples/test_<feature>.py`
8. Document the feature in `README.md` under **Features** and add a **Quick Start** example

## PR Checklist

- All example tests pass: `pytest examples/ -v`
- New classes and enums are exported from `__init__.py` and listed in `__all__`
- PascalCase aliases exist for every new `add` / `remove` / `clear` method
- Colors use 6-digit RRGGBB strings (no `#`)
- Round-trip test: load a file that exercises the feature, save, reload, verify no data lost
- Version bumped in `pyproject.toml` **and** `aspose_cells/__init__.py`
- `README.md` updated if user-facing behaviour changed

## When Stuck

- Consult the [ECMA-376 Part 1 spec](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) for correct XML element and attribute names — the relevant clause number is usually noted in the module docstring
- Mirror the [Aspose.Cells for .NET API reference](https://reference.aspose.com/cells/net/) for class and property names when unsure
- Open an existing feature's loader/saver pair as a template — e.g., `xml_shape_loader.py` / `xml_shape_saver.py` is a good model for new drawing features
- Inspect a real `.xlsx` file (rename to `.zip`, open with any archive tool) to see what XML Excel actually produces for the feature you are implementing
- Ask a clarifying question before making large speculative changes to the XML serialisation layer

## Tech Stack

- **Language**: Python 3.7+
- **XML**: `xml.etree.ElementTree` (stdlib)
- **Archive**: `zipfile` (stdlib)
- **Encryption**: `pycryptodome` (AES)
- **OLE detection**: `olefile`
- **Build**: `setuptools` / `build`
- **Tests**: `pytest`
- **Distribution**: PyPI as `aspose-cells-foss`; imported as `aspose_cells`
