# Aspose.Cells FOSS - Complete Reference

**Aspose.Cells-FOSS-for-Python** is a pure-Python library for creating, reading, and modifying Excel `.xlsx` files without Microsoft Excel. Its public API mirrors [Aspose.Cells for .NET](https://reference.aspose.com/cells/net/), making it straightforward to port .NET code to Python.

## Installation

- Install via `pip install aspose-cells-foss`
- Import as `from aspose.cells_foss import Workbook`
- Requires Python 3.7+; runtime dependencies are `pycryptodome` and `olefile`

## Core Usage

```python
from aspose.cells_foss import Workbook

wb = Workbook()                      # new workbook
wb = Workbook("input.xlsx")          # open existing file
wb = Workbook("protected.xlsx", password="secret")  # open encrypted file

ws = wb.worksheets[0]                # first worksheet (0-based)
ws = wb.worksheets.add("Sheet2")     # add a named worksheet

ws.cells["A1"].put_value("Hello")    # write by A1 reference
ws.cells["A1"].value                 # read value
ws.cells["A2"].formula = "=SUM(B1:B5)"

wb.save("output.xlsx")               # save as XLSX
wb.save("output.csv")                # auto-detects format from extension
wb.save("output.xlsx", password="secret")  # save with AES encryption
```

## Naming Conventions

Every public method has **two names**: a snake_case primary form and a PascalCase alias for .NET-style compatibility.

```python
ws.horizontal_page_breaks.add(19)   # Python style
ws.horizontal_page_breaks.Add(19)   # .NET / Aspose style — identical behaviour
```

## Color Strings

Colors are always **6-digit RRGGBB hex strings without a `#` prefix**.

```python
style.font.color = "FF0000"    # red — correct
style.font.color = "#FF0000"   # wrong — breaks XML serialisation
```

## Cell Coordinates

Row and column indices are **0-based** everywhere in the API.
A1-style string references (`"A1"`, `"B3"`) are also accepted at the cell-access boundary.

## Styling

```python
cell = ws.cells["A1"]
style = cell.get_style()
style.font.is_bold = True
style.font.size = 14
style.font.color = "FF0000"
style.horizontal_alignment = "center"
style.number_format = "#,##0.00"
style.borders["bottom"].line_style = "thin"
cell.apply_style(style)
```

## Save Formats

`SaveFormat` enum controls the output format when the extension is ambiguous:

- `SaveFormat.XLSX` — Excel 2007+ (default)
- `SaveFormat.CSV` — comma-separated values
- `SaveFormat.TSV` — tab-separated values
- `SaveFormat.MARKDOWN` — Markdown table
- `SaveFormat.JSON` — JSON array

```python
from aspose.cells_foss import Workbook, SaveFormat
wb.save("out.csv", SaveFormat.CSV)
```

## Charts

16 chart types supported. Access via `worksheet.charts`.

```python
from aspose.cells_foss import Workbook, ChartType

ws = wb.worksheets[0]
chart = ws.charts.add_line(upper_left_row=0, upper_left_col=4,
                            lower_right_row=20, lower_right_col=12)
chart.title = "Monthly Sales"
chart.n_series.add("B2:B7", category_data="A2:A7", name="Sales")

# Other add methods: add_bar, add_pie, add_area, add_scatter,
# add_waterfall, add_combo, add_stock, add_surface, add_radar,
# add_treemap, add_sunburst, add_histogram, add_funnel,
# add_box_whisker, add_map
```

## Pictures

Embed images anchored between two cells.

```python
ws.pictures.add("logo.png",
    upper_left_row=1, upper_left_column=1,
    lower_right_row=8, lower_right_column=5)

pic = ws.pictures[0]
pic.hyperlink_url = "https://example.com"  # optional click hyperlink
```

## Drawing Shapes

30+ preset shapes via `MsoDrawingType` enum.

```python
from aspose.cells_foss import MsoDrawingType, FillType, TextAlignmentType, TextAnchorType

shape = ws.shapes.add(MsoDrawingType.ROUNDED_RECTANGLE, 1, 1, 5, 5)
shape.text = "Hello"
shape.fill.fore_color = "90EE90"
shape.font.bold = True
shape.text_horizontal_alignment = TextAlignmentType.CENTER
shape.text_vertical_alignment = TextAnchorType.MIDDLE

textbox = ws.shapes.add_text_box(7, 1, 11, 8)
textbox.text = "Notes"
```

Available shape types include: `RECTANGLE`, `ROUNDED_RECTANGLE`, `OVAL`, `DIAMOND`, `TRIANGLE`, `RIGHT_TRIANGLE`, `PARALLELOGRAM`, `TRAPEZOID`, `HEXAGON`, `OCTAGON`, `CROSS`, `STAR_4/5/6/7/8`, `RIGHT_ARROW`, `LEFT_ARROW`, `UP_ARROW`, `DOWN_ARROW`, `TEXT_BOX`, `CALLOUT`, `PENTAGON`, `CLOUD`, `HEART`, `LIGHTNING_BOLT`, `SMILEY_FACE`, `LEFT_RIGHT_ARROW`, `UP_DOWN_ARROW`, `CUBE`, `BEVEL`.

## Sparklines

Mini-charts embedded inside cells. Three types: `LINE`, `COLUMN`, `WIN_LOSS`.

```python
from aspose.cells_foss import SparklineType

group = ws.sparkline_groups.add(
    sparkline_type=SparklineType.LINE,
    data_range="Sheet1!B2:F6",   # source data
    is_vertical=False,
    location_range="G2:G6"       # cells where sparklines appear
)
group.color_series = "0070C0"
group.show_high_point = True
group.color_high = "00B050"
group.color_low = "FF0000"
```

## Excel Tables (ListObject)

Structured tables with auto-filter and named columns.

```python
table = ws.tables.add(start_row=0, start_col=0,
                       end_row=9, end_col=3,
                       has_headers=True, name="SalesTable")
table.table_style_info.name = "TableStyleMedium9"
table.table_style_info.show_row_stripes = True

# Alternative: create from A1-range string
table = ws.tables.add_with_range("A1:D10", name="SalesTable")
```

## Data Validation

```python
from aspose.cells_foss import DataValidationType

v = ws.data_validations.add("A1:A10")
v.type = DataValidationType.LIST
v.formula1 = '"Option1,Option2,Option3"'
```

## Conditional Formatting

```python
cf = ws.conditional_formatting.add("A1:C10")
rule = cf.add_rule()
rule.type = "cellValue"
rule.operator = "greaterThan"
rule.formula1 = "100"
rule.style.font.color = "FF0000"
```

## Hyperlinks

```python
ws.hyperlinks.add("A1", "https://example.com")         # URL
ws.hyperlinks.add("A2", "mailto:info@example.com")     # email
ws.hyperlinks.add("A3", "Sheet2!B5")                   # internal ref
```

## Comments

```python
comment = ws.cells["A1"].add_comment("Author", "This is a note.")
```

## Manual Page Breaks

```python
ws.horizontal_page_breaks.add(19)   # break before row 20 (0-based)
ws.vertical_page_breaks.add(3)      # break before column D (0-based)
ws.horizontal_page_breaks.remove(19)
ws.horizontal_page_breaks.clear()
```

## Merge Cells

```python
ws.cells.merge(0, 0, 1, 3)          # merge 1 row × 3 cols from A1
ws.cells.unmerge(0, 0, 1, 3)
```

## Print Area

```python
ws.page_setup.print_area = "A1:H40"
```

## Auto-Filter

```python
ws.auto_filter.range = "A1:E1"
```

## Workbook & Worksheet Protection

```python
wb.settings.protect(password="pw")                  # protect workbook structure
ws.protect(password="pw")                           # protect worksheet cells
ws.unprotect(password="pw")
```

## Encryption

```python
wb.save("secure.xlsx", password="mypassword")       # encrypt on save
wb2 = Workbook("secure.xlsx", password="mypassword") # decrypt on open
```

## Document Properties

```python
wb.document_properties.title = "My Report"
wb.document_properties.author = "Jane Smith"
wb.document_properties.subject = "Q4 Results"
```

## Formula Evaluator

A lightweight evaluator handles basic formulas at read time for cells without cached values. Supported functions: `CONCATENATE`, `CONCAT`, `TEXT`, `IF`, `AND`, `OR`, `NOT`, `LEN`, `TRIM`, `UPPER`, `LOWER`. Cell references and defined names are also resolved.

## Key Classes

| Class | Description |
|---|---|
| `Workbook` | Root object; manages worksheets, properties, and I/O |
| `Worksheet` | One sheet; exposes `cells`, `charts`, `pictures`, `shapes`, `sparkline_groups`, `tables`, `page_setup`, etc. |
| `Cell` | Single cell; `value`, `formula`, `get_style()`, `set_style()` |
| `Cells` | Cell collection with A1 and (row, col) access |
| `Style` | Cell formatting: font, fill, borders, number format, alignment |
| `Chart` | Chart object with `n_series`, `title`, `legend_position`, axes |
| `Picture` | Embedded image with anchor coordinates and optional hyperlink |
| `Shape` | Drawing shape with fill, line, text, and font formatting |
| `SparklineGroup` | Group of sparklines sharing a visual style |
| `Table` | Structured table with columns, style, and optional totals row |
| `SaveFormat` | Enum for controlling output format (XLSX, CSV, TSV, MARKDOWN, JSON) |
| `DataValidationType` | Enum: `LIST`, `WHOLE`, `DECIMAL`, `DATE`, `TEXT_LENGTH`, `CUSTOM` |
| `ChartType` | Enum: `LINE`, `BAR`, `PIE`, `AREA`, `SCATTER`, `WATERFALL`, … |
| `MsoDrawingType` | Enum of 30+ preset shape geometries |
| `SparklineType` | Enum: `LINE`, `COLUMN`, `WIN_LOSS` |
