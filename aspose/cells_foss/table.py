"""
Aspose.Cells for Python - Table (ListObject) Module

Provides Excel structured-table (ECMA-376 §18.5 <table>) classes:
    TableStyleInfo  — visual style settings
    TableColumn     — per-column settings
    Table           — a single table (ListObject in .NET / VBA)
    TableCollection — ws.tables collection
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_index_to_letter(col_idx: int) -> str:
    """Convert 0-based column index to A1 column letter(s). 0→'A', 25→'Z', 26→'AA'."""
    letters = []
    n = col_idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord('A') + rem))
    return ''.join(reversed(letters))


def _parse_cell_range(ref: str):
    """
    Parse an A1 range string like 'A1:C5' into
    (start_row, start_col, end_row, end_col) as 0-based indices.
    Returns None if parsing fails.
    """
    import re
    ref = ref.strip().upper()
    m = re.match(
        r'^\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$',
        ref,
    )
    if not m:
        return None
    def col_to_idx(col_str):
        idx = 0
        for ch in col_str:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1
    start_col = col_to_idx(m.group(1))
    start_row = int(m.group(2)) - 1
    end_col   = col_to_idx(m.group(3))
    end_row   = int(m.group(4)) - 1
    return start_row, start_col, end_row, end_col


# ---------------------------------------------------------------------------
# TableStyleInfo
# ---------------------------------------------------------------------------

class TableStyleInfo:
    """Visual style settings for an Excel table."""

    def __init__(self):
        self.name = "TableStyleMedium2"
        self.show_first_column = False
        self.show_last_column = False
        self.show_row_stripes = True
        self.show_column_stripes = False

    def copy(self):
        si = TableStyleInfo()
        si.name = self.name
        si.show_first_column = self.show_first_column
        si.show_last_column = self.show_last_column
        si.show_row_stripes = self.show_row_stripes
        si.show_column_stripes = self.show_column_stripes
        return si


# ---------------------------------------------------------------------------
# TableColumn
# ---------------------------------------------------------------------------

class TableColumn:
    """Settings for a single table column."""

    def __init__(self, col_id: int, name: str):
        self.id = col_id                        # 1-based, required by ECMA-376
        self.name = name                        # header cell text
        self.totals_row_function = "none"       # sum|count|average|max|min|stdDev|var|countNums|none
        self.totals_row_label = ""              # label for totals cell
        self.totals_row_formula = ""            # custom formula
        self._source_column_xml = None          # raw <tableColumn .../> for round-trip

    def copy(self):
        tc = TableColumn(self.id, self.name)
        tc.totals_row_function = self.totals_row_function
        tc.totals_row_label = self.totals_row_label
        tc.totals_row_formula = self.totals_row_formula
        tc._source_column_xml = self._source_column_xml
        return tc


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------

class Table:
    """
    Represents an Excel structured table (ECMA-376 §18.5).
    Equivalent to ListObject in Aspose.Cells for .NET / Excel VBA.
    """

    def __init__(self, name: str, ref: str, has_headers: bool = True):
        self.name = name                        # internal name (no spaces)
        self.display_name = name                # displayName attribute
        self.ref = ref                          # A1 range, e.g. "A1:C5"
        self.has_headers = has_headers
        self.show_totals_row = False
        self.show_auto_filter = True
        self.table_style_info = TableStyleInfo()
        self.columns: list = []                 # list of TableColumn
        self._id = 1                            # <table id=""> globally unique per workbook
        self._source_table_xml = None           # raw bytes for round-trip
        self._source_part_path = None           # e.g. "xl/tables/table1.xml"

    def copy(self):
        t = Table(self.name, self.ref, self.has_headers)
        t.display_name = self.display_name
        t.show_totals_row = self.show_totals_row
        t.show_auto_filter = self.show_auto_filter
        t.table_style_info = self.table_style_info.copy()
        t.columns = [c.copy() for c in self.columns]
        t._id = self._id
        t._source_table_xml = self._source_table_xml
        t._source_part_path = self._source_part_path
        return t


# ---------------------------------------------------------------------------
# TableCollection
# ---------------------------------------------------------------------------

class TableCollection:
    """Collection of Table objects belonging to a worksheet (ws.tables)."""

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._tables: list = []

    # ------------------------------------------------------------------
    # Public add API
    # ------------------------------------------------------------------

    def add(self, start_row: int, start_col: int,
            end_row: int, end_col: int,
            has_headers: bool = True,
            name: str = None) -> Table:
        """
        Create a new table from 0-based row/column indices and add it to the worksheet.

        Args:
            start_row: 0-based row index of the top-left cell (including header).
            start_col: 0-based column index of the top-left cell.
            end_row:   0-based row index of the bottom-right cell (including totals if any).
            end_col:   0-based column index of the bottom-right cell.
            has_headers: If True, the first row contains column headers.
            name: Optional table name. Auto-generated as "Table1", "Table2", … if not given.
        Returns:
            The new Table instance.
        """
        if name is None:
            name = self._auto_name()
        # Build A1 range string
        start_letter = _col_index_to_letter(start_col)
        end_letter = _col_index_to_letter(end_col)
        ref = f"{start_letter}{start_row + 1}:{end_letter}{end_row + 1}"

        table = Table(name, ref, has_headers)
        table._id = len(self._tables) + 1  # provisional; re-assigned globally at save time

        # Build columns
        col_count = end_col - start_col + 1
        for i in range(col_count):
            col_name = self._read_header_cell(start_row, start_col + i, has_headers, i)
            table.columns.append(TableColumn(i + 1, col_name))

            # Excel expects the header row cells to contain the table column names.
            # If a table is created over an empty range, materialize default headers
            # into the worksheet so the sheet XML matches the table definition.
            if has_headers:
                self._ensure_header_cell_value(start_row, start_col + i, col_name)

        self._tables.append(table)
        self._worksheet._tables_dirty = True
        return table

    def add_with_range(self, cell_range: str,
                       name: str = None,
                       has_headers: bool = True) -> Table:
        """
        Create a new table from an A1 range string like 'A1:D10'.
        """
        parsed = _parse_cell_range(cell_range)
        if parsed is None:
            raise ValueError(f"Cannot parse cell range: {cell_range!r}")
        start_row, start_col, end_row, end_col = parsed
        return self.add(start_row, start_col, end_row, end_col, has_headers, name)

    # PascalCase aliases
    def Add(self, start_row: int, start_col: int,
            end_row: int, end_col: int,
            has_headers: bool = True,
            name: str = None) -> Table:
        return self.add(start_row, start_col, end_row, end_col, has_headers, name)

    def AddWithRange(self, cell_range: str,
                     name: str = None,
                     has_headers: bool = True) -> Table:
        return self.add_with_range(cell_range, name, has_headers)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _add_loaded(self, table: Table):
        """Append a table loaded from file without marking dirty."""
        self._tables.append(table)

    def _auto_name(self) -> str:
        """Generate the next available table name like 'Table1', 'Table2', …"""
        existing = {t.name.lower() for t in self._tables}
        i = 1
        while f"table{i}" in existing:
            i += 1
        return f"Table{i}"

    def _read_header_cell(self, row: int, col: int, has_headers: bool, col_idx: int) -> str:
        """Read the header cell value from the worksheet, or generate a default name."""
        if has_headers:
            col_letter = _col_index_to_letter(col)
            cell_ref = f"{col_letter}{row + 1}"
            try:
                val = self._worksheet.cells[cell_ref].value
                if val is not None and str(val).strip():
                    return str(val).strip()
            except Exception:
                pass
        return f"Column{col_idx + 1}"

    def _ensure_header_cell_value(self, row: int, col: int, header_name: str) -> None:
        """Write the generated header name back to the worksheet if the cell is blank."""
        col_letter = _col_index_to_letter(col)
        cell_ref = f"{col_letter}{row + 1}"
        try:
            cell = self._worksheet.cells[cell_ref]
            if cell.value is None or not str(cell.value).strip():
                cell.value = header_name
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    @property
    def count(self) -> int:
        return len(self._tables)

    def __len__(self) -> int:
        return len(self._tables)

    def __iter__(self):
        return iter(self._tables)

    def __getitem__(self, index):
        return self._tables[index]

    def copy(self, new_ws) -> 'TableCollection':
        tc = TableCollection(new_ws)
        tc._tables = [t.copy() for t in self._tables]
        return tc
