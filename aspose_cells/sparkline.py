"""
Aspose.Cells for Python - Sparkline Module

Excel sparklines (ECMA-376 §18.3.1 extension / MS-XLSX x14 namespace).
Sparklines are mini-charts embedded directly inside the worksheet XML as
<extLst> extension elements — they do NOT use separate part files.

Classes:
    SparklineType           — line | column | win-loss
    SparklineEmptyCells     — how to render empty cells (gap / zero / connected)
    Sparkline               — one sparkline (data range + display cell)
    SparklineGroup          — a group of sparklines sharing the same style
    SparklineGroupCollection — ws.sparkline_groups collection
"""

from enum import IntEnum
import re


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class SparklineType(IntEnum):
    LINE     = 0   # default; attribute omitted in XML
    COLUMN   = 1   # type="column"
    WIN_LOSS = 2   # type="win-loss"


class SparklineEmptyCells(IntEnum):
    ZERO      = 0  # displayEmptyCellsAs="zero"
    GAP       = 1  # displayEmptyCellsAs="gap"  (default)
    CONNECTED = 2  # displayEmptyCellsAs="connected"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_a1_cell(ref: str):
    """Parse 'F2' → (col_letter, row_num). Ignores $ signs."""
    ref = ref.strip().upper().replace('$', '')
    m = re.match(r'^([A-Z]{1,3})(\d+)$', ref)
    if not m:
        raise ValueError(f"Cannot parse cell reference: {ref!r}")
    return m.group(1), int(m.group(2))


def _parse_a1_range(ref: str):
    """
    Parse 'A1:D4' → list of (col_letter, row_num) for each cell in the range.
    Returns a flat list in row-major order.
    """
    ref = ref.strip().upper().replace('$', '')
    if ':' in ref:
        a, b = ref.split(':', 1)
    else:
        # Single cell
        col_l, row_n = _parse_a1_cell(ref)
        return [(col_l, row_n)]

    def col_to_idx(col_str):
        idx = 0
        for ch in col_str:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx

    def idx_to_col(n):
        letters = []
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters.append(chr(ord('A') + rem))
        return ''.join(reversed(letters))

    col_a, row_a = _parse_a1_cell(a)
    col_b, row_b = _parse_a1_cell(b)
    c1, c2 = col_to_idx(col_a), col_to_idx(col_b)
    r1, r2 = min(row_a, row_b), max(row_a, row_b)
    c1, c2 = min(c1, c2), max(c1, c2)

    cells = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cells.append((idx_to_col(c), r))
    return cells


def _strip_sheet_ref(data_range: str) -> tuple:
    """
    Split 'Sheet1!A2:D4' into ('Sheet1', 'A2:D4').
    Also handles bare ranges like 'A2:D4' → (None, 'A2:D4').
    """
    if '!' in data_range:
        sheet, rng = data_range.split('!', 1)
        sheet = sheet.strip("'")
        return sheet, rng
    return None, data_range


# ---------------------------------------------------------------------------
# Sparkline (individual)
# ---------------------------------------------------------------------------

class Sparkline:
    """One sparkline: a data source range paired with the cell where it appears."""

    def __init__(self, data_range: str, cell_reference: str):
        self.data_range = data_range       # e.g. "Sheet1!A2:D2"
        self.cell_reference = cell_reference  # e.g. "F2"

    def copy(self):
        return Sparkline(self.data_range, self.cell_reference)


# ---------------------------------------------------------------------------
# SparklineGroup
# ---------------------------------------------------------------------------

class SparklineGroup:
    """
    A group of sparklines that share the same visual style.
    Corresponds to <x14:sparklineGroup> in the worksheet XML extension.
    """

    # Default colors match Excel's built-in "dark blue" sparkline theme
    _DEFAULT_COLORS = {
        'color_series':   '376092',
        'color_negative': 'D00000',
        'color_axis':     '000000',
        'color_markers':  'D00000',
        'color_first':    'D00000',
        'color_last':     'D00000',
        'color_high':     'D00000',
        'color_low':      'D00000',
    }

    def __init__(self):
        self.type = SparklineType.LINE
        self.display_empty_cells_as = SparklineEmptyCells.GAP

        # Colors as 6-digit RRGGBB (library convention); full ARGB is FF{color} in XML
        self.color_series   = self._DEFAULT_COLORS['color_series']
        self.color_negative = self._DEFAULT_COLORS['color_negative']
        self.color_axis     = self._DEFAULT_COLORS['color_axis']
        self.color_markers  = self._DEFAULT_COLORS['color_markers']
        self.color_first    = self._DEFAULT_COLORS['color_first']
        self.color_last     = self._DEFAULT_COLORS['color_last']
        self.color_high     = self._DEFAULT_COLORS['color_high']
        self.color_low      = self._DEFAULT_COLORS['color_low']

        self.sparklines: list = []          # list of Sparkline objects
        self._source_group_xml: str = None  # raw <x14:sparklineGroup .../> for round-trip
        self._uid: str = None               # xr2:uid value

    # ------------------------------------------------------------------
    # Sparkline management
    # ------------------------------------------------------------------

    def add_sparkline(self, data_range: str, cell_reference: str) -> Sparkline:
        """Add an individual sparkline entry to this group."""
        sp = Sparkline(data_range, cell_reference)
        self.sparklines.append(sp)
        self._source_group_xml = None  # invalidate round-trip source
        return sp

    # ------------------------------------------------------------------
    # Collection protocol for sparklines list
    # ------------------------------------------------------------------

    @property
    def count(self) -> int:
        return len(self.sparklines)

    def copy(self):
        g = SparklineGroup()
        g.type = self.type
        g.display_empty_cells_as = self.display_empty_cells_as
        g.color_series   = self.color_series
        g.color_negative = self.color_negative
        g.color_axis     = self.color_axis
        g.color_markers  = self.color_markers
        g.color_first    = self.color_first
        g.color_last     = self.color_last
        g.color_high     = self.color_high
        g.color_low      = self.color_low
        g.sparklines     = [s.copy() for s in self.sparklines]
        g._source_group_xml = self._source_group_xml
        g._uid           = self._uid
        return g


# ---------------------------------------------------------------------------
# SparklineGroupCollection
# ---------------------------------------------------------------------------

class SparklineGroupCollection:
    """Collection of SparklineGroup objects (ws.sparkline_groups)."""

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._groups: list = []

    # ------------------------------------------------------------------
    # Public add API
    # ------------------------------------------------------------------

    def add(self, sparkline_type=SparklineType.LINE,
            data_range: str = None,
            is_vertical: bool = False,
            location_range: str = None) -> SparklineGroup:
        """
        Create a SparklineGroup and auto-populate sparklines from ranges.

        Args:
            sparkline_type:  SparklineType.LINE | COLUMN | WIN_LOSS
            data_range:      Source data range, e.g. "Sheet1!A2:D4"
                             Each row (is_vertical=False) or column (True) → one sparkline
            is_vertical:     If True, iterate columns instead of rows for sparklines
            location_range:  Cells where sparklines appear, e.g. "F2:F4"

        Returns:
            The new SparklineGroup.
        """
        group = SparklineGroup()
        group.type = sparkline_type

        if data_range and location_range:
            self._populate_sparklines(group, data_range, is_vertical, location_range)

        self._groups.append(group)
        self._worksheet._sparklines_dirty = True
        self._worksheet._source_sparkline_xml = None
        return group

    def add_group(self, sparkline_type=SparklineType.LINE) -> SparklineGroup:
        """Add an empty SparklineGroup; caller adds sparklines manually."""
        group = SparklineGroup()
        group.type = sparkline_type
        self._groups.append(group)
        self._worksheet._sparklines_dirty = True
        self._worksheet._source_sparkline_xml = None
        return group

    # PascalCase aliases
    def Add(self, sparkline_type=SparklineType.LINE, data_range=None,
            is_vertical=False, location_range=None) -> SparklineGroup:
        return self.add(sparkline_type, data_range, is_vertical, location_range)

    def AddGroup(self, sparkline_type=SparklineType.LINE) -> SparklineGroup:
        return self.add_group(sparkline_type)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _add_loaded(self, group: SparklineGroup):
        """Append a group loaded from file without marking dirty."""
        self._groups.append(group)

    def _populate_sparklines(self, group, data_range, is_vertical, location_range):
        """Build individual Sparkline objects from the given ranges."""
        # Separate sheet prefix from range
        sheet_name, bare_data = _strip_sheet_ref(data_range)
        prefix = f"{sheet_name}!" if sheet_name else ""

        # Parse location cells
        loc_cells = _parse_a1_range(location_range)

        # Parse data range cells
        data_cells = _parse_a1_range(bare_data)
        if not data_cells:
            return

        # Determine row / column boundaries
        rows = sorted(set(r for _, r in data_cells))
        cols = sorted(set(c for c, _ in data_cells))

        if not is_vertical:
            # Each row → one sparkline; columns form the data series
            items = rows
            def make_data_ref(row):
                c1 = cols[0]
                c2 = cols[-1]
                return f"{prefix}{c1}{row}:{c2}{row}"
        else:
            # Each column → one sparkline; rows form the data series
            items = cols
            def make_data_ref(col):
                r1 = rows[0]
                r2 = rows[-1]
                return f"{prefix}{col}{r1}:{col}{r2}"

        for i, item in enumerate(items):
            if i >= len(loc_cells):
                break
            loc_col, loc_row = loc_cells[i]
            cell_ref = f"{loc_col}{loc_row}"
            dr = make_data_ref(item)
            group.sparklines.append(Sparkline(dr, cell_ref))

    # ------------------------------------------------------------------
    # Collection protocol
    # ------------------------------------------------------------------

    @property
    def count(self) -> int:
        return len(self._groups)

    def __len__(self) -> int:
        return len(self._groups)

    def __iter__(self):
        return iter(self._groups)

    def __getitem__(self, index):
        return self._groups[index]

    def copy(self, new_ws) -> 'SparklineGroupCollection':
        sgc = SparklineGroupCollection(new_ws)
        sgc._groups = [g.copy() for g in self._groups]
        return sgc
