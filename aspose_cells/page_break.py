"""
Aspose.Cells for Python - Page Break Module

Manual page break collections aligned with Aspose.Cells/.NET and Excel object model style.
"""

from .cells import Cells


class HorizontalPageBreakCollection:
    """Collection of manual horizontal page breaks (row breaks)."""

    def __init__(self, worksheet):
        self._worksheet = worksheet

    def _normalize_row(self, row_or_cell):
        if isinstance(row_or_cell, str):
            row, _ = Cells.coordinate_from_string(row_or_cell)
            return int(row) - 1
        if row_or_cell is None or int(row_or_cell) < 0:
            raise ValueError("row must be >= 0")
        return int(row_or_cell)

    def add(self, row_or_cell):
        """
        Adds a manual horizontal page break.

        Args:
            row_or_cell (int or str): 0-based row index or A1 cell reference.
        """
        row = self._normalize_row(row_or_cell)
        self._worksheet._horizontal_page_breaks.add(row)
        return row

    def remove_at(self, index):
        """
        Removes the break at zero-based collection index.
        """
        row = self.to_list()[index]
        self._worksheet._horizontal_page_breaks.discard(row)

    def remove(self, row_or_cell):
        """
        Removes a manual horizontal page break by row/cell.
        """
        row = self._normalize_row(row_or_cell)
        self._worksheet._horizontal_page_breaks.discard(row)

    def clear(self):
        """Clears all manual horizontal page breaks."""
        self._worksheet._horizontal_page_breaks.clear()

    @property
    def count(self):
        return len(self._worksheet._horizontal_page_breaks)

    @property
    def Count(self):
        return self.count

    def to_list(self):
        return sorted(self._worksheet._horizontal_page_breaks)

    def __len__(self):
        return self.count

    def __iter__(self):
        return iter(self.to_list())

    def __getitem__(self, index):
        return self.to_list()[index]

    # Aspose-style aliases
    def Add(self, row_or_cell):
        return self.add(row_or_cell)

    def RemoveAt(self, index):
        breaks = self.to_list()
        row = breaks[index]
        self._worksheet._horizontal_page_breaks.discard(row)

    def Remove(self, row_or_cell):
        return self.remove(row_or_cell)

    def Clear(self):
        return self.clear()


class VerticalPageBreakCollection:
    """Collection of manual vertical page breaks (column breaks)."""

    def __init__(self, worksheet):
        self._worksheet = worksheet

    def _normalize_column(self, column_or_cell):
        if isinstance(column_or_cell, str):
            if any(ch.isdigit() for ch in column_or_cell):
                _, col = Cells.coordinate_from_string(column_or_cell)
                return int(col) - 1
            return int(Cells.column_index_from_string(column_or_cell)) - 1
        if column_or_cell is None or int(column_or_cell) < 0:
            raise ValueError("column must be >= 0")
        return int(column_or_cell)

    def add(self, column_or_cell):
        """
        Adds a manual vertical page break.

        Args:
            column_or_cell (int or str): 0-based column index, column letters, or A1 cell reference.
        """
        col = self._normalize_column(column_or_cell)
        self._worksheet._vertical_page_breaks.add(col)
        return col

    def remove(self, column_or_cell):
        """
        Removes a manual vertical page break by column/cell.
        """
        col = self._normalize_column(column_or_cell)
        self._worksheet._vertical_page_breaks.discard(col)

    def clear(self):
        """Clears all manual vertical page breaks."""
        self._worksheet._vertical_page_breaks.clear()

    @property
    def count(self):
        return len(self._worksheet._vertical_page_breaks)

    @property
    def Count(self):
        return self.count

    def to_list(self):
        return sorted(self._worksheet._vertical_page_breaks)

    def __len__(self):
        return self.count

    def __iter__(self):
        return iter(self.to_list())

    def __getitem__(self, index):
        return self.to_list()[index]

    # Aspose-style aliases
    def Add(self, column_or_cell):
        return self.add(column_or_cell)

    def RemoveAt(self, index):
        breaks = self.to_list()
        col = breaks[index]
        self._worksheet._vertical_page_breaks.discard(col)

    def Remove(self, column_or_cell):
        return self.remove(column_or_cell)

    def Clear(self):
        return self.clear()
