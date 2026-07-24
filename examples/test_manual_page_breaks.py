"""
Test cases for manual page break support and XLSX persistence.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from aspose.cells_foss import Workbook


class TestManualPageBreaks(unittest.TestCase):
    """Validates manual row/column page break settings and worksheet XML output."""

    def setUp(self):
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputfiles')
        os.makedirs(self.test_dir, exist_ok=True)

    def test_manual_page_break_api_operations(self):
        wb = Workbook()
        ws = wb.worksheets[0]

        # Aspose/Excel-like object model collection API
        ws.horizontal_page_breaks.Add("A6")  # row index 5 (0-based)
        ws.vertical_page_breaks.Add(2)
        ws.vertical_page_breaks.Add("D")
        self.assertEqual(ws.horizontal_page_breaks.to_list(), [5])
        self.assertEqual(ws.vertical_page_breaks.to_list(), [2, 3])
        self.assertEqual(ws.horizontal_page_breaks.Count, 1)
        self.assertEqual(ws.vertical_page_breaks.Count, 2)

        ws.vertical_page_breaks.Remove("C")
        self.assertEqual(ws.vertical_page_breaks.to_list(), [3])

        # Backward-compatible Cells helper API still works
        ws.cells.SetHorizontalPageBreak(10)
        ws.cells.RemoveHorizontalPageBreak(10)

        ws.horizontal_page_breaks.Clear()
        ws.vertical_page_breaks.Clear()
        self.assertEqual(ws.horizontal_page_breaks.to_list(), [])
        self.assertEqual(ws.vertical_page_breaks.to_list(), [])


if __name__ == '__main__':
    unittest.main()