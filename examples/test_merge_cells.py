"""
Integration test for merged cell feature.

Flow:
1. Create workbook with dummy data.
2. Merge ranges and save to first file.
3. Reload first file, unmerge ranges.
4. Save to second file and verify merged ranges are removed.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from aspose.cells_foss import Workbook


class TestMergeCells(unittest.TestCase):
    """Tests merge -> save -> load -> unmerge -> save workflow."""

    def setUp(self):
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputfiles")
        os.makedirs(self.test_dir, exist_ok=True)

    def test_merge_then_unmerge_roundtrip(self):
        merged_path = os.path.join(self.test_dir, "example_merge_cells_merged.xlsx")
        unmerged_path = os.path.join(self.test_dir, "example_merge_cells_unmerged.xlsx")

        # Create workbook with dummy data
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "MergeDemo"
        ws.cells["A1"].value = "Quarterly Report"
        ws.cells["A3"].value = "Region"
        ws.cells["B3"].value = "Q1"
        ws.cells["C3"].value = "Q2"
        ws.cells["D3"].value = "Q3"
        ws.cells["E3"].value = "Q4"
        ws.cells["A4"].value = "North"
        ws.cells["A5"].value = "South"
        ws.cells["B4"].value = 120
        ws.cells["C4"].value = 135
        ws.cells["D4"].value = 142
        ws.cells["E4"].value = 150
        ws.cells["B5"].value = 98
        ws.cells["C5"].value = 110
        ws.cells["D5"].value = 119
        ws.cells["E5"].value = 130

        # Merge title row and one data label range
        ws.cells.merge(0, 0, 1, 5)  # A1:E1
        ws.cells.Merge(5, 0, 1, 2)  # A6:B6 (alias API)
        ws.cells["A6"].value = "Merged note"

        wb.save(merged_path)
        self.assertTrue(os.path.exists(merged_path))
        self.assertEqual(ws.merged_cells, ["A1:E1", "A6:B6"])

        # Reload and unmerge
        wb2 = Workbook(merged_path)
        ws2 = wb2.worksheets[0]
        self.assertEqual(ws2.merged_cells, ["A1:E1", "A6:B6"])
        ws2.cells.unmerge(0, 0, 1, 5)
        ws2.cells.UnMerge(5, 0, 1, 2)  # alias API
        self.assertEqual(ws2.merged_cells, [])

        wb2.save(unmerged_path)
        self.assertTrue(os.path.exists(unmerged_path))


if __name__ == "__main__":
    unittest.main()
