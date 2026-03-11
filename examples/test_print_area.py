"""
Integration test for worksheet print area setting.

This test demonstrates:
1. Setting print area using worksheet API.
2. Saving to xlsx.
3. Loading back and verifying via API (no XML comparison).
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from aspose.cells_foss import Workbook


class TestPrintArea(unittest.TestCase):
    """Tests print area API with save/load roundtrip."""

    def setUp(self):
        self.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputfiles")
        os.makedirs(self.test_dir, exist_ok=True)

    def test_set_print_area_and_save(self):
        path = os.path.join(self.test_dir, "print_area_demo.xlsx")

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "PrintAreaDemo"

        # Dummy data
        ws.cells["A1"].value = "Product"
        ws.cells["B1"].value = "Q1"
        ws.cells["C1"].value = "Q2"
        ws.cells["A2"].value = "Apple"
        ws.cells["B2"].value = 120
        ws.cells["C2"].value = 132
        ws.cells["A3"].value = "Orange"
        ws.cells["B3"].value = 98
        ws.cells["C3"].value = 105

        ws.set_print_area("A1:C10")
        self.assertEqual(ws.print_area, "A1:C10")

        wb.save(path)
        self.assertTrue(os.path.exists(path))
        self.assertGreater(os.path.getsize(path), 0)

        wb2 = Workbook(path)
        ws2 = wb2.worksheets[0]
        self.assertEqual(ws2.print_area, "A1:C10")

    def test_clear_print_area_and_save(self):
        path = os.path.join(self.test_dir, "print_area_cleared_demo.xlsx")

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells["A1"].value = "Demo"

        ws.SetPrintArea("A1:B5")
        self.assertEqual(ws.print_area, "A1:B5")
        ws.ClearPrintArea()
        self.assertIsNone(ws.print_area)

        wb.save(path)
        self.assertTrue(os.path.exists(path))

        wb2 = Workbook(path)
        ws2 = wb2.worksheets[0]
        self.assertIsNone(ws2.print_area)


if __name__ == "__main__":
    unittest.main()
