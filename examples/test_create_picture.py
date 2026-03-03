"""
Test creating an Excel file with dummy data and adding a picture.
"""

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from aspose_cells import Workbook


class TestCreatePicture(unittest.TestCase):
    """Test creating an Excel file with dummy data and a picture."""

    def test_create_excel_with_picture(self):
        """
        1. Create an Excel file with some dummy data
        2. Add a picture to the file
        3. Save the file
        """
        # Setup paths
        base_dir = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        image_path = base_dir / "input" / "picture" / "R.jpg"
        output_dir = base_dir / "tests" / "outputfiles" / "picture"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "test_create_picture.xlsx"

        # Skip test if image file doesn't exist
        if not image_path.exists():
            self.skipTest(f"Image file not found: {image_path}")

        # Step 1: Create a new workbook and add dummy data
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "Data with Picture"

        # Add some dummy data to cells
        headers = ["Name", "Age", "Department", "Salary"]
        data = [
            ["John Doe", 32, "Engineering", 75000],
            ["Jane Smith", 28, "Marketing", 65000],
            ["Bob Johnson", 45, "Finance", 85000],
            ["Alice Williams", 38, "Engineering", 80000],
            ["Charlie Brown", 29, "HR", 60000],
        ]

        # Write headers
        for col, header in enumerate(headers, start=1):
            ws.cells.set_cell(1, col, header)

        # Write data rows
        for row_idx, row_data in enumerate(data, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cells.set_cell(row_idx, col_idx, value)

        # Verify data was added
        self.assertEqual(ws.cells.get_cell(1, 1).value, "Name")
        self.assertEqual(ws.cells.get_cell(2, 1).value, "John Doe")
        self.assertEqual(ws.cells.get_cell(3, 4).value, 65000)

        # Step 2: Add a picture to the worksheet
        # Position the picture starting at row 8, column 1, ending at row 15, column 4
        picture_index = ws.pictures.add(
            str(image_path),
            upper_left_row=8,
            upper_left_column=1,
            lower_right_row=15,
            lower_right_column=4
        )

        # Verify picture was added
        self.assertEqual(picture_index, 0, "First picture should have index 0")
        self.assertEqual(ws.pictures.count, 1, "Worksheet should have exactly 1 picture")
        
        pic = ws.pictures[0]
        self.assertIn(pic.image_extension.lower(), {"jpg", "jpeg"}, "Picture should be a JPEG file")
        self.assertIsNotNone(pic.image_bytes, "Picture should have image data")
        self.assertGreater(len(pic.name), 0, "Picture name should not be empty")

        # Step 3: Save the workbook
        wb.save(str(output_path))
        self.assertTrue(output_path.exists(), f"Output file should exist: {output_path}")

        # Verify the saved file can be reloaded
        wb_loaded = Workbook(str(output_path))
        ws_loaded = wb_loaded.worksheets[0]

        # Verify data was preserved
        self.assertEqual(ws_loaded.cells.get_cell(1, 1).value, "Name")
        self.assertEqual(ws_loaded.cells.get_cell(2, 1).value, "John Doe")
        self.assertEqual(ws_loaded.cells.get_cell(3, 4).value, 65000)

        # Verify picture was preserved
        self.assertEqual(ws_loaded.pictures.count, 1, "Reloaded worksheet should have 1 picture")
        pic_loaded = ws_loaded.pictures[0]
        self.assertIn(pic_loaded.image_extension.lower(), {"jpg", "jpeg"})
        self.assertIsNotNone(pic_loaded.image_bytes)

        print(f"\nSuccessfully created Excel file with picture at: {output_path}")
        print(f"  - Added {len(data)} rows of dummy data")
        print(f"  - Added 1 picture: {pic.name}")


if __name__ == "__main__":
    unittest.main()
