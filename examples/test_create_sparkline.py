"""
Test case to create an Excel file with dummy data and add sparklines.

This test demonstrates:
1. Creating a new Excel workbook from scratch
2. Adding dummy data to the worksheet
3. Adding several sparklines (line, column, and win-loss types)
4. Saving the workbook to a file
"""

import os
import sys
import zipfile
import re

# Add parent directory to path to import aspose.cells_foss
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose.cells_foss import Workbook
from aspose.cells_foss import SparklineType, SparklineEmptyCells


def test_create_line_sparklines():
    """
    Create a new Excel file with dummy data and add line sparklines.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/createsparkline", exist_ok=True)
    
    output_path = "tests/outputfiles/createsparkline/test_create_line_sparklines.xlsx"
    
    # Create a new workbook
    print("Creating new workbook for line sparklines...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "LineSparklines"
    
    # Add dummy data for sparklines
    print("\nAdding dummy data...")
    # Add headers
    ws.cells["A1"].value = "Store"
    ws.cells["B1"].value = "January"
    ws.cells["C1"].value = "February"
    ws.cells["D1"].value = "March"
    ws.cells["E1"].value = "April"
    ws.cells["F1"].value = "May"
    ws.cells["G1"].value = "Trend"
    
    # Add data rows
    dummy_data = [
        ["Houston", 4873, 11776, 8355, 9241, 10567],
        ["San Diego", 9575, 7135, 5575, 8234, 7892],
        ["Portland", 12011, 9373, 3386, 6789, 8456],
        ["Seattle", 6543, 8765, 9876, 7654, 8765],
        ["Austin", 7890, 6543, 8765, 5432, 6789],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Add line sparklines
    print("\nAdding line sparklines...")
    
    # Create a sparkline group for line sparklines
    sparkline_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B2:F6",
        is_vertical=False,
        location_range="G2:G6"
    )
    
    # Customize sparkline appearance
    sparkline_group.color_series = "0070C0"  # Blue
    sparkline_group.line_weight = 1.0
    sparkline_group.show_high_point = True
    sparkline_group.show_low_point = True
    sparkline_group.color_high = "00B050"  # Green for high point
    sparkline_group.color_low = "FF0000"   # Red for low point
    
    print(f"Added {sparkline_group.count} line sparklines")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())
        # Check that worksheet file exists
        assert "xl/worksheets/sheet1.xml" in names, "Missing worksheet part"
        
        # Read and verify worksheet XML
        worksheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        
        # Verify sparkline extLst exists
        assert "<extLst>" in worksheet_xml, "Missing extLst in worksheet XML"
        assert "sparklineGroups" in worksheet_xml, "Missing sparklineGroups in worksheet XML"
        
        # Verify sparkline type (line is default, so type attribute may be omitted)
        # Check for sparkline elements
        assert "<x14:sparkline>" in worksheet_xml, "Missing sparkline element"
        assert "<xm:f>" in worksheet_xml, "Missing data range formula"
        assert "<xm:sqref>" in worksheet_xml, "Missing cell reference"
        
        # Verify data ranges
        assert f"{ws.name}!B2:F6" in worksheet_xml or f"{ws.name}!B2:F2" in worksheet_xml, "Missing data range"
        assert "G2" in worksheet_xml, "Missing location cell"
        
        # Verify colors
        assert "0070C0" in worksheet_xml, "Missing series color"
        assert "00B050" in worksheet_xml, "Missing high point color"
        assert "FF0000" in worksheet_xml, "Missing low point color"
    
    print(f"[OK] Successfully saved workbook with {sparkline_group.count} line sparklines")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.sparkline_groups.count} sparkline group(s) in saved file")
    for group_idx, group in enumerate(ws_verify.sparkline_groups):
        print(f"  Group {group_idx}: type={group.type}, count={group.count}")
        for sp_idx, sp in enumerate(group.sparklines):
            print(f"    Sparkline {sp_idx}: data_range='{sp.data_range}', cell='{sp.cell_reference}'")
    
    return output_path


def test_create_column_sparklines():
    """
    Create a new Excel file with dummy data and add column sparklines.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/createsparkline", exist_ok=True)
    
    output_path = "tests/outputfiles/createsparkline/test_create_column_sparklines.xlsx"
    
    # Create a new workbook
    print("\nCreating new workbook for column sparklines...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "ColumnSparklines"
    
    # Add dummy data for column sparklines
    print("\nAdding dummy data...")
    # Add headers
    ws.cells["A1"].value = "Product"
    ws.cells["B1"].value = "Q1"
    ws.cells["C1"].value = "Q2"
    ws.cells["D1"].value = "Q3"
    ws.cells["E1"].value = "Q4"
    ws.cells["F1"].value = "Quarterly"
    
    # Add data rows
    dummy_data = [
        ["Product A", 150, 200, 180, 220],
        ["Product B", 120, 140, 160, 190],
        ["Product C", 180, 170, 150, 200],
        ["Product D", 90, 110, 130, 145],
        ["Product E", 200, 210, 190, 230],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Add column sparklines
    print("\nAdding column sparklines...")
    
    # Create a sparkline group for column sparklines
    sparkline_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.COLUMN,
        data_range=f"{ws.name}!B2:E6",
        is_vertical=False,
        location_range="F2:F6"
    )
    
    # Customize sparkline appearance
    sparkline_group.color_series = "FFC000"  # Orange
    sparkline_group.show_high_point = True
    sparkline_group.show_low_point = True
    sparkline_group.color_high = "00B050"  # Green for high point
    sparkline_group.color_low = "C00000"   # Dark red for low point
    
    print(f"Added {sparkline_group.count} column sparklines")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        worksheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        
        # Verify sparkline type is column
        assert 'type="column"' in worksheet_xml, "Missing column type attribute"
        
        # Verify colors
        assert "FFC000" in worksheet_xml, "Missing series color"
        assert "00B050" in worksheet_xml, "Missing high point color"
        assert "C00000" in worksheet_xml, "Missing low point color"
    
    print(f"[OK] Successfully saved workbook with {sparkline_group.count} column sparklines")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.sparkline_groups.count} sparkline group(s) in saved file")
    
    return output_path


def test_create_win_loss_sparklines():
    """
    Create a new Excel file with dummy data and add win-loss sparklines.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/createsparkline", exist_ok=True)
    
    output_path = "tests/outputfiles/createsparkline/test_create_win_loss_sparklines.xlsx"
    
    # Create a new workbook
    print("\nCreating new workbook for win-loss sparklines...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "WinLossSparklines"
    
    # Add dummy data for win-loss sparklines
    print("\nAdding dummy data...")
    # Add headers
    ws.cells["A1"].value = "Team"
    ws.cells["B1"].value = "Game 1"
    ws.cells["C1"].value = "Game 2"
    ws.cells["D1"].value = "Game 3"
    ws.cells["E1"].value = "Game 4"
    ws.cells["F1"].value = "Game 5"
    ws.cells["G1"].value = "Performance"
    
    # Add data rows (positive = win, negative = loss)
    dummy_data = [
        ["Team A", 1, 1, -1, 1, 1],
        ["Team B", -1, 1, 1, -1, 1],
        ["Team C", 1, 1, 1, 1, -1],
        ["Team D", -1, -1, 1, 1, 1],
        ["Team E", 1, -1, 1, -1, 1],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Add win-loss sparklines
    print("\nAdding win-loss sparklines...")
    
    # Create a sparkline group for win-loss sparklines
    sparkline_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.WIN_LOSS,
        data_range=f"{ws.name}!B2:F6",
        is_vertical=False,
        location_range="G2:G6"
    )
    
    # Customize sparkline appearance
    sparkline_group.color_series = "0070C0"  # Blue
    sparkline_group.color_negative = "FF0000"  # Red for negative values
    
    print(f"Added {sparkline_group.count} win-loss sparklines")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        worksheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        
        # Verify sparkline type is win-loss
        assert 'type="win-loss"' in worksheet_xml, "Missing win-loss type attribute"
        
        # Verify colors
        assert "0070C0" in worksheet_xml, "Missing series color"
        assert "FF0000" in worksheet_xml, "Missing negative color"
    
    print(f"[OK] Successfully saved workbook with {sparkline_group.count} win-loss sparklines")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.sparkline_groups.count} sparkline group(s) in saved file")
    
    return output_path


def test_create_multiple_sparkline_groups():
    """
    Create a new Excel file with multiple sparkline groups of different types.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/createsparkline", exist_ok=True)
    
    output_path = "tests/outputfiles/createsparkline/test_create_multiple_sparkline_groups.xlsx"
    
    # Create a new workbook
    print("\nCreating new workbook for multiple sparkline groups...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "MultipleSparklines"
    
    # Add dummy data for multiple sparkline groups
    print("\nAdding dummy data...")
    # Add headers
    ws.cells["A1"].value = "Region"
    ws.cells["B1"].value = "Jan"
    ws.cells["C1"].value = "Feb"
    ws.cells["D1"].value = "Mar"
    ws.cells["E1"].value = "Apr"
    ws.cells["F1"].value = "May"
    ws.cells["G1"].value = "Jun"
    ws.cells["H1"].value = "Line"
    ws.cells["I1"].value = "Column"
    ws.cells["J1"].value = "WinLoss"
    
    # Add data rows
    dummy_data = [
        ["North", 100, 120, 115, 130, 125, 140],
        ["South", 80, 90, 85, 95, 100, 110],
        ["East", 110, 105, 120, 115, 130, 125],
        ["West", 70, 85, 90, 80, 95, 100],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Add multiple sparkline groups
    print("\nAdding multiple sparkline groups...")
    
    # Group 1: Line sparklines
    line_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B2:G5",
        is_vertical=False,
        location_range="H2:H5"
    )
    line_group.color_series = "0070C0"
    line_group.show_high_point = True
    line_group.show_low_point = True
    print(f"Added line sparkline group with {line_group.count} sparklines")
    
    # Group 2: Column sparklines
    column_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.COLUMN,
        data_range=f"{ws.name}!B2:G5",
        is_vertical=False,
        location_range="I2:I5"
    )
    column_group.color_series = "FFC000"
    column_group.show_high_point = True
    print(f"Added column sparkline group with {column_group.count} sparklines")
    
    # Group 3: Win-loss sparklines (using deviations from average)
    winloss_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.WIN_LOSS,
        data_range=f"{ws.name}!B2:G5",
        is_vertical=False,
        location_range="J2:J5"
    )
    winloss_group.color_series = "00B050"
    winloss_group.color_negative = "FF0000"
    print(f"Added win-loss sparkline group with {winloss_group.count} sparklines")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        worksheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        
        # Verify multiple sparkline groups exist
        sparkline_groups_count = worksheet_xml.count("<x14:sparklineGroup")
        assert sparkline_groups_count >= 3, f"Expected at least 3 sparkline groups, found {sparkline_groups_count}"
        
        # Verify different types
        assert 'type="column"' in worksheet_xml, "Missing column type"
        assert 'type="win-loss"' in worksheet_xml, "Missing win-loss type"
        
        # Verify different colors
        assert "0070C0" in worksheet_xml, "Missing blue color"
        assert "FFC000" in worksheet_xml, "Missing orange color"
        assert "00B050" in worksheet_xml, "Missing green color"
        assert "FF0000" in worksheet_xml, "Missing red color"
    
    print(f"[OK] Successfully saved workbook with {ws.sparkline_groups.count} sparkline groups")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.sparkline_groups.count} sparkline group(s) in saved file")
    for group_idx, group in enumerate(ws_verify.sparkline_groups):
        print(f"  Group {group_idx}: type={group.type}, count={group.count}")
    
    return output_path


def test_create_sparkline_with_empty_cells():
    """
    Create a new Excel file with sparklines that handle empty cells.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/createsparkline", exist_ok=True)
    
    output_path = "tests/outputfiles/createsparkline/test_create_sparkline_with_empty_cells.xlsx"
    
    # Create a new workbook
    print("\nCreating new workbook for sparklines with empty cells...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "EmptyCellsSparklines"
    
    # Add dummy data with some empty cells
    print("\nAdding dummy data with empty cells...")
    # Add headers
    ws.cells["A1"].value = "Item"
    ws.cells["B1"].value = "Jan"
    ws.cells["C1"].value = "Feb"
    ws.cells["D1"].value = "Mar"
    ws.cells["E1"].value = "Apr"
    ws.cells["F1"].value = "May"
    ws.cells["G1"].value = "Trend"
    
    # Add data rows with some empty cells
    dummy_data = [
        ["Item 1", 100, 120, None, 140, 150],
        ["Item 2", 80, None, 100, 110, 120],
        ["Item 3", 90, 95, 105, None, 115],
        ["Item 4", 70, 75, 80, 85, None],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data with empty cells")
    
    # Add sparklines with different empty cell handling
    print("\nAdding sparklines with empty cell handling...")
    
    # Group 1: Treat empty cells as gaps
    gap_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B2:F2",
        is_vertical=False,
        location_range="G2"
    )
    gap_group.display_empty_cells_as = SparklineEmptyCells.GAP
    gap_group.color_series = "0070C0"
    print(f"Added sparkline with GAP handling")
    
    # Group 2: Treat empty cells as zero
    zero_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B3:F3",
        is_vertical=False,
        location_range="G3"
    )
    zero_group.display_empty_cells_as = SparklineEmptyCells.ZERO
    zero_group.color_series = "FFC000"
    print(f"Added sparkline with ZERO handling")
    
    # Group 3: Connect across empty cells
    connected_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B4:F4",
        is_vertical=False,
        location_range="G4"
    )
    connected_group.display_empty_cells_as = SparklineEmptyCells.CONNECTED
    connected_group.color_series = "00B050"
    print(f"Added sparkline with CONNECTED handling")
    
    # Group 4: Default (gap) handling
    default_group = ws.sparkline_groups.add(
        sparkline_type=SparklineType.LINE,
        data_range=f"{ws.name}!B5:F5",
        is_vertical=False,
        location_range="G5"
    )
    default_group.color_series = "FF0000"
    print(f"Added sparkline with default (GAP) handling")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        worksheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        
        # Verify empty cell handling attributes
        assert 'displayEmptyCellsAs="gap"' in worksheet_xml, "Missing gap handling"
        assert 'displayEmptyCellsAs="zero"' in worksheet_xml, "Missing zero handling"
        assert 'displayEmptyCellsAs="connected"' in worksheet_xml, "Missing connected handling"
        
        # Verify different colors
        assert "0070C0" in worksheet_xml, "Missing blue color"
        assert "FFC000" in worksheet_xml, "Missing orange color"
        assert "00B050" in worksheet_xml, "Missing green color"
        assert "FF0000" in worksheet_xml, "Missing red color"
    
    print(f"[OK] Successfully saved workbook with {ws.sparkline_groups.count} sparkline groups")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.sparkline_groups.count} sparkline group(s) in saved file")
    for group_idx, group in enumerate(ws_verify.sparkline_groups):
        print(f"  Group {group_idx}: type={group.type}, empty_cells={group.display_empty_cells_as}")
    
    return output_path


def test_all_create_sparklines():
    """
    Run all test cases for creating Excel files with sparklines.
    """
    print("\n" + "="*70)
    print("Test: Create Excel Files with Sparklines")
    print("="*70)
    
    test_create_line_sparklines()
    test_create_column_sparklines()
    test_create_win_loss_sparklines()
    test_create_multiple_sparkline_groups()
    test_create_sparkline_with_empty_cells()
    
    print("\n" + "="*70)
    print("All tests completed successfully!")
    print("="*70 + "\n")


if __name__ == "__main__":
    test_all_create_sparklines()
