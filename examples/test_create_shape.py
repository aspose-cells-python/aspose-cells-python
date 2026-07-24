"""
Test case to create an Excel file with dummy data and add workflow shapes.

This test demonstrates:
1. Creating a new Excel workbook from scratch
2. Adding dummy data to the worksheet
3. Adding several shapes to show a workflow (rectangles, arrows, diamonds, etc.)
4. Saving the workbook to a file
"""

import os
import sys
import zipfile
import re

# Add parent directory to path to import aspose.cells_foss
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose.cells_foss import Workbook
from aspose.cells_foss import MsoDrawingType, FillType, MsoLineDashStyle, TextAlignmentType, TextAnchorType
from examples.output_path_helper import examples_output_path, ensure_examples_output_dir


def test_create_workflow_with_shapes():
    """
    Create a new Excel file with dummy data and add workflow shapes.
    """
    # Ensure output directory exists
    ensure_examples_output_dir("createshape")
    
    output_path = examples_output_path("createshape", "example_test_create_workflow.xlsx")
    
    # Create a new workbook
    print("Creating new workbook...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "Workflow"
    
    # Add dummy data for the workflow
    print("\nAdding dummy data...")
    # Add headers
    ws.cells["A1"].value = "Step"
    ws.cells["B1"].value = "Task Name"
    ws.cells["C1"].value = "Description"
    ws.cells["D1"].value = "Status"
    ws.cells["E1"].value = "Owner"
    
    # Add data rows
    dummy_data = [
        [1, "Start", "Begin the process", "Completed", "John"],
        [2, "Data Collection", "Gather required information", "In Progress", "Jane"],
        [3, "Analysis", "Analyze the collected data", "Pending", "Bob"],
        [4, "Decision", "Review and make decision", "Pending", "Alice"],
        [5, "Implementation", "Implement the solution", "Not Started", "Charlie"],
        [6, "Testing", "Test the implementation", "Not Started", "David"],
        [7, "Deployment", "Deploy to production", "Not Started", "Eve"],
        [8, "End", "Process complete", "Not Started", "Frank"],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Add workflow shapes
    print("\nAdding workflow shapes...")
    
    # Shape 1: Start (Rounded Rectangle) - Green
    start_shape = ws.shapes.add(
        MsoDrawingType.ROUNDED_RECTANGLE,
        upper_left_row=1,
        upper_left_column=7,
        lower_right_row=4,
        lower_right_column=10
    )
    start_shape.name = "Start"
    start_shape.text = "START"
    start_shape.fill.fill_type = FillType.SOLID
    start_shape.fill.fore_color = "90EE90"  # Light green
    start_shape.line.is_visible = True
    start_shape.line.color = "006400"  # Dark green
    start_shape.line.weight = 12700  # 1 pt
    start_shape.font.name = "Arial"
    start_shape.font.size = 14.0
    start_shape.font.bold = True
    start_shape.font.color = "000000"
    start_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    start_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {start_shape.name} (Rounded Rectangle)")
    
    # Shape 2: Arrow 1 (Right Arrow) - Gray
    arrow1 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=2,
        upper_left_column=10,
        lower_right_row=4,
        lower_right_column=12
    )
    arrow1.name = "Arrow1"
    arrow1.text = "→"
    arrow1.fill.fill_type = FillType.SOLID
    arrow1.fill.fore_color = "D3D3D3"  # Light gray
    arrow1.line.is_visible = True
    arrow1.line.color = "696969"  # Dim gray
    arrow1.line.weight = 9525  # 0.75 pt
    arrow1.font.name = "Arial"
    arrow1.font.size = 16.0
    arrow1.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow1.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow1.name} (Right Arrow)")
    
    # Shape 3: Process 1 (Rectangle) - Blue
    process1 = ws.shapes.add(
        MsoDrawingType.RECTANGLE,
        upper_left_row=1,
        upper_left_column=12,
        lower_right_row=4,
        lower_right_column=15
    )
    process1.name = "Process1"
    process1.text = "Data\nCollection"
    process1.fill.fill_type = FillType.SOLID
    process1.fill.fore_color = "87CEEB"  # Sky blue
    process1.line.is_visible = True
    process1.line.color = "00008B"  # Dark blue
    process1.line.weight = 12700  # 1 pt
    process1.font.name = "Arial"
    process1.font.size = 11.0
    process1.font.bold = True
    process1.text_horizontal_alignment = TextAlignmentType.CENTER
    process1.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {process1.name} (Rectangle)")
    
    # Shape 4: Arrow 2 (Right Arrow) - Gray
    arrow2 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=2,
        upper_left_column=15,
        lower_right_row=4,
        lower_right_column=17
    )
    arrow2.name = "Arrow2"
    arrow2.text = "→"
    arrow2.fill.fill_type = FillType.SOLID
    arrow2.fill.fore_color = "D3D3D3"  # Light gray
    arrow2.line.is_visible = True
    arrow2.line.color = "696969"  # Dim gray
    arrow2.line.weight = 9525  # 0.75 pt
    arrow2.font.name = "Arial"
    arrow2.font.size = 16.0
    arrow2.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow2.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow2.name} (Right Arrow)")
    
    # Shape 5: Decision (Diamond) - Yellow
    decision = ws.shapes.add(
        MsoDrawingType.DIAMOND,
        upper_left_row=1,
        upper_left_column=17,
        lower_right_row=4,
        lower_right_column=20
    )
    decision.name = "Decision"
    decision.text = "Decision?"
    decision.fill.fill_type = FillType.SOLID
    decision.fill.fore_color = "FFD700"  # Gold
    decision.line.is_visible = True
    decision.line.color = "B8860B"  # Dark goldenrod
    decision.line.weight = 12700  # 1 pt
    decision.font.name = "Arial"
    decision.font.size = 11.0
    decision.font.bold = True
    decision.text_horizontal_alignment = TextAlignmentType.CENTER
    decision.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {decision.name} (Diamond)")
    
    # Shape 6: Arrow Down (Down Arrow) - Gray
    arrow_down = ws.shapes.add(
        MsoDrawingType.DOWN_ARROW,
        upper_left_row=4,
        upper_left_column=18,
        lower_right_row=7,
        lower_right_column=20
    )
    arrow_down.name = "ArrowDown"
    arrow_down.text = "↓"
    arrow_down.fill.fill_type = FillType.SOLID
    arrow_down.fill.fore_color = "D3D3D3"  # Light gray
    arrow_down.line.is_visible = True
    arrow_down.line.color = "696969"  # Dim gray
    arrow_down.line.weight = 9525  # 0.75 pt
    arrow_down.font.name = "Arial"
    arrow_down.font.size = 16.0
    arrow_down.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow_down.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow_down.name} (Down Arrow)")
    
    # Shape 7: Process 2 (Rectangle) - Orange
    process2 = ws.shapes.add(
        MsoDrawingType.RECTANGLE,
        upper_left_row=7,
        upper_left_column=17,
        lower_right_row=10,
        lower_right_column=20
    )
    process2.name = "Process2"
    process2.text = "Implementation"
    process2.fill.fill_type = FillType.SOLID
    process2.fill.fore_color = "FFA500"  # Orange
    process2.line.is_visible = True
    process2.line.color = "8B4500"  # Saddle brown
    process2.line.weight = 12700  # 1 pt
    process2.font.name = "Arial"
    process2.font.size = 11.0
    process2.font.bold = True
    process2.text_horizontal_alignment = TextAlignmentType.CENTER
    process2.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {process2.name} (Rectangle)")
    
    # Shape 8: Arrow 3 (Right Arrow) - Gray
    arrow3 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=8,
        upper_left_column=20,
        lower_right_row=10,
        lower_right_column=22
    )
    arrow3.name = "Arrow3"
    arrow3.text = "→"
    arrow3.fill.fill_type = FillType.SOLID
    arrow3.fill.fore_color = "D3D3D3"  # Light gray
    arrow3.line.is_visible = True
    arrow3.line.color = "696969"  # Dim gray
    arrow3.line.weight = 9525  # 0.75 pt
    arrow3.font.name = "Arial"
    arrow3.font.size = 16.0
    arrow3.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow3.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow3.name} (Right Arrow)")
    
    # Shape 9: End (Rounded Rectangle) - Red
    end_shape = ws.shapes.add(
        MsoDrawingType.ROUNDED_RECTANGLE,
        upper_left_row=7,
        upper_left_column=22,
        lower_right_row=10,
        lower_right_column=25
    )
    end_shape.name = "End"
    end_shape.text = "END"
    end_shape.fill.fill_type = FillType.SOLID
    end_shape.fill.fore_color = "FF6347"  # Tomato red
    end_shape.line.is_visible = True
    end_shape.line.color = "8B0000"  # Dark red
    end_shape.line.weight = 12700  # 1 pt
    end_shape.font.name = "Arial"
    end_shape.font.size = 14.0
    end_shape.font.bold = True
    end_shape.font.color = "000000"
    end_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    end_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {end_shape.name} (Rounded Rectangle)")
    
    # Shape 10: Text Box for notes
    notes_box = ws.shapes.add_text_box(
        upper_left_row=12,
        upper_left_column=7,
        lower_right_row=17,
        lower_right_column=15
    )
    notes_box.name = "Notes"
    notes_box.text = "Workflow Notes:\n\n1. Start process\n2. Collect data\n3. Make decision\n4. Implement solution\n5. End process"
    notes_box.font.name = "Calibri"
    notes_box.font.size = 10.0
    notes_box.font.color = "000000"
    notes_box.text_horizontal_alignment = TextAlignmentType.LEFT
    notes_box.text_vertical_alignment = TextAnchorType.TOP
    print(f"  Added: {notes_box.name} (Text Box)")
    
    print(f"\nTotal shapes added: {len(ws.shapes)}")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())
        # Check that drawing file exists
        assert "xl/drawings/drawing1.xml" in names, "Missing drawing part"
        
        # Read and verify drawing XML
        drawing_xml = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        
        # Verify shape anchors are valid
        anchors = re.findall(
            r'<xdr:from>\s*<xdr:col>(\d+)</xdr:col>.*?</xdr:from>\s*<xdr:to>\s*<xdr:col>(\d+)</xdr:col>',
            drawing_xml,
            re.DOTALL,
        )
        assert all(int(c1) <= int(c2) for c1, c2 in anchors), "Invalid shape anchor: from.col > to.col"
        
        # Verify shape elements exist
        assert "<xdr:sp" in drawing_xml, "Missing shape element in drawing XML"
        
        # Verify different shape types
        assert '<a:prstGeom prst="roundRect">' in drawing_xml, "Missing rounded rectangle geometry"
        assert '<a:prstGeom prst="rect">' in drawing_xml, "Missing rectangle geometry"
        assert '<a:prstGeom prst="diamond">' in drawing_xml, "Missing diamond geometry"
        assert '<a:prstGeom prst="rightArrow">' in drawing_xml, "Missing right arrow geometry"
        assert '<a:prstGeom prst="downArrow">' in drawing_xml, "Missing down arrow geometry"
        
        # Verify text box
        assert 'txBox="1"' in drawing_xml, "Missing txBox attribute for text box"
        
        # Verify fill colors
        assert '<a:srgbClr val="90EE90"/>' in drawing_xml, "Missing green fill"
        assert '<a:srgbClr val="87CEEB"/>' in drawing_xml, "Missing blue fill"
        assert '<a:srgbClr val="FFD700"/>' in drawing_xml, "Missing gold fill"
        assert '<a:srgbClr val="FFA500"/>' in drawing_xml, "Missing orange fill"
        assert '<a:srgbClr val="FF6347"/>' in drawing_xml, "Missing red fill"
        
        # Verify text content
        assert "START" in drawing_xml, "Missing START text"
        # Newlines are converted to separate paragraph elements, so check for both parts
        assert "Data" in drawing_xml and "Collection" in drawing_xml, "Missing process text"
        assert "Decision?" in drawing_xml, "Missing decision text"
        assert "Implementation" in drawing_xml, "Missing implementation text"
        assert "END" in drawing_xml, "Missing END text"
        assert "Workflow Notes" in drawing_xml, "Missing notes text"
        
        # Verify font properties
        assert '<a:latin typeface="Arial"/>' in drawing_xml, "Missing Arial font"
        assert 'b="1"' in drawing_xml, "Missing bold font property"
    
    print(f"[OK] Successfully saved workbook with {len(ws.shapes)} shapes")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {len(ws_verify.shapes)} shapes in saved file")
    for i, shape in enumerate(ws_verify.shapes):
        print(f"  Shape {i}: type={shape.drawing_type}, name='{shape.name}', text='{shape.text.replace(chr(10), ' ')}'")
    
    return output_path


def test_create_approval_workflow():
    """
    Create a new Excel file with an approval workflow using shapes.
    """
    # Ensure output directory exists
    ensure_examples_output_dir("createshape")
    
    output_path = examples_output_path("createshape", "example_test_create_approval_workflow.xlsx")
    
    # Create a new workbook
    print("\nCreating new workbook for approval workflow...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "ApprovalWorkflow"
    
    # Add dummy data for the approval workflow
    print("\nAdding approval workflow data...")
    # Add headers
    ws.cells["A1"].value = "Request ID"
    ws.cells["B1"].value = "Requester"
    ws.cells["C1"].value = "Type"
    ws.cells["D1"].value = "Amount"
    ws.cells["E1"].value = "Status"
    
    # Add data rows
    approval_data = [
        ["REQ001", "Alice", "Expense", 1500.00, "Pending Manager"],
        ["REQ002", "Bob", "Purchase", 5000.00, "Pending Finance"],
        ["REQ003", "Charlie", "Expense", 250.00, "Approved"],
        ["REQ004", "David", "Purchase", 12000.00, "Pending Director"],
        ["REQ005", "Eve", "Expense", 800.00, "Rejected"],
    ]
    
    for row_idx, row_data in enumerate(approval_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(approval_data)} rows of approval data")
    
    # Add approval workflow shapes
    print("\nAdding approval workflow shapes...")
    
    # Shape 1: Request Submission (Oval) - Purple
    request_shape = ws.shapes.add(
        MsoDrawingType.OVAL,
        upper_left_row=1,
        upper_left_column=7,
        lower_right_row=4,
        lower_right_column=10
    )
    request_shape.name = "Request"
    request_shape.text = "Submit\nRequest"
    request_shape.fill.fill_type = FillType.SOLID
    request_shape.fill.fore_color = "DDA0DD"  # Plum
    request_shape.line.is_visible = True
    request_shape.line.color = "800080"  # Purple
    request_shape.line.weight = 12700  # 1 pt
    request_shape.font.name = "Arial"
    request_shape.font.size = 11.0
    request_shape.font.bold = True
    request_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    request_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {request_shape.name} (Oval)")
    
    # Shape 2: Arrow 1 (Right Arrow)
    arrow1 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=2,
        upper_left_column=10,
        lower_right_row=4,
        lower_right_column=12
    )
    arrow1.name = "Arrow1"
    arrow1.text = "→"
    arrow1.fill.fill_type = FillType.SOLID
    arrow1.fill.fore_color = "D3D3D3"
    arrow1.line.is_visible = True
    arrow1.line.color = "696969"
    arrow1.line.weight = 9525
    arrow1.font.name = "Arial"
    arrow1.font.size = 16.0
    arrow1.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow1.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow1.name} (Right Arrow)")
    
    # Shape 3: Manager Approval (Rectangle) - Blue
    manager_shape = ws.shapes.add(
        MsoDrawingType.RECTANGLE,
        upper_left_row=1,
        upper_left_column=12,
        lower_right_row=4,
        lower_right_column=15
    )
    manager_shape.name = "Manager"
    manager_shape.text = "Manager\nApproval"
    manager_shape.fill.fill_type = FillType.SOLID
    manager_shape.fill.fore_color = "4169E1"  # Royal blue
    manager_shape.line.is_visible = True
    manager_shape.line.color = "000080"  # Navy
    manager_shape.line.weight = 12700
    manager_shape.font.name = "Arial"
    manager_shape.font.size = 11.0
    manager_shape.font.bold = True
    manager_shape.font.color = "FFFFFF"  # White text
    manager_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    manager_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {manager_shape.name} (Rectangle)")
    
    # Shape 4: Arrow 2 (Right Arrow)
    arrow2 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=2,
        upper_left_column=15,
        lower_right_row=4,
        lower_right_column=17
    )
    arrow2.name = "Arrow2"
    arrow2.text = "→"
    arrow2.fill.fill_type = FillType.SOLID
    arrow2.fill.fore_color = "D3D3D3"
    arrow2.line.is_visible = True
    arrow2.line.color = "696969"
    arrow2.line.weight = 9525
    arrow2.font.name = "Arial"
    arrow2.font.size = 16.0
    arrow2.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow2.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow2.name} (Right Arrow)")
    
    # Shape 5: Finance Approval (Rectangle) - Teal
    finance_shape = ws.shapes.add(
        MsoDrawingType.RECTANGLE,
        upper_left_row=1,
        upper_left_column=17,
        lower_right_row=4,
        lower_right_column=20
    )
    finance_shape.name = "Finance"
    finance_shape.text = "Finance\nApproval"
    finance_shape.fill.fill_type = FillType.SOLID
    finance_shape.fill.fore_color = "008080"  # Teal
    finance_shape.line.is_visible = True
    finance_shape.line.color = "004040"
    finance_shape.line.weight = 12700
    finance_shape.font.name = "Arial"
    finance_shape.font.size = 11.0
    finance_shape.font.bold = True
    finance_shape.font.color = "FFFFFF"
    finance_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    finance_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {finance_shape.name} (Rectangle)")
    
    # Shape 6: Arrow 3 (Right Arrow)
    arrow3 = ws.shapes.add(
        MsoDrawingType.RIGHT_ARROW,
        upper_left_row=2,
        upper_left_column=20,
        lower_right_row=4,
        lower_right_column=22
    )
    arrow3.name = "Arrow3"
    arrow3.text = "→"
    arrow3.fill.fill_type = FillType.SOLID
    arrow3.fill.fore_color = "D3D3D3"
    arrow3.line.is_visible = True
    arrow3.line.color = "696969"
    arrow3.line.weight = 9525
    arrow3.font.name = "Arial"
    arrow3.font.size = 16.0
    arrow3.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow3.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow3.name} (Right Arrow)")
    
    # Shape 7: Final Decision (Diamond) - Yellow
    decision_shape = ws.shapes.add(
        MsoDrawingType.DIAMOND,
        upper_left_row=1,
        upper_left_column=22,
        lower_right_row=4,
        lower_right_column=25
    )
    decision_shape.name = "Decision"
    decision_shape.text = "Approved?"
    decision_shape.fill.fill_type = FillType.SOLID
    decision_shape.fill.fore_color = "FFD700"
    decision_shape.line.is_visible = True
    decision_shape.line.color = "B8860B"
    decision_shape.line.weight = 12700
    decision_shape.font.name = "Arial"
    decision_shape.font.size = 11.0
    decision_shape.font.bold = True
    decision_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    decision_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {decision_shape.name} (Diamond)")
    
    # Shape 8: Arrow Down (Down Arrow)
    arrow_down = ws.shapes.add(
        MsoDrawingType.DOWN_ARROW,
        upper_left_row=4,
        upper_left_column=23,
        lower_right_row=7,
        lower_right_column=25
    )
    arrow_down.name = "ArrowDown"
    arrow_down.text = "↓"
    arrow_down.fill.fill_type = FillType.SOLID
    arrow_down.fill.fore_color = "D3D3D3"
    arrow_down.line.is_visible = True
    arrow_down.line.color = "696969"
    arrow_down.line.weight = 9525
    arrow_down.font.name = "Arial"
    arrow_down.font.size = 16.0
    arrow_down.text_horizontal_alignment = TextAlignmentType.CENTER
    arrow_down.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {arrow_down.name} (Down Arrow)")
    
    # Shape 9: Complete (Star) - Green
    complete_shape = ws.shapes.add(
        MsoDrawingType.STAR_5,
        upper_left_row=7,
        upper_left_column=22,
        lower_right_row=11,
        lower_right_column=26
    )
    complete_shape.name = "Complete"
    complete_shape.text = "★\nComplete"
    complete_shape.fill.fill_type = FillType.SOLID
    complete_shape.fill.fore_color = "32CD32"  # Lime green
    complete_shape.line.is_visible = True
    complete_shape.line.color = "006400"
    complete_shape.line.weight = 12700
    complete_shape.font.name = "Arial"
    complete_shape.font.size = 12.0
    complete_shape.font.bold = True
    complete_shape.text_horizontal_alignment = TextAlignmentType.CENTER
    complete_shape.text_vertical_alignment = TextAnchorType.MIDDLE
    print(f"  Added: {complete_shape.name} (Star)")
    
    print(f"\nTotal shapes added: {len(ws.shapes)}")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        drawing_xml = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        
        # Verify different shape types
        assert '<a:prstGeom prst="ellipse">' in drawing_xml, "Missing oval geometry"
        assert '<a:prstGeom prst="rect">' in drawing_xml, "Missing rectangle geometry"
        assert '<a:prstGeom prst="diamond">' in drawing_xml, "Missing diamond geometry"
        assert '<a:prstGeom prst="star5">' in drawing_xml, "Missing star geometry"
        
        # Verify text content (newlines are converted to separate paragraph elements)
        assert "Submit" in drawing_xml and "Request" in drawing_xml, "Missing request text"
        assert "Manager" in drawing_xml and "Approval" in drawing_xml, "Missing manager text"
        assert "Finance" in drawing_xml and "Approval" in drawing_xml, "Missing finance text"
        assert "Approved?" in drawing_xml, "Missing decision text"
        assert "Complete" in drawing_xml, "Missing complete text"
    
    print(f"[OK] Successfully saved workbook with {len(ws.shapes)} shapes")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {len(ws_verify.shapes)} shapes in saved file")
    for i, shape in enumerate(ws_verify.shapes):
        print(f"  Shape {i}: type={shape.drawing_type}, name='{shape.name}', text='{shape.text.replace(chr(10), ' ')}'")
    
    return output_path


def test_all_create_shapes():
    """
    Run all test cases for creating Excel files with workflow shapes.
    """
    print("\n" + "="*70)
    print("Test: Create Excel Files with Workflow Shapes")
    print("="*70)
    
    test_create_workflow_with_shapes()
    test_create_approval_workflow()
    
    print("\n" + "="*70)
    print("All tests completed successfully!")
    print("="*70 + "\n")


if __name__ == "__main__":
    test_all_create_shapes()
