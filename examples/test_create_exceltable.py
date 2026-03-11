"""
Test case to create an Excel file with dummy data and add an Excel table.

This test demonstrates:
1. Creating a new Excel workbook from scratch
2. Adding dummy data to the worksheet
3. Creating an Excel table from the data
4. Saving the workbook to a file
"""

import os
import sys
import zipfile

# Add parent directory to path to import aspose.cells_foss
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose.cells_foss import Workbook


def test_create_excel_with_table():
    """
    Create a new Excel file with dummy data and add an Excel table.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/exceltable", exist_ok=True)
    
    output_path = "tests/outputfiles/exceltable/test_create_table.xlsx"
    
    # Create a new workbook
    print("Creating new workbook...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # Set worksheet name
    ws.name = "SalesData"
    
    # Add dummy data for the table
    print("\nAdding dummy data for table...")
    # Add headers
    ws.cells["A1"].value = "Product"
    ws.cells["B1"].value = "Category"
    ws.cells["C1"].value = "Quantity"
    ws.cells["D1"].value = "Price"
    ws.cells["E1"].value = "Total"
    
    # Add data rows
    dummy_data = [
        ["Laptop", "Electronics", 5, 999.99, "=C2*D2"],
        ["Mouse", "Electronics", 20, 29.99, "=C3*D3"],
        ["Keyboard", "Electronics", 15, 79.99, "=C4*D4"],
        ["Monitor", "Electronics", 8, 299.99, "=C5*D5"],
        ["Headphones", "Electronics", 12, 149.99, "=C6*D6"],
        ["Desk Chair", "Furniture", 3, 249.99, "=C7*D7"],
        ["Desk Lamp", "Furniture", 10, 49.99, "=C8*D8"],
        ["Notebook", "Stationery", 50, 4.99, "=C9*D9"],
        ["Pen Set", "Stationery", 30, 12.99, "=C10*D10"],
        ["USB Cable", "Accessories", 25, 9.99, "=C11*D11"],
    ]
    
    for row_idx, row_data in enumerate(dummy_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    print(f"Added {len(dummy_data)} rows of data")
    
    # Create a table from the dummy data (A1:E11)
    print("\nCreating Excel table from dummy data...")
    table = ws.tables.add(
        start_row=0,      # Row 1 (0-based)
        start_col=0,      # Column A (0-based)
        end_row=10,       # Row 11 (0-based)
        end_col=4,        # Column E (0-based)
        has_headers=True,
        name="SalesTable"
    )
    
    # Customize table style
    table.table_style_info.name = "TableStyleMedium9"
    table.table_style_info.show_row_stripes = True
    table.table_style_info.show_first_column = True
    table.table_style_info.show_last_column = False
    table.table_style_info.show_column_stripes = False
    
    print(f"Table created: name='{table.name}', ref='{table.ref}', columns={len(table.columns)}")
    for j, col in enumerate(table.columns):
        print(f"  Column {j}: name='{col.name}'")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())
        # Check that table file exists
        assert "xl/tables/table1.xml" in names, "Missing table part"
        
        # Read and verify table XML
        table_xml = zf.read("xl/tables/table1.xml").decode("utf-8")
        
        # Verify table elements exist
        assert "<table" in table_xml, "Missing table element in table XML"
        
        # Verify table reference
        assert 'ref="A1:E11"' in table_xml, "Missing correct table reference"
        
        # Verify table name
        assert 'name="SalesTable"' in table_xml, "Missing table name"
        
        # Verify table style
        assert 'name="TableStyleMedium9"' in table_xml, "Missing table style"
        
        # Verify column names
        assert 'name="Product"' in table_xml, "Missing Product column"
        assert 'name="Category"' in table_xml, "Missing Category column"
        assert 'name="Quantity"' in table_xml, "Missing Quantity column"
        assert 'name="Price"' in table_xml, "Missing Price column"
        assert 'name="Total"' in table_xml, "Missing Total column"
    
    print(f"[OK] Successfully saved workbook with table")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.tables.count} table(s) in saved file")
    for i, table in enumerate(ws_verify.tables):
        print(f"  Table {i}: name='{table.name}', ref='{table.ref}', columns={len(table.columns)}")
        for j, col in enumerate(table.columns):
            print(f"    Column {j}: name='{col.name}'")
    
    return output_path


def test_create_multiple_tables():
    """
    Create a new Excel file with multiple tables in different ranges.
    """
    # Ensure output directory exists
    os.makedirs("tests/outputfiles/exceltable", exist_ok=True)
    
    output_path = "tests/outputfiles/exceltable/test_create_multiple_tables.xlsx"
    
    # Create a new workbook
    print("\nCreating new workbook with multiple tables...")
    wb = Workbook()
    ws = wb.worksheets[0]
    
    ws.name = "MultipleTables"
    
    # Add first table data (Employees)
    print("\nAdding first table data (Employees)...")
    ws.cells["A1"].value = "Employee ID"
    ws.cells["B1"].value = "Name"
    ws.cells["C1"].value = "Department"
    ws.cells["D1"].value = "Salary"
    
    employee_data = [
        ["E001", "John Smith", "Engineering", 85000],
        ["E002", "Jane Doe", "Marketing", 75000],
        ["E003", "Bob Johnson", "Engineering", 90000],
        ["E004", "Alice Brown", "HR", 65000],
        ["E005", "Charlie Wilson", "Finance", 80000],
    ]
    
    for row_idx, row_data in enumerate(employee_data, start=2):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    # Create first table
    table1 = ws.tables.add(
        start_row=0, start_col=0, end_row=5, end_col=3,
        has_headers=True, name="EmployeesTable"
    )
    table1.table_style_info.name = "TableStyleMedium7"
    table1.table_style_info.show_row_stripes = True
    
    print(f"First table created: name='{table1.name}', ref='{table1.ref}'")
    
    # Add second table data (Projects) - starting at row 8
    print("\nAdding second table data (Projects)...")
    ws.cells["A8"].value = "Project ID"
    ws.cells["B8"].value = "Project Name"
    ws.cells["C8"].value = "Status"
    ws.cells["D8"].value = "Budget"
    
    project_data = [
        ["P001", "Website Redesign", "In Progress", 50000],
        ["P002", "Mobile App", "Planning", 75000],
        ["P003", "Database Migration", "Completed", 30000],
        ["P004", "Cloud Integration", "On Hold", 60000],
    ]
    
    for row_idx, row_data in enumerate(project_data, start=9):
        for col_idx, value in enumerate(row_data):
            cell_ref = f"{chr(65 + col_idx)}{row_idx}"
            ws.cells[cell_ref].value = value
    
    # Create second table
    table2 = ws.tables.add(
        start_row=7, start_col=0, end_row=11, end_col=3,
        has_headers=True, name="ProjectsTable"
    )
    table2.table_style_info.name = "TableStyleMedium11"
    table2.table_style_info.show_row_stripes = True
    
    print(f"Second table created: name='{table2.name}', ref='{table2.ref}'")
    
    # Save the workbook
    print(f"\nSaving workbook to: {output_path}")
    wb.save(output_path)
    
    # Verify the file was created
    assert os.path.exists(output_path), f"Output file not created: {output_path}"
    
    # Verify the XML structure
    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())
        # Check that both table files exist
        assert "xl/tables/table1.xml" in names, "Missing first table part"
        assert "xl/tables/table2.xml" in names, "Missing second table part"
        
        # Read and verify table XMLs
        table1_xml = zf.read("xl/tables/table1.xml").decode("utf-8")
        table2_xml = zf.read("xl/tables/table2.xml").decode("utf-8")
        
        # Verify first table
        assert 'name="EmployeesTable"' in table1_xml, "Missing first table name"
        assert 'ref="A1:D6"' in table1_xml, "Missing correct reference in first table"
        
        # Verify second table
        assert 'name="ProjectsTable"' in table2_xml, "Missing second table name"
        assert 'ref="A8:D12"' in table2_xml, "Missing correct reference in second table"
    
    print(f"[OK] Successfully saved workbook with {ws.tables.count} table(s)")
    
    # Reload and verify
    print("\nReloading to verify...")
    wb_verify = Workbook(output_path)
    ws_verify = wb_verify.worksheets[0]
    print(f"[OK] Verified: {ws_verify.tables.count} table(s) in saved file")
    for i, table in enumerate(ws_verify.tables):
        print(f"  Table {i}: name='{table.name}', ref='{table.ref}', columns={len(table.columns)}")
    
    return output_path


def test_all_create_tables():
    """
    Run all test cases for creating Excel files with tables.
    """
    print("\n" + "="*70)
    print("Test: Create Excel Files with Tables")
    print("="*70)
    
    test_create_excel_with_table()
    test_create_multiple_tables()
    
    print("\n" + "="*70)
    print("All tests completed successfully!")
    print("="*70 + "\n")


if __name__ == "__main__":
    test_all_create_tables()
