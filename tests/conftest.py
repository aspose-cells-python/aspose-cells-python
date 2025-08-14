"""
Pytest configuration and fixtures for aspose.cells.python tests.
"""

import pytest
import os
import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Configure pytest for comprehensive testing
pytest_plugins = []

@pytest.fixture
def testdata_dir():
    """Get testdata directory path."""
    return Path(__file__).parent / "testdata"

@pytest.fixture
def ensure_testdata_dir(testdata_dir):
    """Ensure testdata directory exists."""
    testdata_dir.mkdir(exist_ok=True)
    return testdata_dir

@pytest.fixture
def sample_data():
    """Sample data for testing."""
    return [
        ["Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total", "Average", "Growth %"],
        ["Laptop", 15000, 18000, 22000, 25000, "", "", ""],
        ["Desktop", 8000, 9500, 11000, 12500, "", "", ""],
        ["Tablet", 12000, 14000, 16000, 18000, "", "", ""],
        ["Phone", 25000, 28000, 32000, 35000, "", "", ""],
        ["Monitor", 5000, 5500, 6000, 6500, "", "", ""],
        ["", "", "", "", "TOTALS:", "", "", ""],
    ]

@pytest.fixture
def financial_data():
    """Financial report data for testing."""
    return {
        "company_info": {
            "name": "Tech Solutions Inc.",
            "period": "Fiscal Year 2024",
            "prepared_by": "Finance Department"
        },
        "revenue_data": [
            ["Month", "Revenue", "Expenses", "Profit", "Margin %"],
            ["January", 120000, 85000, "", ""],
            ["February", 135000, 92000, "", ""],
            ["March", 148000, 98000, "", ""],
            ["April", 156000, 105000, "", ""],
            ["May", 162000, 108000, "", ""],
            ["June", 175000, 115000, "", ""],
        ]
    }

@pytest.fixture
def employee_data():
    """Employee data for testing."""
    return [
        ["ID", "Name", "Department", "Salary", "Bonus %", "Total Comp", "Performance"],
        [1001, "Alice Johnson", "Engineering", 95000, 15, "", "Excellent"],
        [1002, "Bob Smith", "Sales", 75000, 20, "", "Good"],
        [1003, "Carol Davis", "Marketing", 68000, 12, "", "Good"],
        [1004, "David Wilson", "Engineering", 105000, 18, "", "Excellent"],
        [1005, "Eva Brown", "HR", 72000, 10, "", "Average"],
        [1006, "Frank Miller", "Sales", 82000, 25, "", "Excellent"],
    ]

@pytest.fixture
def cleanup_test_files(ensure_testdata_dir):
    """Clean up test files after test execution."""
    yield ensure_testdata_dir
    # Cleanup happens after test - could implement if needed

@pytest.fixture(scope="session")
def test_session_info():
    """Session-wide test information."""
    return {
        "test_run_id": "excel_test_comprehensive",
        "features_tested": [
            "workbook_creation", "data_entry", "styling", "formulas",
            "worksheets", "export_formats", "data_operations", "reading"
        ]
    }