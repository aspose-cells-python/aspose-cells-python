"""
MarkItDown Plugin Test - Simplified
Compare plugin vs no-plugin output for PRESAMPLE.xlsx
"""

import pytest
import time
from tests.test_advanced_features import create_sales_workbook
from aspose.cells import FileFormat

try:
    from markitdown import MarkItDown
    MARKITDOWN_AVAILABLE = True
except ImportError:
    MARKITDOWN_AVAILABLE = False


@pytest.mark.skipif(not MARKITDOWN_AVAILABLE, reason="MarkItDown not installed")
def test_sales_report_plugin_comparison(ensure_testdata_dir):
    """Compare MarkItDown output with/without plugin for PRESAMPLE.xlsx"""
    # Create comprehensive sales workbook
    # wb = create_sales_workbook() test100
    xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
    # wb.save(str(xlsx_file))
    # wb.close()
    
    # Without plugin (baseline)
    md_basic = MarkItDown(enable_plugins=False)
    t0 = time.perf_counter()
    result_basic = md_basic.convert(str(xlsx_file))
    basic_ms = (time.perf_counter() - t0) * 1000.0
    
    basic_output = ensure_testdata_dir / "test_markitdown_basic.md"
    with open(basic_output, "w", encoding="utf-8") as f:
        f.write(result_basic.text_content)
    
    # With plugin (enhanced)
    md_enhanced = MarkItDown(enable_plugins=True)
    t1 = time.perf_counter()
    #result_enhanced = md_enhanced.convert(str(xlsx_file),include_hyperlinks=False, cell_value_mode="formula")
    result_enhanced = md_enhanced.convert(str(xlsx_file),include_hyperlinks=True)
    enhanced_ms = (time.perf_counter() - t1) * 1000.0
    
    enhanced_output = ensure_testdata_dir / "test_markitdown_enhanced.md"
    with open(enhanced_output, "w", encoding="utf-8") as f:
        f.write(result_enhanced.text_content)
    
    # Record timing comparison
    timings_path = ensure_testdata_dir / "test_markitdown_timings.txt"
    with open(timings_path, "w", encoding="utf-8") as f:
        f.write("MarkItDown conversion timings (ms)\n")
        f.write(f"- No plugin: {basic_ms:.2f} ms\n")
        f.write(f"- With plugin: {enhanced_ms:.2f} ms\n")
        f.write(f"- Difference (plugin - no-plugin): {enhanced_ms - basic_ms:.2f} ms\n")
    
    print(f"[Timing] No plugin: {basic_ms:.2f} ms; With plugin: {enhanced_ms:.2f} ms; Diff: {enhanced_ms - basic_ms:.2f} ms")
    
    # Verify files created and different sizes
    assert basic_output.exists()
    assert enhanced_output.exists()
    assert basic_output.stat().st_size != enhanced_output.stat().st_size
    assert timings_path.exists()