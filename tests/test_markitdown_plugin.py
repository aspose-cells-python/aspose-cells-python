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


@pytest.mark.skipif(not MARKITDOWN_AVAILABLE, reason="MarkItDown not installed")
def test_markitdown_parameters_variations(ensure_testdata_dir):
    """Test each MarkItDown plugin parameter individually with output files for comparison"""
    xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
    
    # Test simplified parameters with their variations
    test_cases = [
        # Default configuration (baseline)
        ("default", {}),
        
        # Test sheet_name parameter - convert specific sheet
        ("sheet_name_specific", {"sheet_name": "Summary"}),
        ("sheet_name_nonexistent", {"sheet_name": "NonExistent"}),
        
        # Test include_metadata parameter
        ("include_metadata_false", {"include_metadata": False}),
        ("include_metadata_true", {"include_metadata": True}),
        
        # Test value_mode parameter (renamed from cell_value_mode)
        ("value_mode_value", {"value_mode": "value"}),
        ("value_mode_formula", {"value_mode": "formula"}),
        
        # Test include_hyperlinks parameter
        ("include_hyperlinks_false", {"include_hyperlinks": False}),
        ("include_hyperlinks_true", {"include_hyperlinks": True}),
        
        # Test include_generator_info parameter
        ("include_generator_info_false", {"include_generator_info": False}),
        ("include_generator_info_true", {"include_generator_info": True}),
        
        # Test combination scenarios with simplified parameters
        ("minimal_output", {
            "include_metadata": False,
            "value_mode": "value",
            "include_hyperlinks": False,
            "include_generator_info": False
        }),
        ("detailed_output", {
            "include_metadata": True,
            "value_mode": "value",
            "include_hyperlinks": True,
            "include_generator_info": True
        }),
        ("formula_focused", {
            "include_metadata": True,
            "value_mode": "formula",
            "include_hyperlinks": True,
            "include_generator_info": True
        })
    ]
    
    md_enhanced = MarkItDown(enable_plugins=True)
    timing_results = []
    
    for test_name, kwargs in test_cases:
        t0 = time.perf_counter()
        result = md_enhanced.convert(str(xlsx_file), **kwargs)
        conversion_time = (time.perf_counter() - t0) * 1000.0
        
        # Save output to individual files for comparison
        output_file = ensure_testdata_dir / f"markitdown_param_{test_name}.md"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(result.text_content)
        
        # Record timing and parameters
        timing_results.append({
            "name": test_name,
            "time_ms": conversion_time,
            "params": kwargs,
            "output_size": len(result.text_content)
        })
        
        # Verify file was created
        assert output_file.exists()
        assert output_file.stat().st_size > 0
    
    # Generate comparison report
    report_path = ensure_testdata_dir / "markitdown_parameters_report.md"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# MarkItDown Plugin Parameters Test Report\n\n")
        f.write("## Test Cases Overview\n\n")
        f.write("| Test Case | Conversion Time (ms) | Output Size (chars) | Parameters |\n")
        f.write("|-----------|---------------------|---------------------|------------|\n")
        
        for result in timing_results:
            params_str = ", ".join([f"{k}={v}" for k, v in result["params"].items()]) if result["params"] else "default"
            f.write(f"| {result['name']} | {result['time_ms']:.2f} | {result['output_size']} | {params_str} |\n")
        
        f.write("\n## Simplified Parameter Descriptions\n\n")
        f.write("- **sheet_name**: Convert specific sheet by name (None = all sheets)\n")
        f.write("- **include_metadata**: Include workbook metadata (title, author, etc.)\n")
        f.write("- **value_mode**: 'value' (calculated values) vs 'formula' (raw formulas)\n")
        f.write("- **include_hyperlinks**: Convert Excel hyperlinks to Markdown links\n")
        f.write("- **include_generator_info**: Add Aspose plugin identification banner\n")
        
        f.write(f"\n## Generated Files\n\n")
        for result in timing_results:
            f.write(f"- `markitdown_param_{result['name']}.md`: {result['params'] if result['params'] else 'Default configuration'}\n")
    
    print(f"Generated {len(test_cases)} test output files and comparison report")
    print(f"Report saved to: {report_path}")
    assert report_path.exists()