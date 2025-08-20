"""
Comprehensive Docling Backend Plugin Tests
Test CellsDocumentBackend integration with comparison outputs and timing.
Based on markitdown plugin test structure.
"""

import pytest
import time
from pathlib import Path

try:
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.document import InputDocument
    from docling_core.types.doc import DoclingDocument
    from docling.document_converter import DocumentConverter
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False

from aspose.cells.plugins.docling_backend import CellsDocumentBackend
from aspose.cells.plugins.docling_backend.backend import AsposeCellsDoclingDocument


@pytest.mark.skipif(not DOCLING_AVAILABLE, reason="Docling not installed")
class TestDoclingBackend:
    """Test Docling backend plugin with comprehensive comparisons."""
    
    def test_backend_initialization(self, ensure_testdata_dir):
        """Test backend initializes correctly."""
        xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
        
        in_doc = InputDocument(
            path_or_stream=xlsx_file,
            format=InputFormat.XLSX,
            backend=CellsDocumentBackend,
            filename="test.xlsx",
        )
        
        backend = CellsDocumentBackend(in_doc=in_doc, path_or_stream=xlsx_file)
        assert backend.is_valid()
        assert backend.supports_pagination()
        assert backend.page_count() > 0
    
    def test_supported_formats(self):
        """Test supported formats."""
        formats = CellsDocumentBackend.supported_formats()
        assert InputFormat.XLSX in formats
    
    def test_convert_to_markdown(self, ensure_testdata_dir):
        """Test Excel to Markdown conversion."""
        xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
        
        in_doc = InputDocument(
            path_or_stream=xlsx_file,
            format=InputFormat.XLSX,
            backend=CellsDocumentBackend,
            filename="test.xlsx",
        )
        
        backend = CellsDocumentBackend(in_doc=in_doc, path_or_stream=xlsx_file)
        doc = backend.convert()
        
        assert isinstance(doc, AsposeCellsDoclingDocument)
        assert doc.name == "sales_report_comprehensive"
        assert len(doc.pages) > 0
        
        # Convert to markdown
        markdown = doc.export_to_markdown()
        assert isinstance(markdown, str)
        assert len(markdown) > 0

    def test_docling_plugin_vs_native_comparison(self, ensure_testdata_dir):
        """Compare Docling output with/without Aspose plugin for comprehensive files."""
        # Set up dedicated output folder
        output_dir = Path(__file__).parent / "testdata" / "test_docling_backend"
        output_dir.mkdir(exist_ok=True)
        
        xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
        file_stem = xlsx_file.stem
        
        # Native docling conversion (baseline)
        converter = DocumentConverter()
        t0 = time.perf_counter()
        native_result = converter.convert(str(xlsx_file))
        native_time = (time.perf_counter() - t0) * 1000.0
        native_markdown = native_result.document.export_to_markdown()
        
        native_output = output_dir / f"docling_native_{file_stem}.md"
        with open(native_output, "w", encoding="utf-8") as f:
            f.write(native_markdown)
        
        # Plugin conversion (enhanced)
        in_doc = InputDocument(
            path_or_stream=xlsx_file,
            format=InputFormat.XLSX,
            backend=CellsDocumentBackend,
            filename=xlsx_file.name,
        )
        
        t1 = time.perf_counter()
        backend = CellsDocumentBackend(in_doc=in_doc, path_or_stream=xlsx_file)
        plugin_doc = backend.convert()
        plugin_time = (time.perf_counter() - t1) * 1000.0
        plugin_markdown = plugin_doc.export_to_markdown()
        
        plugin_output = output_dir / f"plugin_{file_stem}.md"
        with open(plugin_output, "w", encoding="utf-8") as f:
            f.write(plugin_markdown)
        
        # Record timing comparison
        comparison_path = output_dir / f"comparison_{file_stem}_timing.txt"
        with open(comparison_path, "w", encoding="utf-8") as f:
            f.write(f"Docling conversion comparison for {xlsx_file.name}\n")
            f.write(f"- Native docling: {native_time:.2f} ms\n")
            f.write(f"- Aspose plugin: {plugin_time:.2f} ms\n")
            f.write(f"- Difference (plugin - native): {plugin_time - native_time:.2f} ms\n")
            f.write(f"- Native output size: {len(native_markdown)} chars\n")
            f.write(f"- Plugin output size: {len(plugin_markdown)} chars\n")
            f.write(f"- Plugin pages: {len(plugin_doc.pages)}\n")
        
        print(f"[{file_stem}] Native: {native_time:.2f}ms, Plugin: {plugin_time:.2f}ms")
        
        # Verify files created
        assert native_output.exists()
        assert plugin_output.exists()
        assert comparison_path.exists()

    def test_docling_export_formats(self, ensure_testdata_dir):
        """Test multiple export formats from Docling backend."""
        output_dir = Path(__file__).parent / "testdata" / "test_docling_backend"
        output_dir.mkdir(exist_ok=True)
        
        xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
        
        # Convert with plugin
        in_doc = InputDocument(
            path_or_stream=xlsx_file,
            format=InputFormat.XLSX,
            backend=CellsDocumentBackend,
            filename=xlsx_file.name,
        )
        
        backend = CellsDocumentBackend(in_doc=in_doc, path_or_stream=xlsx_file)
        doc = backend.convert()
        
        # Test markdown export
        markdown = doc.export_to_markdown()
        md_output = output_dir / "plugin_export_test.md"
        with open(md_output, "w", encoding="utf-8") as f:
            f.write(markdown)
        assert md_output.exists()
        assert len(markdown) > 0
        
        # Test JSON export if available
        try:
            json_content = doc.export_to_json()
            json_output = output_dir / "plugin_export_test.json"
            with open(json_output, "w", encoding="utf-8") as f:
                f.write(json_content)
            assert json_output.exists()
            assert len(json_content) > 0
        except AttributeError:
            # JSON export might not be available in all docling versions
            print("JSON export not available in this docling version")

    def test_docling_plugin_parameter_variations(self, ensure_testdata_dir):
        """Test different parameter configurations with the Docling plugin."""
        output_dir = Path(__file__).parent / "testdata" / "test_docling_backend"
        output_dir.mkdir(exist_ok=True)
        
        xlsx_file = ensure_testdata_dir / "sales_report_comprehensive.xlsx"
        
        # Test different backend configurations
        test_cases = [
            ("default", {}),
            ("with_metadata", {"include_metadata": True}),
            ("without_metadata", {"include_metadata": False})
        ]
        
        timing_results = []
        
        for test_name, kwargs in test_cases:
            try:
                in_doc = InputDocument(
                    path_or_stream=xlsx_file,
                    format=InputFormat.XLSX,
                    backend=CellsDocumentBackend,
                    filename=xlsx_file.name,
                )
                
                t0 = time.perf_counter()
                backend = CellsDocumentBackend(in_doc=in_doc, path_or_stream=xlsx_file)
                doc = backend.convert(**kwargs)
                conversion_time = (time.perf_counter() - t0) * 1000.0
                
                markdown = doc.export_to_markdown()
                
                # Save output
                output_file = output_dir / f"plugin_param_{test_name}.md"
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(markdown)
                
                timing_results.append({
                    "name": test_name,
                    "time_ms": conversion_time,
                    "params": kwargs,
                    "output_size": len(markdown),
                    "pages": len(doc.pages)
                })
                
                assert output_file.exists()
                
            except Exception as e:
                print(f"Parameter test {test_name} failed: {e}")
        
        # Generate parameter comparison report
        if timing_results:
            report_path = output_dir / "docling_parameters_report.md"
            with open(report_path, "w", encoding="utf-8") as f:
                f.write("# Docling Plugin Parameters Test Report\n\n")
                f.write("## Test Cases Overview\n\n")
                f.write("| Test Case | Conversion Time (ms) | Output Size (chars) | Pages | Parameters |\n")
                f.write("|-----------|---------------------|---------------------|-------|------------|\n")
                
                for result in timing_results:
                    params_str = ", ".join([f"{k}={v}" for k, v in result["params"].items()]) if result["params"] else "default"
                    f.write(f"| {result['name']} | {result['time_ms']:.2f} | {result['output_size']} | {result['pages']} | {params_str} |\n")
                
                f.write("\n## Parameter Descriptions\n\n")
                f.write("- **include_metadata**: Include document metadata in conversion\n")
                f.write(f"\n## Generated Files\n\n")
                for result in timing_results:
                    f.write(f"- `plugin_param_{result['name']}.md`: {result['params'] if result['params'] else 'Default configuration'}\n")
            
            assert report_path.exists()
            print(f"Generated {len(timing_results)} parameter test files and comparison report")