"""
MarkItDown Excel plugin that leverages Aspose.Cells.Python's Markdown converter.

This plugin integrates Aspose.Cells.Python with Microsoft MarkItDown to provide
enhanced Excel-to-Markdown conversion with metadata, multi-sheet support,
and professional formatting.

Part of the Aspose.org open source ecosystem.
"""
from typing import BinaryIO, Any
import tempfile
import os
import logging

logger = logging.getLogger(__name__)

__plugin_interface_version__ = 1


def register_converters(markitdown, **kwargs):
    """Register Aspose.Cells.Python's enhanced Excel converter with MarkItDown."""
    markitdown.register_converter(ExcelEnhancerConverter())


class ExcelEnhancerConverter:
    """Enhanced Excel converter using Aspose.Cells.Python's MarkItDownConverter."""

    # Hints for MarkItDown converter discovery systems
    file_extensions = [".xlsx"]
    mimetypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
    name = "Excel Enhancer"
    priority = 50  # Prefer this converter over generic ones for .xlsx

    def accepts(self, file_stream: BinaryIO, stream_info, **kwargs: Any) -> bool:
        """Return True if the stream describes an .xlsx file."""
        # MarkItDown's StreamInfo may expose different fields depending on source
        extension = (
            (getattr(stream_info, "extension", None) or
             getattr(stream_info, "suffix", None) or
             "").lower()
        )
        filename = (getattr(stream_info, "filename", None) or "").lower()
        mimetype = (getattr(stream_info, "mimetype", None) or "").lower()

        # Only support modern .xlsx
        return (
            extension == ".xlsx"
            or (filename and filename.endswith(".xlsx") and not filename.endswith(".xls"))
            or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" in mimetype
        )

    def convert(self, file_stream: BinaryIO, stream_info, **kwargs: Any):
        """Convert given Excel content to Markdown using our converter."""
        try:
            from markitdown import DocumentConverterResult
        except ImportError:
            # Fallback lightweight result object if markitdown is not installed
            class DocumentConverterResult:  # type: ignore
                def __init__(self, text_content):
                    self.text_content = text_content

        try:
            # Persist incoming stream to a temporary .xlsx file
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                if hasattr(file_stream, "read"):
                    content = file_stream.read()
                    if hasattr(file_stream, "seek"):
                        file_stream.seek(0)  # Reset for potential re-use elsewhere
                    tmp.write(content)
                else:
                    # file_stream may be a file path
                    with open(file_stream, "rb") as f:  # type: ignore[arg-type]
                        tmp.write(f.read())
                tmp_path = tmp.name

            # Load workbook using our implementation
            from ...workbook import Workbook

            workbook = Workbook.load(tmp_path)

            # Convert to MarkItDown format optimized for LLMs using enhanced MarkdownConverter
            from ...converters.markdown_converter import MarkdownConverter

            converter = MarkdownConverter()

            # Simplified and optimized parameters for better user experience
            convert_kwargs = {
                "sheet_name": kwargs.get("sheet_name", None),  # Convert specific sheet by name, None means all sheets
                "include_metadata": kwargs.get("include_metadata", True),
                "value_mode": kwargs.get("value_mode", "value"),  # "value" shows calculated results, "formula" shows formulas
                "include_hyperlinks": kwargs.get("include_hyperlinks", True),  # Convert hyperlinks to markdown
            }
            markdown_content = converter.convert_workbook(workbook, **convert_kwargs)

            # Optional generator banner for disambiguation in outputs
            if kwargs.get("include_generator_info", False):
                banner = "<!-- Generator: Aspose.Cells.Python MarkItDown Plugin -->\n\n"
                markdown_content = banner + markdown_content

            # Cleanup temp file
            try:
                os.unlink(tmp_path)
            except OSError:
                logger.debug("Temp file already removed or locked: %s", tmp_path)

            logger.info("Converted .xlsx using enhanced Excel converter")
            return DocumentConverterResult(markdown_content)

        except Exception as e:  # pragma: no cover - defensive path
            logger.error("Excel conversion failed: %s", e)
            error_msg = (
                "# Excel conversion error\n\n"
                f"Conversion failed: {str(e)}\n\n"
                "Please verify the Excel file is a valid .xlsx workbook."
            )
            return DocumentConverterResult(error_msg)


# MarkItDown plugin interface
__plugin_interface_version__ = 1

def register_converters(markitdown, **kwargs):
    """
    Register Aspose.Cells.Python's enhanced Excel converter with MarkItDown.
    
    This function is called by MarkItDown when enable_plugins=True.
    """
    markitdown.register_converter(ExcelEnhancerConverter())