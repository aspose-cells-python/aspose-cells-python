"""
MarkItDown Excel Enhancement Plugin

Integrates Aspose.Cells.Python Excel-to-Markdown conversion with Microsoft MarkItDown.
Part of the Aspose.org open source ecosystem.
"""

from .plugin import register_converters  # re-export for convenience

__version__ = "1.1.0"
__plugin_name__ = "Excel Enhancer"
__plugin_description__ = "Enhanced Excel processing for MarkItDown (.xlsx only)"

# MarkItDown plugin interface
__plugin_interface_version__ = 1