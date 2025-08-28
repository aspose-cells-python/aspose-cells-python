"""
Format handler factory for unified file processing.
"""

from typing import Dict, Optional, Type, List
from pathlib import Path
from .interfaces import IFormatHandler


class FormatHandlerFactory:
    """Factory for managing format handlers."""
    
    _handlers: Dict[str, Type[IFormatHandler]] = {}
    _instances: Dict[str, IFormatHandler] = {}
    
    @classmethod
    def register(cls, extension: str, handler_class: Type[IFormatHandler]) -> None:
        """Register format handler for file extension."""
        if not extension.startswith('.'):
            extension = '.' + extension
        cls._handlers[extension.lower()] = handler_class
        # Clear instance cache when registering new handler
        if extension.lower() in cls._instances:
            del cls._instances[extension.lower()]
    
    @classmethod
    def get_handler(cls, file_path: str) -> Optional[IFormatHandler]:
        """Get format handler for file extension."""
        ext = Path(file_path).suffix.lower()
        
        if ext not in cls._handlers:
            return None
        
        # Use singleton pattern for handler instances
        if ext not in cls._instances:
            cls._instances[ext] = cls._handlers[ext]()
        
        return cls._instances[ext]
    
    @classmethod
    def get_supported_formats(cls) -> List[str]:
        """Get list of supported file extensions."""
        return list(cls._handlers.keys())
    
    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        """Check if file format is supported."""
        ext = Path(file_path).suffix.lower()
        return ext in cls._handlers
    
    @classmethod
    def clear_cache(cls) -> None:
        """Clear handler instance cache."""
        cls._instances.clear()


def _register_builtin_formats():
    """Register built-in format handlers."""
    
    # XLSX Handler
    class XlsxHandler(IFormatHandler):
        """Handler for XLSX format."""
        
        def __init__(self):
            from .xlsx.reader import XlsxReader
            from .xlsx.writer import XlsxWriter
            self._reader = XlsxReader()
            self._writer = XlsxWriter()
        
        def load_workbook(self, workbook, file_path: str, **options):
            return self._reader.load_workbook(workbook, file_path, **options)
        
        def save_workbook(self, workbook, file_path: str, **options):
            return self._writer.save_workbook(workbook, file_path, **options)
    
    FormatHandlerFactory.register('.xlsx', XlsxHandler)
    FormatHandlerFactory.register('.xlsm', XlsxHandler)
    FormatHandlerFactory.register('.xltx', XlsxHandler)
    FormatHandlerFactory.register('.xltm', XlsxHandler)
    
    # JSON Handler
    class JsonHandler(IFormatHandler):
        """Handler for JSON format."""
        
        def __init__(self):
            from .json.reader import JsonReader
            from .json.writer import JsonWriter
            self._reader = JsonReader()
            self._writer = JsonWriter()
        
        def load_workbook(self, workbook, file_path: str, **options):
            return self._reader.load_workbook(workbook, file_path, **options)
        
        def save_workbook(self, workbook, file_path: str, **options):
            return self._writer.save_workbook(workbook, file_path, **options)
    
    FormatHandlerFactory.register('.json', JsonHandler)
    
    # CSV Handler
    class CsvHandler(IFormatHandler):
        """Handler for CSV format."""
        
        def __init__(self):
            from .csv.reader import CsvReader
            from .csv.writer import CsvWriter
            self._reader = CsvReader()
            self._writer = CsvWriter()
        
        def load_workbook(self, workbook, file_path: str, **options):
            return self._reader.load_workbook(workbook, file_path, **options)
        
        def save_workbook(self, workbook, file_path: str, **options):
            return self._writer.save_workbook(workbook, file_path, **options)
    
    FormatHandlerFactory.register('.csv', CsvHandler)
    
    # Markdown Handler
    class MarkdownHandler(IFormatHandler):
        """Handler for Markdown format."""
        
        def __init__(self):
            from .md.reader import MarkdownReader
            from .md.writer import MarkdownWriter
            self._reader = MarkdownReader()
            self._writer = MarkdownWriter()
        
        def load_workbook(self, workbook, file_path: str, **options):
            return self._reader.load_workbook(workbook, file_path, **options)
        
        def save_workbook(self, workbook, file_path: str, **options):
            return self._writer.save_workbook(workbook, file_path, **options)
    
    FormatHandlerFactory.register('.md', MarkdownHandler)
    FormatHandlerFactory.register('.markdown', MarkdownHandler)


# Initialize built-in formats
_register_builtin_formats()