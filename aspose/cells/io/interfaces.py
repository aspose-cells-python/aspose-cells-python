"""
Unified interfaces for format handlers.
"""

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..workbook import Workbook
    from .models import WorkbookData


class IFormatHandler(ABC):
    """Unified interface for format handlers."""
    
    @abstractmethod
    def load_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """
        Load file into workbook object.
        Maintains compatibility with existing interface.
        """
        pass
    
    @abstractmethod
    def save_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """
        Save workbook object to file.
        Maintains compatibility with existing interface.
        """
        pass
    
    def read_to_data(self, file_path: str, **options) -> 'WorkbookData':
        """Read file and return unified data model."""
        from ..workbook import Workbook
        from .models import WorkbookData
        
        temp_workbook = Workbook()
        # Clear default sheet since we're loading from file
        temp_workbook._worksheets.clear()
        temp_workbook._active_sheet = None
        
        self.load_workbook(temp_workbook, file_path, **options)
        return WorkbookData.from_workbook(temp_workbook)
    
    def write_from_data(self, data: 'WorkbookData', file_path: str, **options) -> None:
        """Write unified data model to file."""
        workbook = data.to_workbook()
        self.save_workbook(workbook, file_path, **options)