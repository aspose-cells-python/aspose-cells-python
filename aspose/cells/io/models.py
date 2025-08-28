"""
Unified data models for cross-format workbook operations.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from ..worksheet import Worksheet
    from ..workbook import Workbook


@dataclass
class WorkbookData:
    """Unified workbook data model for cross-format operations."""
    
    worksheets: Dict[str, 'Worksheet'] = field(default_factory=dict)
    active_sheet_name: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def add_worksheet(self, name: str, worksheet: 'Worksheet') -> None:
        """Add worksheet to the data model."""
        self.worksheets[name] = worksheet
        if self.active_sheet_name is None:
            self.active_sheet_name = name
    
    def to_workbook(self) -> 'Workbook':
        """Convert unified data model to Workbook object."""
        from ..workbook import Workbook
        
        wb = Workbook.__new__(Workbook)  # Create without calling __init__
        wb._worksheets = {}
        wb._active_sheet = None
        wb._shared_strings = []
        wb._properties = self.metadata.copy()
        wb._filename = None
        
        # Copy worksheets
        for name, worksheet in self.worksheets.items():
            wb._worksheets[name] = worksheet
            worksheet._parent = wb  # Update parent reference
        
        # Set active sheet
        if self.active_sheet_name and self.active_sheet_name in wb._worksheets:
            wb._active_sheet = wb._worksheets[self.active_sheet_name]
        elif wb._worksheets:
            wb._active_sheet = next(iter(wb._worksheets.values()))
        
        return wb
    
    @classmethod
    def from_workbook(cls, workbook: 'Workbook') -> 'WorkbookData':
        """Create unified data model from Workbook object."""
        active_name = None
        if workbook._active_sheet:
            active_name = workbook._active_sheet.name
        
        metadata = getattr(workbook, '_properties', {}).copy()
        
        return cls(
            worksheets=workbook._worksheets.copy(),
            active_sheet_name=active_name,
            metadata=metadata
        )