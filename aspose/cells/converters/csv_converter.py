"""
CSV converter for exporting Excel data to CSV format.
"""

import csv
import io
from typing import Optional, TYPE_CHECKING
from ..io.csv import CsvWriter

if TYPE_CHECKING:
    from ..workbook import Workbook


class CsvConverter:
    """Convert Excel workbook data to CSV format."""
    
    def __init__(self):
        self._writer = CsvWriter()
    
    def convert_workbook(self, workbook: 'Workbook', **kwargs) -> str:
        """Convert active worksheet to CSV string."""
        sheet_name = kwargs.get('sheet_name')
        delimiter = kwargs.get('delimiter', ',')
        quotechar = kwargs.get('quotechar', '"')
        
        # Get target worksheet
        if sheet_name and sheet_name in workbook._worksheets:
            worksheet = workbook._worksheets[sheet_name]
        else:
            worksheet = workbook.active
        
        if not worksheet or not worksheet._cells:
            return ""
        
        # Convert worksheet to data
        data = self._writer._worksheet_to_data(worksheet)
        
        if not data:
            return ""
        
        # Create CSV in memory
        output = io.StringIO()
        writer = csv.writer(output, delimiter=delimiter, quotechar=quotechar, 
                           quoting=csv.QUOTE_MINIMAL)
        
        # Write data rows
        for row_data in data:
            formatted_row = []
            for cell in row_data:
                formatted_row.append(self._writer._format_cell_value(cell))
            writer.writerow(formatted_row)
        
        csv_content = output.getvalue()
        output.close()
        return csv_content