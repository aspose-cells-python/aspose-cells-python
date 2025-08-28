"""
Markdown table reader for loading Markdown table data into workbook format.
"""

import re
from typing import Dict, List, Optional, Union, TYPE_CHECKING
from pathlib import Path
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook


class MarkdownReader:
    """Reader for Markdown table files."""
    
    def __init__(self):
        pass
    
    def read(self, file_path: str, **kwargs) -> Union[List[List[CellValue]], Dict[str, List[List[CellValue]]]]:
        """Read Markdown file and extract table data."""
        encoding = kwargs.get('encoding', 'utf-8')
        
        try:
            with open(file_path, 'r', encoding=encoding) as file:
                content = file.read()
            
            return self._parse_markdown_tables(content)
                
        except FileNotFoundError:
            raise FileNotFoundError(f"Markdown file not found: {file_path}")
        except Exception as e:
            raise ValueError(f"Error reading Markdown file: {e}")
    
    def _parse_markdown_tables(self, content: str) -> Union[List[List[CellValue]], Dict[str, List[List[CellValue]]]]:
        """Parse markdown content and extract tables."""
        sections = self._split_by_headers(content)
        
        if len(sections) == 1 and sections[0]['name'] == 'default':
            # Single table, return as list
            return self._extract_tables_from_text(sections[0]['content'])
        else:
            # Multiple sections, return as dict
            result = {}
            for section in sections:
                tables = self._extract_tables_from_text(section['content'])
                if tables:  # Only add sections with tables
                    result[section['name']] = tables
            return result if result else [[]]
    
    def _split_by_headers(self, content: str) -> List[Dict[str, str]]:
        """Split content by markdown headers."""
        lines = content.split('\n')
        sections = []
        current_section = {'name': 'default', 'content': ''}
        
        for line in lines:
            header_match = re.match(r'^#+\s+(.+)$', line.strip())
            if header_match:
                # Save current section if it has content
                if current_section['content'].strip():
                    sections.append(current_section)
                # Start new section
                current_section = {
                    'name': header_match.group(1).strip(),
                    'content': ''
                }
            else:
                current_section['content'] += line + '\n'
        
        # Add final section
        if current_section['content'].strip():
            sections.append(current_section)
        
        return sections if sections else [{'name': 'default', 'content': content}]
    
    def _extract_tables_from_text(self, text: str) -> List[List[CellValue]]:
        """Extract table data from text content."""
        lines = text.split('\n')
        table_rows = []
        in_table = False
        
        for line in lines:
            line = line.strip()
            if not line:
                if in_table:
                    break  # End of table
                continue
            
            # Check if this is a table row (starts and ends with |)
            if line.startswith('|') and line.endswith('|'):
                # Skip separator lines (contain only |, -, :, and spaces)
                if re.match(r'^[\|\-:\s]+$', line):
                    continue
                
                in_table = True
                # Parse table row
                cells = [cell.strip() for cell in line[1:-1].split('|')]
                converted_row = []
                for cell in cells:
                    # Unescape markdown special characters
                    cell = cell.replace('\\|', '|')
                    converted_row.append(self._convert_cell_value(cell))
                table_rows.append(converted_row)
            elif in_table:
                # End of table
                break
        
        return table_rows
    
    def _convert_cell_value(self, value: str) -> CellValue:
        """Convert string value to appropriate Python type."""
        if not value or value.strip() == "":
            return None
        
        value = value.strip()
        
        # Try boolean first
        if value.upper() in ('TRUE', 'FALSE'):
            return value.upper() == 'TRUE'
        
        # Try integer
        try:
            if '.' not in value and 'e' not in value.lower():
                return int(value)
        except ValueError:
            pass
        
        # Try float
        try:
            return float(value)
        except ValueError:
            pass
        
        # Return as string
        return value
    
    def load_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """Load Markdown file into workbook object."""
        data = self.read(file_path, **options)
        
        # Clear existing worksheets
        workbook._worksheets.clear()
        workbook._active_sheet = None
        
        if isinstance(data, dict):
            # Multi-sheet format (multiple headers/sections)
            for sheet_name, table_data in data.items():
                worksheet = workbook.create_sheet(sheet_name)
                self._populate_worksheet(worksheet, table_data)
        else:
            # Single table format
            worksheet = workbook.create_sheet("Sheet1")
            self._populate_worksheet(worksheet, data)
    
    def _populate_worksheet(self, worksheet, rows: List[List[CellValue]]) -> None:
        """Populate worksheet with table data."""
        for row_idx, row_data in enumerate(rows, 1):
            for col_idx, cell_value in enumerate(row_data, 1):
                if cell_value is not None:
                    worksheet.cell(row_idx, col_idx, cell_value)