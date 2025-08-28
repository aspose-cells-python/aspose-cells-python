"""
Markdown file writer for saving workbook data to Markdown table format.
"""

from typing import List, Optional, TYPE_CHECKING
from ...formats import CellValue

if TYPE_CHECKING:
    from ...workbook import Workbook
    from ...worksheet import Worksheet


class MarkdownWriter:
    """Writer for Markdown table files."""
    
    def __init__(self):
        pass
    
    def write(self, file_path: str, data: List[List[CellValue]], **kwargs) -> None:
        """Write data to Markdown file."""
        encoding = kwargs.get('encoding', 'utf-8')
        include_headers = kwargs.get('include_headers', True)
        table_alignment = kwargs.get('table_alignment', 'left')
        max_col_width = kwargs.get('max_col_width', 50)
        
        try:
            markdown_content = self._convert_data_to_markdown(
                data, include_headers, table_alignment, max_col_width
            )
            
            with open(file_path, 'w', encoding=encoding) as file:
                file.write(markdown_content)
                    
        except Exception as e:
            raise ValueError(f"Error writing Markdown file: {e}")
    
    def write_workbook(self, file_path: str, workbook: 'Workbook', **kwargs) -> None:
        """Write workbook data to Markdown file."""
        sheet_name = kwargs.get('sheet_name')
        include_headers = kwargs.get('include_headers', True)
        table_alignment = kwargs.get('table_alignment', 'left')
        max_col_width = kwargs.get('max_col_width', 50)
        all_sheets = kwargs.get('all_sheets', False)
        encoding = kwargs.get('encoding', 'utf-8')
        
        result_parts = []
        
        if sheet_name and sheet_name in workbook._worksheets:
            # Convert specific sheet
            worksheet = workbook._worksheets[sheet_name]
            sheet_md = self._convert_single_sheet(worksheet, include_headers, table_alignment, max_col_width)
            if sheet_md:
                result_parts.append(sheet_md)
        elif all_sheets:
            # Convert all sheets with headers
            for worksheet in workbook._worksheets.values():
                sheet_md = self._convert_single_sheet(worksheet, include_headers, table_alignment, max_col_width)
                if sheet_md:
                    result_parts.append(sheet_md)
                    result_parts.append("")  # Add empty line between sheets
        else:
            # Convert active sheet only
            worksheet = workbook.active
            sheet_md = self._convert_single_sheet(worksheet, include_headers, table_alignment, max_col_width)
            if sheet_md:
                result_parts.append(sheet_md)
        
        markdown_content = "\n".join(result_parts).strip()
        
        try:
            with open(file_path, 'w', encoding=encoding) as file:
                file.write(markdown_content)
        except Exception as e:
            raise ValueError(f"Error writing Markdown file: {e}")
    
    def _convert_single_sheet(self, worksheet: 'Worksheet', include_headers: bool, 
                             table_alignment: str, max_col_width: int) -> str:
        """Convert single worksheet to markdown with header."""
        if not worksheet or not worksheet._cells:
            return ""
        
        sheet_parts = []
        
        # Add worksheet title with markdown header
        sheet_parts.append(f"# {worksheet.name}")
        sheet_parts.append("")  # Empty line after header
        
        # Convert worksheet to data first
        data = self._worksheet_to_data(worksheet)
        # Heuristically trim leading non-tabular/title rows so header aligns with table
        if data:
            # Use improved detection logic to find the best starting row
            best_idx = self._detect_table_start_index(data)
            if best_idx > 0:
                data = data[best_idx:]
        
        # Convert data to markdown table
        table_md = self._convert_data_to_markdown(
            data, include_headers, table_alignment, max_col_width
        )
        if table_md:
            sheet_parts.append(table_md)
        
        return "\n".join(sheet_parts)
    
    def _worksheet_to_data(self, worksheet: 'Worksheet') -> List[List]:
        """Convert worksheet to list of rows with cell objects for hyperlink support."""
        max_row = worksheet.max_row
        max_col = worksheet.max_column
        
        if max_row == 0 or max_col == 0:
            return []
        
        # Collect all data with cell objects to preserve hyperlinks
        table_data = []
        for row in range(1, max_row + 1):
            row_data = []
            for col in range(1, max_col + 1):
                cell = worksheet._cells.get((row, col))
                if cell:
                    row_data.append(cell)  # Pass cell object instead of just value
                else:
                    row_data.append(None)
            table_data.append(row_data)
        
        return table_data
    
    def _convert_data_to_markdown(self, data: List[List], include_headers: bool, 
                                 alignment: str, max_width: int) -> str:
        """Convert data to markdown table."""
        if not data:
            return ""
        
        # Format all cell values (now handling cell objects)
        formatted_data = []
        for row_data in data:
            formatted_row = []
            for cell_or_value in row_data:
                formatted_cell = self._format_cell_for_markdown(cell_or_value, max_width)
                formatted_row.append(formatted_cell)
            formatted_data.append(formatted_row)
        
        if not formatted_data:
            return ""
        
        # Determine column widths
        max_col = max(len(row) for row in formatted_data) if formatted_data else 0
        col_widths = [0] * max_col
        
        for row_data in formatted_data:
            for i, cell_value in enumerate(row_data):
                if i < len(col_widths):
                    col_widths[i] = max(col_widths[i], len(str(cell_value)))
        
        # Generate markdown table
        result = []
        
        # Add header row if requested and data exists
        if include_headers and formatted_data:
            header_row = formatted_data[0]
            data_rows = formatted_data[1:]
            
            # Format header
            header_line = "| " + " | ".join(
                str(cell).ljust(col_widths[i]) for i, cell in enumerate(header_row)
            ) + " |"
            result.append(header_line)
            
            # Add separator line
            separator_parts = []
            align_char = self._get_alignment_chars(alignment)
            for width in col_widths:
                separator_parts.append(align_char[0] + "-" * max(1, width - 2) + align_char[1])
            separator_line = "| " + " | ".join(separator_parts) + " |"
            result.append(separator_line)
            
            # Add data rows
            for row_data in data_rows:
                formatted_row = [self._format_cell_for_markdown(cell, max_width) for cell in row_data]
                data_line = "| " + " | ".join(
                    str(cell).ljust(col_widths[i]) for i, cell in enumerate(formatted_row)
                ) + " |"
                result.append(data_line)
        else:
            # No headers, treat all as data
            for row_data in formatted_data:
                data_line = "| " + " | ".join(
                    str(cell).ljust(col_widths[i]) for i, cell in enumerate(row_data)
                ) + " |"
                result.append(data_line)
        
        return "\n".join(result)
    
    def _get_alignment_chars(self, alignment: str) -> tuple:
        """Get alignment characters for markdown table."""
        if alignment == 'center':
            return (":", ":")
        elif alignment == 'right':
            return ("-", ":")
        else:  # left
            return ("-", "-")
    
    def _format_cell_for_markdown(self, cell_or_value, max_width: int) -> str:
        """Format cell or value for markdown output with hyperlink support."""
        # Handle cell objects vs direct values
        if hasattr(cell_or_value, 'value') and hasattr(cell_or_value, 'hyperlink'):
            # This is a Cell object
            cell = cell_or_value
            value = cell.value
            hyperlink = cell.hyperlink
        else:
            # This is a direct value
            value = cell_or_value
            hyperlink = None
        
        if value is None:
            return ""
        
        # Convert to string
        if isinstance(value, bool):
            text = "TRUE" if value else "FALSE"
        else:
            text = str(value)
        
        # Escape markdown special characters
        text = text.replace("|", "\\|")
        text = text.replace("\n", " ")
        text = text.replace("\r", "")
        
        # Create hyperlink if present
        if hyperlink:
            # Escape hyperlink URL for markdown
            escaped_url = hyperlink.replace(")", "\\)")
            text = f"[{text}]({escaped_url})"
        
        # Truncate if too long (account for hyperlink syntax)
        if len(text) > max_width:
            if hyperlink:
                # For hyperlinks, try to preserve the link structure
                display_text = str(value)
                if len(display_text) > max_width - len(hyperlink) - 4:  # Account for []() syntax
                    display_text = display_text[:max_width - len(hyperlink) - 7] + "..."
                escaped_url = hyperlink.replace(")", "\\)")
                text = f"[{display_text}]({escaped_url})"
            else:
                text = text[:max_width - 3] + "..."
        
        return text
    
    def _detect_table_start_index(self, data: List[List]) -> int:
        """Detect the best starting index for the table data.
        
        Uses similar logic to the enhanced converter to skip rows with many "Unnamed" columns.
        """
        if not data:
            return 0
            
        best_idx = 0
        best_score = -1
        
        for idx, row in enumerate(data):
            score = self._score_row_as_table_start(row)
            if score > best_score:
                best_score = score
                best_idx = idx
        
        return best_idx
    
    def _score_row_as_table_start(self, row: List) -> float:
        """Score a row's likelihood of being the actual table start."""
        non_empty = 0
        unnamed_count = 0
        meaningful_content = 0
        total_chars = 0
        
        for cell_or_value in row:
            # Handle both cell objects and direct values
            if hasattr(cell_or_value, 'value'):
                value = cell_or_value.value
            else:
                value = cell_or_value
                
            if value is None:
                continue
                
            value_str = str(value).strip()
            if value_str == "":
                continue
                
            non_empty += 1
            total_chars += len(value_str)
            
            # Check for pandas-style "Unnamed" columns
            if value_str.startswith("Unnamed"):
                unnamed_count += 1
            else:
                meaningful_content += 1
        
        if non_empty == 0:
            return 0
        
        # Calculate score components
        unnamed_ratio = unnamed_count / non_empty if non_empty > 0 else 0
        meaningful_ratio = meaningful_content / non_empty if non_empty > 0 else 0
        avg_content_length = total_chars / non_empty if non_empty > 0 else 0
        
        score = 0
        
        # Penalize unnamed columns heavily
        if unnamed_ratio > 0.5:  # More than half are "Unnamed"
            score -= 100 * unnamed_ratio
        
        # Reward meaningful content
        score += 50 * meaningful_ratio
        
        # Reward reasonable content length
        if 2 <= avg_content_length <= 20:
            score += 20
        elif avg_content_length > 20:
            score += 10
        
        # Reward having multiple non-empty cells (but not too many unnamed ones)
        if non_empty >= 2 and unnamed_ratio < 0.5:
            score += min(non_empty * 5, 25)
        
        return score
    
    def _format_cell_value(self, value: CellValue, max_width: int) -> str:
        """Legacy method for backward compatibility."""
        return self._format_cell_for_markdown(value, max_width)
    
    def save_workbook(self, workbook: 'Workbook', file_path: str, **options) -> None:
        """Save workbook to Markdown file - unified interface method."""
        self.write_workbook(file_path, workbook, **options)