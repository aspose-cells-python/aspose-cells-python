"""Optimized Markdown converter for Excel data."""

from typing import List, Optional, TYPE_CHECKING
from datetime import datetime
from pathlib import Path

if TYPE_CHECKING:
    from ..workbook import Workbook
    from ..worksheet import Worksheet
    from ..cell import Cell


class MarkdownConverter:
    """Optimized Excel to Markdown converter."""

    def convert_workbook(self, workbook: 'Workbook', **kwargs) -> str:
        """Convert workbook to Markdown."""
        config = self._create_config(**kwargs)
        result_parts = []
        
        if config['include_metadata']:
            result_parts.extend([self._create_metadata(workbook), ""])
        
        sheets = self._get_sheets(workbook, config)
        for i, sheet in enumerate(sheets):
            if i > 0:
                result_parts.append("\n---\n" if config['include_metadata'] else "")
            content = self._process_sheet(sheet, config)
            if content:
                result_parts.append(content)
        
        return "\n".join(result_parts).strip()
    
    def _create_config(self, **kwargs) -> dict:
        """Create config with simplified defaults."""
        return {
            'sheet_name': kwargs.get('sheet_name'),
            'include_metadata': kwargs.get('include_metadata', False),
            'value_mode': kwargs.get('value_mode', 'value'),  # "value" shows calculated results, "formula" shows formulas
            'include_hyperlinks': kwargs.get('include_hyperlinks', True),
            'image_export_mode': kwargs.get('image_export_mode', 'none'),  # 'none', 'base64', 'folder'
            'image_folder': kwargs.get('image_folder', 'images'),
            'output_dir': kwargs.get('output_dir', '.')
        }
    
    def _get_sheets(self, workbook: 'Workbook', config: dict) -> List['Worksheet']:
        """Get sheets to process - specific sheet by name or all sheets."""
        sheet_name = config['sheet_name']
        if sheet_name and sheet_name in workbook._worksheets:
            # Convert specific sheet by name
            return [workbook._worksheets[sheet_name]]
        # Convert all sheets if no specific sheet is requested
        return list(workbook._worksheets.values())
    
    def _create_metadata(self, workbook: 'Workbook') -> str:
        """Create simplified metadata section without source file."""
        lines = [
            "# Document Metadata", "",
            f"- **Source Type**: Excel Workbook",
            f"- **Conversion Date**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"- **Total Sheets**: {len(workbook._worksheets)}",
            f"- **Sheet Names**: {', '.join(workbook._worksheets.keys())}",
            f"- **Active Sheet**: {workbook.active.name if workbook.active else 'None'}"
        ]
        return "\n".join(lines)
    
    def _process_sheet(self, worksheet: 'Worksheet', config: dict) -> str:
        """Process sheet to markdown with clean formatting."""
        if not worksheet or not worksheet._cells:
            return ""
        
        parts = [f"## {worksheet.name}", ""]
        
        # Add images if present and image export is enabled
        if hasattr(worksheet, '_images') and len(worksheet._images) > 0:
            image_content = self._process_images(worksheet, config)
            if image_content:
                parts.extend(["### Images", "", image_content, ""])
        
        table = self._create_table(worksheet, config)
        if table:
            parts.append(table)
        
        return "\n".join(parts)
    
    def _create_table(self, worksheet: 'Worksheet', config: dict) -> str:
        """Create markdown table."""
        if not worksheet._cells:
            return ""
        
        start_row = self._detect_start_row(worksheet)
        
        # Check for merged cells before start_row that should be included
        merged_title_rows = self._find_merged_title_rows(worksheet, start_row)
        
        all_data = []
        
        # Add merged title rows first
        for title_row in merged_title_rows:
            title_data = self._extract_data(worksheet, title_row, config, single_row=True)
            if title_data:
                all_data.extend(title_data)
        
        # Add main table data
        table_data = self._extract_data(worksheet, start_row, config)
        if table_data:
            all_data.extend(table_data)
        
        if not all_data:
            return ""
        
        result = []
        if all_data:
            # Always include headers with simplified logic
            header = all_data[0]
            header_line = "| " + " | ".join(self._generate_column_header(cell, i) for i, cell in enumerate(header)) + " |"
            separator = "| " + " | ".join("---" for _ in header) + " |"
            result.extend([header_line, separator])
            
            for row in all_data[1:]:
                data_line = "| " + " | ".join(str(cell) if cell else "" for cell in row) + " |"
                result.append(data_line)
        
        return "\n".join(result)
    
    def _find_merged_title_rows(self, worksheet: 'Worksheet', start_row: int) -> list:
        """Find merged cell rows before start_row that contain titles."""
        title_rows = []
        
        if hasattr(worksheet, '_merged_ranges') and worksheet._merged_ranges:
            for merge_range in worksheet._merged_ranges:
                if ':' in merge_range:
                    start_ref, end_ref = merge_range.split(':')
                    import re
                    merged_row = int(re.search(r'\d+', start_ref).group())
                    
                    # If merged row is before start_row and contains meaningful content
                    if merged_row < start_row:
                        cell = worksheet._cells.get((merged_row, 1))
                        if cell and cell.value and str(cell.value).strip():
                            title_rows.append(merged_row)
        
        return sorted(title_rows)
    
    def _extract_data(self, worksheet: 'Worksheet', start_row: int, config: dict, single_row: bool = False) -> List[List[str]]:
        """Extract table data."""
        data = []
        end_row = start_row if single_row else worksheet.max_row
        
        for row in range(start_row, end_row + 1):
            row_data = []
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet._cells.get((row, col))
                value = self._format_cell(cell, config)
                row_data.append(value)
            data.append(row_data)
        return data
    
    def _format_cell(self, cell: Optional['Cell'], config: dict) -> str:
        """Format cell value with enhanced processing."""
        if not cell or cell.is_empty():
            return ""
        
        # Enhanced hyperlink detection and formatting
        if config['include_hyperlinks'] and cell.has_hyperlink():
            return cell.get_markdown_link()
        
        # Auto-detect URLs in text values and convert to hyperlinks
        if config['include_hyperlinks'] and isinstance(cell.value, str):
            url_detected = self._detect_and_format_urls(cell.value)
            if url_detected != cell.value:
                return url_detected
        
        if config['value_mode'] == 'formula' and cell.is_formula():
            value = cell.formula or cell.value
        else:
            value = cell.calculated_value
        
        return self._format_value(value)
    
    def _detect_and_format_urls(self, text: str) -> str:
        """Detect URLs in text and format them as markdown links."""
        import re
        
        # Don't process text that already contains markdown links
        if '[' in text and '](' in text:
            return text
        
        # Patterns for different types of URLs (in order of specificity)
        url_patterns = [
            # Full HTTP/HTTPS URLs
            (r'\bhttps?://[^\s<>"\'|\[\]()]+', lambda m: f"[{m.group()}]({m.group()})"),
            # www domains
            (r'\bwww\.[^\s<>"\'|\[\]()]+', lambda m: f"[{m.group()}](http://{m.group()})"),
            # Email addresses
            (r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b', lambda m: f"[{m.group()}](mailto:{m.group()})")
        ]
        
        result = text
        for pattern, formatter in url_patterns:
            result = re.sub(pattern, formatter, result)
        
        return result
    
    def _format_value(self, value) -> str:
        """Format value for display with improved formatting."""
        if value is None:
            return ""
        
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        
        if isinstance(value, (int, float)):
            # Improve number formatting
            if isinstance(value, float):
                # Use scientific notation only for extremely large/small numbers
                if abs(value) >= 1e9 or (abs(value) < 1e-4 and value != 0):
                    return f"{value:.2e}"  # Scientific notation for very large/small numbers
                elif value.is_integer():
                    return str(int(value))  # Remove .0 from whole numbers
                else:
                    # Format large numbers with commas and reasonable decimal places
                    if abs(value) >= 1000:
                        return f"{value:,.2f}".rstrip('0').rstrip('.')
                    else:
                        return f"{value:.2f}".rstrip('0').rstrip('.')
            else:
                # Format large integers with commas
                if abs(value) >= 1000:
                    return f"{value:,}"
                else:
                    return str(value)
        
        if isinstance(value, str):
            # Enhanced string formatting
            text = value.replace("|", "\\|")  # Escape pipe characters
            text = text.replace("\n", " ").replace("\r", " ")  # Handle line breaks
            text = " ".join(text.split())  # Normalize whitespace
            return text.strip()
        
        # Fallback for other types
        return str(value).replace("|", "\\|").replace("\n", " ").strip()
    
    def _generate_column_header(self, cell_value: str, column_index: int) -> str:
        """Generate intelligent column headers."""
        if cell_value and str(cell_value).strip():
            header = str(cell_value).strip()
            # Don't use generic names if we have meaningful content
            if not header.startswith(('Unnamed', 'Col', 'Column')):
                return header
        
        # Generate Excel-style column names (A, B, C, ..., AA, AB, etc.)
        result = ""
        col_num = column_index
        while col_num >= 0:
            result = chr(col_num % 26 + ord('A')) + result
            col_num = col_num // 26 - 1
            if col_num < 0:
                break
        
        return result if result else f"Col{column_index + 1}"
    
    def _detect_start_row(self, worksheet: 'Worksheet') -> int:
        """Detect optimal table start row."""
        if worksheet.max_row <= 3:
            return 1
        
        best_row, best_score = 1, -1
        for row in range(1, min(worksheet.max_row + 1, 6)):
            score = self._score_row(worksheet, row)
            if score > best_score:
                best_score, best_row = score, row
        
        return best_row
    
    def _score_row(self, worksheet: 'Worksheet', row: int) -> float:
        """Score row as potential table start."""
        non_empty = meaningful = unnamed = 0
        merged_bonus = 0
        
        for col in range(1, worksheet.max_column + 1):
            cell = worksheet._cells.get((row, col))
            if not cell or cell.value is None:
                continue
            
            value_str = str(cell.value).strip()
            if not value_str:
                continue
            
            non_empty += 1
            if value_str.startswith("Unnamed"):
                unnamed += 1
            else:
                meaningful += 1
        
        # Check if this row contains merged cells
        if hasattr(worksheet, '_merged_ranges') and worksheet._merged_ranges:
            for merge_range in worksheet._merged_ranges:
                # Parse merge range like "A1:F1"
                if ':' in merge_range:
                    start_ref, end_ref = merge_range.split(':')
                    # Extract row number from references
                    import re
                    start_row = int(re.search(r'\d+', start_ref).group())
                    if start_row == row:
                        # This row has merged cells, give it a bonus
                        merged_bonus = 20
                        break
        
        if non_empty == 0:
            return 0
        
        score = 50 * (meaningful / non_empty) - 100 * (unnamed / non_empty) + merged_bonus
        if non_empty >= 2 and unnamed / non_empty < 0.5:
            score += min(non_empty * 5, 25)
        
        return score
    
    def _process_images(self, worksheet: 'Worksheet', config: dict) -> str:
        """Process images in worksheet based on export mode."""
        if config['image_export_mode'] == 'none':
            return ""
        
        image_lines = []
        
        for i, image in enumerate(worksheet._images):
            if config['image_export_mode'] == 'base64':
                # Export as base64 data URL
                image_md = self._image_to_base64_markdown(image, i)
            elif config['image_export_mode'] == 'folder':
                # Export to file and reference
                image_md = self._image_to_file_markdown(image, i, config)
            else:
                continue
            
            if image_md:
                image_lines.append(image_md)
        
        return "\n\n".join(image_lines)
    
    def _image_to_base64_markdown(self, image, index: int) -> str:
        """Convert image to base64 markdown."""
        import base64
        
        if not image.data:
            return f"*Image {index + 1}: {image.name or 'Unnamed'} (No data available)*"
        
        # Create base64 data URL
        format_str = image.format.value if hasattr(image.format, 'value') else str(image.format)
        base64_data = base64.b64encode(image.data).decode('utf-8')
        data_url = f"data:image/{format_str};base64,{base64_data}"
        
        # Create markdown
        alt_text = image.description or image.name or f"Image {index + 1}"
        anchor_info = self._get_anchor_description(image.anchor)
        
        md_lines = [
            f"**{image.name or f'Image {index + 1}'}**",
            f"- Position: {anchor_info}",
            f"- Size: {image.width}x{image.height}px" if image.width and image.height else "- Size: Unknown",
            f"- Format: {format_str.upper()}",
            f"- Description: {image.description}" if image.description else "",
            "",
            f"![{alt_text}]({data_url})"
        ]
        
        return "\n".join(line for line in md_lines if line)
    
    def _image_to_file_markdown(self, image, index: int, config: dict) -> str:
        """Convert image to file reference markdown."""
        import os
        
        if not image.data:
            return f"*Image {index + 1}: {image.name or 'Unnamed'} (No data available)*"
        
        # Create images directory
        output_dir = Path(config['output_dir'])
        image_dir = output_dir / config['image_folder']
        image_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate unique filename
        base_name = image.name or f"image_{index + 1}"
        format_ext = self._get_file_extension(image.format)
        filename = self._generate_unique_filename(image_dir, base_name, format_ext)
        
        # Save image file
        image_path = image_dir / filename
        with open(image_path, 'wb') as f:
            f.write(image.data)
        
        # Create markdown
        alt_text = image.description or image.name or f"Image {index + 1}"
        anchor_info = self._get_anchor_description(image.anchor)
        relative_path = f"{config['image_folder']}/{filename}"
        
        md_lines = [
            f"**{image.name or f'Image {index + 1}'}**",
            f"- Position: {anchor_info}",
            f"- Size: {image.width}x{image.height}px" if image.width and image.height else "- Size: Unknown",
            f"- Format: {image.format.value.upper() if hasattr(image.format, 'value') else str(image.format).upper()}",
            f"- Description: {image.description}" if image.description else "",
            f"- File: [{filename}]({relative_path})",
            "",
            f"![{alt_text}]({relative_path})"
        ]
        
        return "\n".join(line for line in md_lines if line)
    
    def _get_anchor_description(self, anchor) -> str:
        """Get human-readable anchor description."""
        from ..drawing.anchor import AnchorType
        
        if anchor.type == AnchorType.ONE_CELL:
            from ..utils.coordinates import tuple_to_coordinate
            cell_ref = tuple_to_coordinate(anchor.from_position[0] + 1, anchor.from_position[1] + 1)
            if anchor.from_offset != (0, 0):
                return f"Cell {cell_ref} + offset {anchor.from_offset}"
            return f"Cell {cell_ref}"
        elif anchor.type == AnchorType.TWO_CELL:
            from ..utils.coordinates import tuple_to_coordinate
            start_ref = tuple_to_coordinate(anchor.from_position[0] + 1, anchor.from_position[1] + 1)
            end_ref = tuple_to_coordinate(anchor.to_position[0] + 1, anchor.to_position[1] + 1)
            return f"Range {start_ref}:{end_ref}"
        elif anchor.type == AnchorType.ABSOLUTE:
            return f"Absolute ({anchor.absolute_position[0]}, {anchor.absolute_position[1]})"
        else:
            return "Unknown position"
    
    def _get_file_extension(self, image_format) -> str:
        """Get file extension for image format."""
        format_map = {
            'png': '.png',
            'jpeg': '.jpg',
            'jpg': '.jpg',
            'gif': '.gif'
        }
        format_str = image_format.value if hasattr(image_format, 'value') else str(image_format)
        return format_map.get(format_str.lower(), '.png')
    
    def _generate_unique_filename(self, directory: Path, base_name: str, extension: str) -> str:
        """Generate unique filename to avoid conflicts."""
        # Sanitize base name
        import re
        safe_name = re.sub(r'[^\w\-_.]', '_', base_name)
        
        filename = f"{safe_name}{extension}"
        counter = 1
        
        while (directory / filename).exists():
            filename = f"{safe_name}_{counter}{extension}"
            counter += 1
        
        return filename