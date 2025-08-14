"""Excel XLSX file writer with full OOXML implementation."""

import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Set, TYPE_CHECKING, Tuple
from pathlib import Path
import io

from ...utils import FileFormatError, tuple_to_coordinate
from .constants import XlsxConstants, XlsxTemplates

if TYPE_CHECKING:
    from ...workbook import Workbook, Worksheet
    from ...cell import Cell


class StyleManager:
    """Manages styles for Excel generation."""
    
    def __init__(self):
        self.fonts = []
        self.fills = []
        self.borders = []
        self.number_formats = {}
        self.cell_formats = []
        self.font_map = {}
        self.fill_map = {}
        self.border_map = {}
        self.format_map = {}
        
        # Initialize default styles
        self._init_default_styles()
    
    def _init_default_styles(self):
        """Initialize default Excel styles."""
        # Default font
        self.fonts.append(XlsxConstants.DEFAULT_FONT.copy())
        self.font_map[self._font_key(self.fonts[0])] = 0
        
        # Default fills
        self.fills.extend([fill.copy() for fill in XlsxConstants.DEFAULT_FILLS])
        self.fill_map[self._fill_key(self.fills[0])] = 0
        self.fill_map[self._fill_key(self.fills[1])] = 1
        
        # Default border
        self.borders.append(XlsxConstants.DEFAULT_BORDER.copy())
        self.border_map[self._border_key(self.borders[0])] = 0
        
        # Default cell format
        self.cell_formats.append(XlsxConstants.DEFAULT_CELL_FORMAT.copy())
    
    def _font_key(self, font):
        """Generate key for font lookup."""
        return f"{font['name']}|{font['size']}|{font['bold']}|{font['italic']}|{font['color']}"
    
    def _fill_key(self, fill):
        """Generate key for fill lookup."""
        return f"{fill['pattern']}|{fill.get('color', '')}"
    
    def _border_key(self, border):
        """Generate key for border lookup."""
        if hasattr(border, '_left'):
            # New border format
            left = f"{getattr(border._left, 'style', 'none')}:{getattr(border._left, 'color', 'black')}" if border._left else "none:black"
            right = f"{getattr(border._right, 'style', 'none')}:{getattr(border._right, 'color', 'black')}" if border._right else "none:black"
            top = f"{getattr(border._top, 'style', 'none')}:{getattr(border._top, 'color', 'black')}" if border._top else "none:black"
            bottom = f"{getattr(border._bottom, 'style', 'none')}:{getattr(border._bottom, 'color', 'black')}" if border._bottom else "none:black"
            return f"{left}|{right}|{top}|{bottom}"
        else:
            # Legacy border format
            return f"{border.get('left', '')}|{border.get('right', '')}|{border.get('top', '')}|{border.get('bottom', '')}"
    
    def get_font_id(self, font_props):
        """Get or create font ID."""
        font = {
            'name': getattr(font_props, 'name', 'Calibri'),
            'size': getattr(font_props, 'size', 11),
            'bold': getattr(font_props, 'bold', False),
            'italic': getattr(font_props, 'italic', False),
            'color': self._normalize_color(getattr(font_props, 'color', 'black'))
        }
        
        key = self._font_key(font)
        if key not in self.font_map:
            self.font_map[key] = len(self.fonts)
            self.fonts.append(font)
        
        return self.font_map[key]
    
    def get_fill_id(self, fill_props):
        """Get or create fill ID."""
        fill = {
            'pattern': 'solid',
            'color': self._normalize_color(getattr(fill_props, 'color', 'white'))
        }
        
        key = self._fill_key(fill)
        if key not in self.fill_map:
            self.fill_map[key] = len(self.fills)
            self.fills.append(fill)
        
        return self.fill_map[key]
    
    def get_border_id(self, border_props):
        """Get or create border ID."""
        if hasattr(border_props, '_left'):
            # New Border object
            border = {
                'left': getattr(border_props._left, 'style', 'none') if border_props._left else 'none',
                'left_color': getattr(border_props._left, 'color', 'black') if border_props._left else 'black',
                'right': getattr(border_props._right, 'style', 'none') if border_props._right else 'none',
                'right_color': getattr(border_props._right, 'color', 'black') if border_props._right else 'black',
                'top': getattr(border_props._top, 'style', 'none') if border_props._top else 'none',
                'top_color': getattr(border_props._top, 'color', 'black') if border_props._top else 'black',
                'bottom': getattr(border_props._bottom, 'style', 'none') if border_props._bottom else 'none',
                'bottom_color': getattr(border_props._bottom, 'color', 'black') if border_props._bottom else 'black'
            }
        else:
            # Legacy border
            border = {
                'left': 'none', 'left_color': 'black',
                'right': 'none', 'right_color': 'black',
                'top': 'none', 'top_color': 'black',
                'bottom': 'none', 'bottom_color': 'black'
            }
        
        key = self._border_key(border_props)
        if key not in self.border_map:
            self.border_map[key] = len(self.borders)
            self.borders.append(border)
        
        return self.border_map[key]
    
    def get_number_format_id(self, format_code):
        """Get or create number format ID."""
        if format_code in XlsxConstants.BUILTIN_NUMBER_FORMATS:
            return XlsxConstants.BUILTIN_NUMBER_FORMATS[format_code]
        
        # Custom format
        if format_code not in self.format_map:
            format_id = 164 + len(self.number_formats)
            self.number_formats[format_id] = format_code
            self.format_map[format_code] = format_id
        
        return self.format_map[format_code]
    
    def get_cell_format_id(self, cell):
        """Get cell format ID based on cell styling."""
        font_id = 0
        fill_id = 0
        border_id = 0
        number_format_id = 0
        
        if hasattr(cell, '_style') and cell._style:
            style = cell._style
            
            if style._font:
                font_id = self.get_font_id(style._font)
            
            if style._fill:
                fill_id = self.get_fill_id(style._fill)
            
            if style._border:
                border_id = self.get_border_id(style._border)
        
        if hasattr(cell, '_number_format') and cell._number_format != "General":
            number_format_id = self.get_number_format_id(cell._number_format)
        
        # Find or create cell format
        cell_format = {
            'font_id': font_id,
            'fill_id': fill_id,
            'border_id': border_id,
            'number_format_id': number_format_id
        }
        
        # Check if this format already exists
        for i, existing_format in enumerate(self.cell_formats):
            if existing_format == cell_format:
                return i
        
        # Create new format
        format_id = len(self.cell_formats)
        self.cell_formats.append(cell_format)
        return format_id
    
    def _normalize_color(self, color):
        """Normalize color to hex format."""
        if color in XlsxConstants.COLOR_MAP:
            return XlsxConstants.COLOR_MAP[color]
        elif color.startswith('#'):
            return color[1:].upper()
        elif len(color) == 6:
            return color.upper()
        else:
            return '000000'  # Default to black


class XlsxWriter:
    """Excel XLSX file writer with OOXML protocol support and proper styling."""
    
    def __init__(self):
        self.namespaces = XlsxConstants.NAMESPACES
        self.style_manager = StyleManager()
    
    def write(self, file_path: str, workbook: 'Workbook', **kwargs) -> None:
        """Write workbook to Excel XLSX file."""
        self.save_workbook(workbook, file_path, **kwargs)
    
    def save_workbook(self, workbook: 'Workbook', filename: str, **kwargs):
        """Save workbook to XLSX file with proper styling."""
        # Reset style manager for new file
        self.style_manager = StyleManager()
        
        with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Pre-process all cells to build styles
            self._analyze_styles(workbook)
            
            # Build shared strings
            shared_strings = self._build_shared_strings(workbook)
            
            # Write core structure files
            self._write_content_types(zip_file, workbook, bool(shared_strings))
            self._write_rels(zip_file)
            self._write_app_properties(zip_file)
            self._write_core_properties(zip_file)
            self._write_workbook_xml(zip_file, workbook)
            self._write_workbook_rels(zip_file, workbook)
            
            # Write shared strings only if they exist
            if shared_strings:
                self._write_shared_strings(zip_file, shared_strings)
            
            # Write styles with proper formatting
            self._write_styles(zip_file)
            
            # Write theme
            self._write_theme(zip_file)
            
            # Write worksheets
            for idx, worksheet in enumerate(workbook._worksheets.values(), 1):
                self._write_worksheet(zip_file, worksheet, idx, shared_strings)
    
    def _analyze_styles(self, workbook: 'Workbook'):
        """Pre-analyze all cells to build style tables."""
        for worksheet in workbook._worksheets.values():
            for cell in worksheet._cells.values():
                if cell.value is not None:
                    # This will register the style
                    self.style_manager.get_cell_format_id(cell)
    
    def _build_shared_strings(self, workbook: 'Workbook') -> Dict[str, int]:
        """Build shared strings table."""
        strings = {}
        string_list = []
        
        for worksheet in workbook._worksheets.values():
            for cell in worksheet._cells.values():
                if isinstance(cell.value, str) and not cell.is_formula():
                    if cell.value not in strings:
                        strings[cell.value] = len(string_list)
                        string_list.append(cell.value)
        
        return strings
    
    def _write_content_types(self, zip_file: zipfile.ZipFile, workbook: 'Workbook', has_shared_strings: bool = True):
        """Write [Content_Types].xml with proper formatting."""
        root = ET.Element("Types")
        root.set("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
        
        # Default content types
        for ext, content_type in XlsxConstants.CONTENT_TYPES['defaults']:
            default = ET.SubElement(root, "Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
        
        # Override content types
        overrides = [
            XlsxConstants.CONTENT_TYPES['overrides']['workbook'],
            XlsxConstants.CONTENT_TYPES['overrides']['styles'],
            XlsxConstants.CONTENT_TYPES['overrides']['theme'],
            XlsxConstants.CONTENT_TYPES['overrides']['core_props'],
            XlsxConstants.CONTENT_TYPES['overrides']['app_props']
        ]
        
        # Only add shared strings if they exist
        if has_shared_strings:
            overrides.append(XlsxConstants.CONTENT_TYPES['overrides']['shared_strings'])
        
        # Add worksheet overrides
        for idx in range(1, len(workbook._worksheets) + 1):
            overrides.append((
                f"/xl/worksheets/sheet{idx}.xml",
                XlsxConstants.CONTENT_TYPES['overrides']['worksheet']
            ))
        
        for part_name, content_type in overrides:
            override = ET.SubElement(root, "Override")
            override.set("PartName", part_name)
            override.set("ContentType", content_type)
        
        self._write_xml_to_zip(zip_file, "[Content_Types].xml", root)
    
    def _write_rels(self, zip_file: zipfile.ZipFile):
        """Write _rels/.rels."""
        root = ET.Element("Relationships")
        root.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        for rel_id, rel_type, target in XlsxTemplates.get_rels_data():
            rel = ET.SubElement(root, "Relationship")
            rel.set("Id", rel_id)
            rel.set("Type", rel_type)
            rel.set("Target", target)
        
        self._write_xml_to_zip(zip_file, "_rels/.rels", root)
    
    def _write_workbook_xml(self, zip_file: zipfile.ZipFile, workbook: 'Workbook'):
        """Write xl/workbook.xml."""
        root = ET.Element("workbook")
        root.set("xmlns", self.namespaces['main'])
        root.set("xmlns:r", self.namespaces['r'])
        
        # Book views
        book_views = ET.SubElement(root, "bookViews")
        work_book_view = ET.SubElement(book_views, "workbookView")
        work_book_view.set("xWindow", XlsxConstants.SHEET_DEFAULTS['window_x'])
        work_book_view.set("yWindow", XlsxConstants.SHEET_DEFAULTS['window_y'])
        work_book_view.set("windowWidth", XlsxConstants.SHEET_DEFAULTS['window_width'])
        work_book_view.set("windowHeight", XlsxConstants.SHEET_DEFAULTS['window_height'])
        
        # Sheets
        sheets = ET.SubElement(root, "sheets")
        for idx, (name, worksheet) in enumerate(workbook._worksheets.items(), 1):
            sheet = ET.SubElement(sheets, "sheet")
            sheet.set("name", name)
            sheet.set("sheetId", str(idx))
            sheet.set("r:id", f"rId{idx}")
        
        self._write_xml_to_zip(zip_file, "xl/workbook.xml", root)
    
    def _write_workbook_rels(self, zip_file: zipfile.ZipFile, workbook: 'Workbook'):
        """Write xl/_rels/workbook.xml.rels."""
        root = ET.Element("Relationships")
        root.set("xmlns", self.namespaces['pkg'])
        
        rel_id = 1
        
        # Worksheet relationships
        for idx in range(1, len(workbook._worksheets) + 1):
            rel = ET.SubElement(root, "Relationship")
            rel.set("Id", f"rId{rel_id}")
            rel.set("Type", XlsxConstants.REL_TYPES['worksheet'])
            rel.set("Target", f"worksheets/sheet{idx}.xml")
            rel_id += 1
        
        # Styles relationship
        styles_rel = ET.SubElement(root, "Relationship")
        styles_rel.set("Id", f"rId{rel_id}")
        styles_rel.set("Type", XlsxConstants.REL_TYPES['styles'])
        styles_rel.set("Target", "styles.xml")
        rel_id += 1
        
        # Theme relationship
        theme_rel = ET.SubElement(root, "Relationship")
        theme_rel.set("Id", f"rId{rel_id}")
        theme_rel.set("Type", XlsxConstants.REL_TYPES['theme'])
        theme_rel.set("Target", "theme/theme1.xml")
        rel_id += 1
        
        # Only add shared strings relationship if there are shared strings
        shared_strings = self._build_shared_strings(workbook)
        if shared_strings:
            shared_rel = ET.SubElement(root, "Relationship")
            shared_rel.set("Id", f"rId{rel_id}")
            shared_rel.set("Type", XlsxConstants.REL_TYPES['shared_strings'])
            shared_rel.set("Target", "sharedStrings.xml")
        
        self._write_xml_to_zip(zip_file, "xl/_rels/workbook.xml.rels", root)
    
    def _write_worksheet_rels(self, zip_file: zipfile.ZipFile, sheet_id: int, hyperlinks: list):
        """Write worksheet relationships for hyperlinks."""
        if not hyperlinks:
            return
            
        root = ET.Element("Relationships")
        root.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        for idx, cell in enumerate(hyperlinks, 1):
            relationship = ET.SubElement(root, "Relationship")
            relationship.set("Id", f"rId{idx}")
            relationship.set("Type", XlsxConstants.REL_TYPES['hyperlink'])
            relationship.set("Target", cell.hyperlink)
            relationship.set("TargetMode", "External")
        
        self._write_xml_to_zip(zip_file, f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels", root)
    
    def _write_worksheet(self, zip_file: zipfile.ZipFile, worksheet: 'Worksheet', 
                        sheet_id: int, shared_strings: Dict[str, int]):
        """Write individual worksheet XML with proper styling."""
        root = ET.Element("worksheet")
        root.set("xmlns", self.namespaces['main'])
        root.set("xmlns:r", self.namespaces['r'])
        
        # Sheet views
        sheet_views = ET.SubElement(root, "sheetViews")
        sheet_view = ET.SubElement(sheet_views, "sheetView")
        sheet_view.set("tabSelected", "1" if worksheet == worksheet._parent.active else "0")
        sheet_view.set("workbookViewId", "0")
        
        # Sheet format properties
        sheet_format = ET.SubElement(root, "sheetFormatPr")
        sheet_format.set("defaultRowHeight", XlsxConstants.SHEET_DEFAULTS['default_row_height'])
        sheet_format.set("defaultColWidth", XlsxConstants.SHEET_DEFAULTS['default_col_width'])
        
        # Column widths
        if worksheet._column_widths:
            cols = ET.SubElement(root, "cols")
            for col_num, width in sorted(worksheet._column_widths.items()):
                col = ET.SubElement(cols, "col")
                col.set("min", str(col_num))
                col.set("max", str(col_num))
                col.set("width", str(width))
                col.set("customWidth", "1")
        
        # Sheet data (always required, even for empty worksheets)
        sheet_data = ET.SubElement(root, "sheetData")
        
        if worksheet._cells:
            # Group cells by row
            rows_data = {}
            for (row, col), cell in worksheet._cells.items():
                if row not in rows_data:
                    rows_data[row] = {}
                rows_data[row][col] = cell
            
            # Write rows
            for row_num in sorted(rows_data.keys()):
                row_elem = ET.SubElement(sheet_data, "row")
                row_elem.set("r", str(row_num))
                
                # Add custom row height if set
                if row_num in worksheet._row_heights:
                    row_elem.set("ht", str(worksheet._row_heights[row_num]))
                    row_elem.set("customHeight", "1")
                
                for col_num in sorted(rows_data[row_num].keys()):
                    cell = rows_data[row_num][col_num]
                    if cell.value is not None:
                        self._write_cell(row_elem, cell, shared_strings)
        
        # Merged cells
        if worksheet._merged_ranges:
            merge_cells = ET.SubElement(root, "mergeCells")
            merge_cells.set("count", str(len(worksheet._merged_ranges)))
            for range_ref in worksheet._merged_ranges:
                merge_cell = ET.SubElement(merge_cells, "mergeCell")
                merge_cell.set("ref", range_ref)
        
        # Hyperlinks
        hyperlinks = []
        if worksheet._cells:
            for cell in worksheet._cells.values():
                if cell.has_hyperlink():
                    hyperlinks.append(cell)
        
        if hyperlinks:
            hyperlinks_elem = ET.SubElement(root, "hyperlinks")
            for idx, cell in enumerate(hyperlinks, 1):
                hyperlink = ET.SubElement(hyperlinks_elem, "hyperlink")
                hyperlink.set("ref", cell.coordinate)
                hyperlink.set("r:id", f"rId{idx}")
            # Create worksheet relationships for hyperlinks
            self._write_worksheet_rels(zip_file, sheet_id, hyperlinks)
        
        self._write_xml_to_zip(zip_file, f"xl/worksheets/sheet{sheet_id}.xml", root)
    
    def _write_cell(self, row_elem: ET.Element, cell, shared_strings: Dict[str, int]):
        """Write individual cell element with proper styling."""
        cell_elem = ET.SubElement(row_elem, "c")
        cell_elem.set("r", cell.coordinate)
        
        # Apply style
        style_id = self.style_manager.get_cell_format_id(cell)
        if style_id > 0:  # Only set if not default style
            cell_elem.set("s", str(style_id))
        
        value = cell.value
        if isinstance(value, str) and not cell.is_formula():
            # Use shared strings for non-formula strings
            if value in shared_strings:
                cell_elem.set("t", "s")
                v_elem = ET.SubElement(cell_elem, "v")
                v_elem.text = str(shared_strings[value])
            else:
                cell_elem.set("t", "inlineStr")
                is_elem = ET.SubElement(cell_elem, "is")
                t_elem = ET.SubElement(is_elem, "t")
                t_elem.text = value
        elif isinstance(value, bool):
            cell_elem.set("t", "b")
            v_elem = ET.SubElement(cell_elem, "v")
            v_elem.text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            v_elem = ET.SubElement(cell_elem, "v")
            v_elem.text = str(value)
        elif cell.is_formula():
            # Write formula
            f_elem = ET.SubElement(cell_elem, "f")
            f_elem.text = str(value)[1:]  # Remove = prefix
            
            # Always write calculated value for formulas
            calc_value = None
            if hasattr(cell, '_calculated_value') and cell._calculated_value is not None:
                calc_value = cell._calculated_value
            elif hasattr(cell, 'calculated_value') and cell.calculated_value is not None:
                calc_value = cell.calculated_value
            else:
                # Provide fallback calculated value to ensure Excel can display something
                calc_value = self._get_fallback_formula_value(str(value))
            
            # Write the calculated value
            if calc_value is not None:
                if isinstance(calc_value, bool):
                    cell_elem.set("t", "b")
                    v_elem = ET.SubElement(cell_elem, "v")
                    v_elem.text = "1" if calc_value else "0"
                elif isinstance(calc_value, (int, float)):
                    v_elem = ET.SubElement(cell_elem, "v")
                    v_elem.text = str(calc_value)
                elif isinstance(calc_value, str):
                    # String result from formula
                    if calc_value in shared_strings:
                        cell_elem.set("t", "s")
                        v_elem = ET.SubElement(cell_elem, "v")
                        v_elem.text = str(shared_strings[calc_value])
                    else:
                        cell_elem.set("t", "str")  # Formula string result
                        v_elem = ET.SubElement(cell_elem, "v")
                        v_elem.text = calc_value
                else:
                    # Fallback to string representation
                    v_elem = ET.SubElement(cell_elem, "v")
                    v_elem.text = str(calc_value)
        else:
            # Fallback to string
            v_elem = ET.SubElement(cell_elem, "v")
            v_elem.text = str(value) if value is not None else ""
    
    def _get_fallback_formula_value(self, formula: str):
        """Provide basic fallback calculated values for common formulas."""
        # This method should rarely be used since Cell class now handles this better
        # Keep it simple for edge cases
        formula_upper = formula.upper().strip()
        
        # Remove = prefix if present
        if formula_upper.startswith('='):
            formula_upper = formula_upper[1:]
        
        # Handle simple cases
        if formula_upper.startswith(('SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN')):
            return 0
        elif formula_upper.startswith(('NOW', 'TODAY')):
            return "2024-01-01"
        elif formula_upper.startswith('TRUE'):
            return True
        elif formula_upper.startswith('FALSE'):
            return False
        elif formula_upper.startswith(('CONCATENATE', 'TEXT')):
            return ""
        elif all(c in '0123456789+-*/.() ' for c in formula_upper):
            # Pure numeric formula - use safe expression evaluation
            try:
                import ast
                node = ast.parse(formula_upper, mode='eval')
                if self._is_safe_expression(node):
                    return eval(compile(node, '<string>', 'eval'))
                else:
                    return 0
            except (ValueError, SyntaxError, TypeError):
                return 0
        else:
            # Default - let Excel handle it
            return 0
    
    def _write_shared_strings(self, zip_file: zipfile.ZipFile, shared_strings: Dict[str, int]):
        """Write xl/sharedStrings.xml."""
        root = ET.Element("sst")
        root.set("xmlns", self.namespaces['main'])
        root.set("count", str(len(shared_strings)))
        root.set("uniqueCount", str(len(shared_strings)))
        
        # Sort by index to maintain order
        sorted_strings = sorted(shared_strings.items(), key=lambda x: x[1])
        
        for string_value, _ in sorted_strings:
            si = ET.SubElement(root, "si")
            t = ET.SubElement(si, "t")
            t.text = string_value
        
        self._write_xml_to_zip(zip_file, "xl/sharedStrings.xml", root)
    
    def _write_styles(self, zip_file: zipfile.ZipFile):
        """Write comprehensive xl/styles.xml with all styling information."""
        root = ET.Element("styleSheet")
        root.set("xmlns", self.namespaces['main'])
        
        # Number formats (custom formats only)
        if self.style_manager.number_formats:
            num_fmts = ET.SubElement(root, "numFmts")
            num_fmts.set("count", str(len(self.style_manager.number_formats)))
            
            for format_id, format_code in self.style_manager.number_formats.items():
                num_fmt = ET.SubElement(num_fmts, "numFmt")
                num_fmt.set("numFmtId", str(format_id))
                num_fmt.set("formatCode", format_code)
        
        # Fonts
        fonts = ET.SubElement(root, "fonts")
        fonts.set("count", str(len(self.style_manager.fonts)))
        
        for font in self.style_manager.fonts:
            font_elem = ET.SubElement(fonts, "font")
            
            # Font size
            sz = ET.SubElement(font_elem, "sz")
            sz.set("val", str(font['size']))
            
            # Font color
            if font['color'] != '000000':  # Only add if not black (default)
                color = ET.SubElement(font_elem, "color")
                color.set("rgb", f"FF{font['color']}")  # Add alpha channel
            
            # Font name
            name = ET.SubElement(font_elem, "name")
            name.set("val", font['name'])
            
            # Font attributes
            if font['bold']:
                ET.SubElement(font_elem, "b")
            if font['italic']:
                ET.SubElement(font_elem, "i")
        
        # Fills
        fills = ET.SubElement(root, "fills")
        fills.set("count", str(len(self.style_manager.fills)))
        
        for fill in self.style_manager.fills:
            fill_elem = ET.SubElement(fills, "fill")
            pattern_fill = ET.SubElement(fill_elem, "patternFill")
            pattern_fill.set("patternType", fill['pattern'])
            
            if fill['color'] and fill['pattern'] == 'solid':
                fg_color = ET.SubElement(pattern_fill, "fgColor")
                fg_color.set("rgb", f"FF{fill['color']}")
        
        # Borders
        borders = ET.SubElement(root, "borders")
        borders.set("count", str(len(self.style_manager.borders)))
        
        for border in self.style_manager.borders:
            border_elem = ET.SubElement(borders, "border")
            
            # Write each border side with style and color
            for side in ["left", "right", "top", "bottom", "diagonal"]:
                side_elem = ET.SubElement(border_elem, side)
                
                side_style = border.get(side, 'none')
                side_color = border.get(f'{side}_color', 'black')
                
                if side_style and side_style != 'none':
                    side_elem.set("style", side_style)
                    if side_color and side_color != 'black':
                        color_elem = ET.SubElement(side_elem, "color")
                        color_elem.set("rgb", f"FF{self.style_manager._normalize_color(side_color)}")
        
        # Cell style formats (cellStyleXfs)
        cell_style_xfs = ET.SubElement(root, "cellStyleXfs")
        cell_style_xfs.set("count", "1")
        xf = ET.SubElement(cell_style_xfs, "xf")
        xf.set("numFmtId", "0")
        xf.set("fontId", "0")
        xf.set("fillId", "0")
        xf.set("borderId", "0")
        
        # Cell formats (cellXfs)
        cell_xfs = ET.SubElement(root, "cellXfs")
        cell_xfs.set("count", str(len(self.style_manager.cell_formats)))
        
        for cell_format in self.style_manager.cell_formats:
            xf = ET.SubElement(cell_xfs, "xf")
            xf.set("numFmtId", str(cell_format['number_format_id']))
            xf.set("fontId", str(cell_format['font_id']))
            xf.set("fillId", str(cell_format['fill_id']))
            xf.set("borderId", str(cell_format['border_id']))
            xf.set("xfId", "0")
            
            # Apply formatting flags
            if cell_format['font_id'] > 0:
                xf.set("applyFont", "1")
            if cell_format['fill_id'] > 0:
                xf.set("applyFill", "1")
            if cell_format['border_id'] > 0:
                xf.set("applyBorder", "1")
            if cell_format['number_format_id'] > 0:
                xf.set("applyNumberFormat", "1")
        
        # Cell styles
        cell_styles = ET.SubElement(root, "cellStyles")
        cell_styles.set("count", "1")
        cell_style = ET.SubElement(cell_styles, "cellStyle")
        cell_style.set("name", "Normal")
        cell_style.set("xfId", "0")
        cell_style.set("builtinId", "0")
        
        self._write_xml_to_zip(zip_file, "xl/styles.xml", root)
    
    def _write_theme(self, zip_file: zipfile.ZipFile):
        """Write xl/theme/theme1.xml with comprehensive theme."""
        zip_file.writestr("xl/theme/theme1.xml", XlsxTemplates.get_theme_xml())
    
    def _write_app_properties(self, zip_file: zipfile.ZipFile):
        """Write docProps/app.xml."""
        root = ET.Element("Properties")
        root.set("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
        root.set("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
        
        app_data = XlsxTemplates.get_app_properties_data()
        
        app = ET.SubElement(root, "Application")
        app.text = app_data['application']
        
        doc_security = ET.SubElement(root, "DocSecurity")
        doc_security.text = app_data['doc_security']
        
        lines_of_text = ET.SubElement(root, "LinksUpToDate")
        lines_of_text.text = app_data['links_up_to_date']
        
        shared_doc = ET.SubElement(root, "SharedDoc")
        shared_doc.text = app_data['shared_doc']
        
        hyperlinks_changed = ET.SubElement(root, "HyperlinksChanged")
        hyperlinks_changed.text = app_data['hyperlinks_changed']
        
        app_version = ET.SubElement(root, "AppVersion")
        app_version.text = app_data['app_version']
        
        self._write_xml_to_zip(zip_file, "docProps/app.xml", root)
    
    def _write_core_properties(self, zip_file: zipfile.ZipFile):
        """Write docProps/core.xml."""
        root = ET.Element("cp:coreProperties")
        root.set("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
        root.set("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        root.set("xmlns:dcterms", "http://purl.org/dc/terms/")
        root.set("xmlns:dcmitype", "http://purl.org/dc/dcmitype/")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        
        core_data = XlsxTemplates.get_core_properties_data()
        
        creator = ET.SubElement(root, "dc:creator")
        creator.text = core_data['creator']
        
        last_modified = ET.SubElement(root, "dcterms:modified")
        last_modified.set("xsi:type", "dcterms:W3CDTF")
        last_modified.text = core_data['modified']
        
        created = ET.SubElement(root, "dcterms:created")
        created.set("xsi:type", "dcterms:W3CDTF")
        created.text = core_data['created']
        
        self._write_xml_to_zip(zip_file, "docProps/core.xml", root)
    
    def _is_safe_expression(self, node) -> bool:
        """Check if AST node contains only safe mathematical operations."""
        import ast
        
        allowed_nodes = (
            ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Num,
            ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod, ast.Pow,
            ast.USub, ast.UAdd
        )
        
        for child in ast.walk(node):
            if not isinstance(child, allowed_nodes):
                return False
        return True
    
    def _write_xml_to_zip(self, zip_file: zipfile.ZipFile, path: str, root: ET.Element):
        """Write XML element to ZIP file with proper formatting."""
        self._indent_xml(root)
        xml_str = ET.tostring(root, encoding='utf-8', xml_declaration=True).decode('utf-8')
        zip_file.writestr(path, xml_str)
    
    def _indent_xml(self, elem, level=0):
        """Add proper indentation to XML for readability."""
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for child in elem:
                self._indent_xml(child, level + 1)
            if not child.tail or not child.tail.strip():
                child.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i