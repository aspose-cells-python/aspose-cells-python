"""Excel XLSX file reader with full OOXML implementation."""

import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, TYPE_CHECKING
from pathlib import Path

from ...utils import FileFormatError, coordinate_to_tuple
from .constants import XlsxConstants

if TYPE_CHECKING:
    from ...workbook import Workbook
    from ...worksheet import Worksheet


class XlsxReader:
    """Excel XLSX file reader with OOXML protocol support."""
    
    def __init__(self):
        self.namespaces = XlsxConstants.NAMESPACES
    
    def read(self, file_path: str, **kwargs) -> 'Workbook':
        """Read Excel file and return workbook object."""
        from ...workbook import Workbook
        
        workbook = Workbook()
        self.load_workbook(workbook, file_path)
        return workbook
    
    def load_workbook(self, workbook: 'Workbook', filename: str):
        """Load Excel file into workbook object."""
        try:
            with zipfile.ZipFile(filename, 'r') as zip_file:
                # Read core files
                shared_strings = self._read_shared_strings(zip_file)
                workbook_data = self._read_workbook_structure(zip_file)
                relationships = self._read_workbook_relationships(zip_file)
                
                # Clear existing worksheets
                workbook._worksheets.clear()
                workbook._shared_strings = shared_strings
                
                # Load worksheets with proper relationship mapping
                for sheet_info in workbook_data['sheets']:
                    self._load_worksheet(zip_file, workbook, sheet_info, shared_strings, relationships)
                
                # Set active sheet
                if workbook._worksheets:
                    first_sheet = next(iter(workbook._worksheets.values()))
                    workbook._active_sheet = first_sheet
                
        except zipfile.BadZipFile:
            raise FileFormatError(f"Invalid ZIP file: {filename}")
        except Exception as e:
            raise FileFormatError(f"Failed to read Excel file: {e}")
    
    def _read_shared_strings(self, zip_file: zipfile.ZipFile) -> List[str]:
        """Read shared strings table."""
        try:
            content = zip_file.read('xl/sharedStrings.xml')
            root = ET.fromstring(content)
            
            strings = []
            for si in root.findall('.//main:si', self.namespaces):
                t_elem = si.find('main:t', self.namespaces)
                if t_elem is not None:
                    strings.append(t_elem.text or "")
                else:
                    strings.append("")
            
            return strings
        except KeyError:
            # No shared strings file
            return []
    
    def _read_workbook_structure(self, zip_file: zipfile.ZipFile) -> Dict:
        """Read workbook structure and sheet information."""
        try:
            content = zip_file.read('xl/workbook.xml')
            root = ET.fromstring(content)
            
            sheets = []
            for sheet in root.findall('.//main:sheet', self.namespaces):
                sheet_info = {
                    'name': sheet.get('name', 'Sheet1'),
                    'sheet_id': sheet.get('sheetId', '1'),
                    'r_id': sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                }
                sheets.append(sheet_info)
            
            return {'sheets': sheets}
        except KeyError:
            raise FileFormatError("Invalid workbook structure - missing workbook.xml")
    
    def _read_workbook_relationships(self, zip_file: zipfile.ZipFile) -> Dict[str, str]:
        """Read workbook relationships to map sheet IDs to file paths."""
        try:
            content = zip_file.read('xl/_rels/workbook.xml.rels')
            root = ET.fromstring(content)
            
            relationships = {}
            for rel in root.findall('rel:Relationship', {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                if rel_id and target:
                    relationships[rel_id] = target
            
            return relationships
        except KeyError:
            # No relationships file, return empty dict
            return {}
    
    def _load_worksheet(self, zip_file: zipfile.ZipFile, workbook: 'Workbook', 
                       sheet_info: Dict, shared_strings: List[str], relationships: Dict[str, str]):
        """Load individual worksheet data."""
        from ...worksheet import Worksheet
        
        # Create worksheet
        worksheet = Worksheet(workbook, sheet_info['name'])
        workbook._worksheets[sheet_info['name']] = worksheet
        
        # Read worksheet XML
        try:
            # Determine worksheet path using relationships
            sheet_path = None
            r_id = sheet_info.get('r_id')
            if r_id and r_id in relationships:
                sheet_path = f"xl/{relationships[r_id]}"
            
            # Fallback to naming convention
            if not sheet_path or sheet_path not in zip_file.namelist():
                sheet_path = f"xl/worksheets/sheet{sheet_info['sheet_id']}.xml"
            
            # Final fallback - but don't use sheet1.xml for all sheets!
            if sheet_path not in zip_file.namelist():
                # Skip this sheet if we can't find its file
                return
            
            content = zip_file.read(sheet_path)
            root = ET.fromstring(content)
            
            # Process sheet data
            sheet_data = root.find('.//main:sheetData', self.namespaces)
            if sheet_data is not None:
                self._process_sheet_data(worksheet, sheet_data, shared_strings)
            
            # Process merged cells
            merge_cells = root.find('.//main:mergeCells', self.namespaces)
            if merge_cells is not None:
                for merge_cell in merge_cells.findall('main:mergeCell', self.namespaces):
                    ref = merge_cell.get('ref')
                    if ref:
                        worksheet._merged_ranges.add(ref)
            
            # Process hyperlinks
            self._process_hyperlinks(zip_file, worksheet, root, sheet_info['sheet_id'])
        
        except KeyError:
            # Worksheet file not found, create empty worksheet
            pass
    
    def _process_sheet_data(self, worksheet: 'Worksheet', sheet_data: ET.Element, 
                           shared_strings: List[str]):
        """Process sheet data and populate cells."""
        for row in sheet_data.findall('main:row', self.namespaces):
            for cell_elem in row.findall('main:c', self.namespaces):
                # Get cell reference
                cell_ref = cell_elem.get('r')
                if not cell_ref:
                    continue
                
                try:
                    row_idx, col_idx = coordinate_to_tuple(cell_ref)
                except (ValueError, TypeError, AttributeError):
                    # Skip invalid cell references
                    continue
                
                # Get cell value and formula
                cell_type = cell_elem.get('t', 'n')  # Default to number
                value_elem = cell_elem.find('main:v', self.namespaces)
                formula_elem = cell_elem.find('main:f', self.namespaces)
                
                # Create cell first
                cell = worksheet.cell(row_idx, col_idx)
                
                # Handle formula if present
                if formula_elem is not None:
                    formula_text = formula_elem.text
                    if formula_text:
                        # Store formula
                        cell._formula = '=' + formula_text if not formula_text.startswith('=') else formula_text
                        cell._data_type = 'formula'
                        cell._value = cell._formula
                
                # Handle calculated value
                if value_elem is not None:
                    raw_value = value_elem.text or ""
                    calculated_value = self._parse_cell_value(raw_value, cell_type, shared_strings)
                    
                    if cell.is_formula():
                        # Store calculated result for formula cells
                        cell._calculated_value = calculated_value
                    else:
                        # Regular cell value
                        cell.value = calculated_value
                
                # Handle hyperlinks (basic implementation)
                # Note: Full hyperlink support would require reading relationships
                
                # Handle number format if present
                style_id = cell_elem.get('s')
                if style_id:
                    # In a full implementation, would look up style from styles.xml
                    pass
    
    def _process_hyperlinks(self, zip_file: zipfile.ZipFile, worksheet: 'Worksheet', 
                           worksheet_root: ET.Element, sheet_id: int):
        """Process hyperlinks for the worksheet."""
        # Find hyperlinks in the worksheet XML
        hyperlinks_elem = worksheet_root.find('.//main:hyperlinks', self.namespaces)
        if hyperlinks_elem is None:
            return
        
        # Read worksheet relationships to get hyperlink targets
        rels_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
        relationships = {}
        
        try:
            rels_content = zip_file.read(rels_path).decode('utf-8')
            rels_root = ET.fromstring(rels_content)
            
            # Build relationships map
            # The relationships XML uses the package relationships namespace as default
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                if rel_id and target:
                    relationships[rel_id] = target
        except KeyError:
            # No relationships file found
            return
        
        # Apply hyperlinks to cells
        for hyperlink in hyperlinks_elem.findall('main:hyperlink', self.namespaces):
            cell_ref = hyperlink.get('ref')
            # Get the relationship ID using the proper namespace
            rel_id = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            
            if cell_ref and rel_id and rel_id in relationships:
                try:
                    row_idx, col_idx = coordinate_to_tuple(cell_ref)
                    cell = worksheet.cell(row_idx, col_idx)
                    cell._hyperlink = relationships[rel_id]
                except (ValueError, TypeError, AttributeError, KeyError):
                    # Skip invalid cell references or missing relationships
                    continue
    
    def _parse_cell_value(self, raw_value: str, cell_type: str, shared_strings: List[str]):
        """Parse cell value based on type."""
        if cell_type == 's':  # Shared string
            try:
                index = int(raw_value)
                if 0 <= index < len(shared_strings):
                    return shared_strings[index]
                return raw_value
            except (ValueError, IndexError):
                return raw_value
        elif cell_type == 'n':  # Number
            try:
                # Try int first, then float
                if '.' in raw_value or 'e' in raw_value.lower():
                    return float(raw_value)
                else:
                    return int(raw_value)
            except ValueError:
                return raw_value
        elif cell_type == 'b':  # Boolean
            return raw_value == '1'
        elif cell_type == 'str':  # Formula string
            return raw_value
        elif cell_type == 'inlineStr':  # Inline string
            return raw_value
        else:
            return raw_value