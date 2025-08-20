"""
XLSX Image Writer - Handles image embedding in Excel files
"""

import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple
import base64
from pathlib import Path

from ...drawing import Image, ImageFormat, ImageCollection, Anchor, AnchorType
from ...utils import tuple_to_coordinate


class ImageWriter:
    """Handles writing images to XLSX format following OOXML specifications."""
    
    def __init__(self):
        self.image_counter = 0
        self.relationship_counter = 0
        self.image_files: Dict[str, bytes] = {}  # path -> data
        self.image_relationships: List[Dict] = []  # relationship info
    
    def add_image(self, image: Image) -> str:
        """Add image and return its relationship ID."""
        self.image_counter += 1
        self.relationship_counter += 1
        
        # Generate paths
        rel_id = f"rId{self.relationship_counter}"
        image_filename = f"image{self.image_counter}.{image.format.value}"
        image_path = f"../media/{image_filename}"
        
        # Store image data - ensure it's bytes
        image_data = image.data
        if isinstance(image_data, str):
            # If data is base64 string, decode it
            try:
                image_data = base64.b64decode(image_data)
            except Exception:
                # If not base64, treat as raw string and encode
                image_data = image_data.encode('utf-8')
        self.image_files[f"xl/media/{image_filename}"] = image_data
        
        # Store relationship info
        self.image_relationships.append({
            'id': rel_id,
            'type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'target': image_path
        })
        
        return rel_id
    
    def create_drawing_xml(self, images: List[Image]) -> str:
        """Create drawing XML for worksheet images."""
        if not images:
            return ""
        
        # Create drawing XML structure - match OpenCells/Excel standard exactly
        drawing = ET.Element('xdr:wsDr')
        drawing.set('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
        drawing.set('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        # Note: r namespace declared locally in each blip element (OpenCells standard)
        
        for image in images:
            rel_id = self.add_image(image)
            anchor_elem = self._create_anchor_element(image, rel_id)
            drawing.append(anchor_elem)
        
        # Format XML with proper declaration and indentation
        self._indent_xml(drawing)
        xml_str = ET.tostring(drawing, encoding='unicode')
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n{xml_str}'
    
    def _create_anchor_element(self, image: Image, rel_id: str) -> ET.Element:
        """Create anchor element for image based on its positioning type."""
        anchor = image.anchor
        
        if anchor.type == AnchorType.ONE_CELL:
            return self._create_one_cell_anchor(image, rel_id)
        elif anchor.type == AnchorType.TWO_CELL:
            return self._create_two_cell_anchor(image, rel_id)
        elif anchor.type == AnchorType.ABSOLUTE:
            return self._create_absolute_anchor(image, rel_id)
        else:
            # Default to two cell anchor (Excel standard for images)
            return self._create_two_cell_anchor(image, rel_id)
    
    def _create_one_cell_anchor(self, image: Image, rel_id: str) -> ET.Element:
        """Create one-cell anchor element."""
        anchor_elem = ET.Element('xdr:oneCellAnchor')
        
        # From position
        from_elem = ET.SubElement(anchor_elem, 'xdr:from')
        from_row, from_col = image.anchor.from_position
        
        ET.SubElement(from_elem, 'xdr:col').text = str(from_col)
        ET.SubElement(from_elem, 'xdr:colOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[0] * 9525)  # Convert to EMU
        ET.SubElement(from_elem, 'xdr:row').text = str(from_row)
        ET.SubElement(from_elem, 'xdr:rowOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[1] * 9525)  # Convert to EMU
        
        # Extent (size) - EMU units (English Metric Units)
        ext_elem = ET.SubElement(anchor_elem, 'xdr:ext')
        width_emu = int((image.width or 100) * 9525)  # Convert pixels to EMU
        height_emu = int((image.height or 100) * 9525)  # Convert pixels to EMU
        ET.SubElement(ext_elem, 'xdr:cx').text = str(width_emu)
        ET.SubElement(ext_elem, 'xdr:cy').text = str(height_emu)
        
        # Picture
        pic_elem = self._create_picture_element(image, rel_id)
        anchor_elem.append(pic_elem)
        
        # Client data
        client_data = ET.SubElement(anchor_elem, 'xdr:clientData')
        
        return anchor_elem
    
    def _create_two_cell_anchor(self, image: Image, rel_id: str) -> ET.Element:
        """Create two-cell anchor element."""
        anchor_elem = ET.Element('xdr:twoCellAnchor')
        # Add editAs attribute like Excel standard
        anchor_elem.set('editAs', 'oneCell')
        
        # From position
        from_elem = ET.SubElement(anchor_elem, 'xdr:from')
        from_row, from_col = image.anchor.from_position
        
        ET.SubElement(from_elem, 'xdr:col').text = str(from_col)
        ET.SubElement(from_elem, 'xdr:colOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[0] * 9525)
        ET.SubElement(from_elem, 'xdr:row').text = str(from_row)
        ET.SubElement(from_elem, 'xdr:rowOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[1] * 9525)
        
        # To position
        to_elem = ET.SubElement(anchor_elem, 'xdr:to')
        if image.anchor.to_position:
            to_row, to_col = image.anchor.to_position
            to_col_off, to_row_off = image.anchor.to_offset
        else:
            # Calculate end position based on size
            to_row = from_row + 5  # Default span
            to_col = from_col + 3
            to_col_off, to_row_off = (0, 0)
        
        ET.SubElement(to_elem, 'xdr:col').text = str(to_col)
        ET.SubElement(to_elem, 'xdr:colOff').text = str(int(to_col_off * 9525))
        ET.SubElement(to_elem, 'xdr:row').text = str(to_row)
        ET.SubElement(to_elem, 'xdr:rowOff').text = str(int(to_row_off * 9525))
        
        # Picture
        pic_elem = self._create_picture_element(image, rel_id)
        anchor_elem.append(pic_elem)
        
        # Client data
        client_data = ET.SubElement(anchor_elem, 'xdr:clientData')
        
        return anchor_elem
    
    def _create_absolute_anchor(self, image: Image, rel_id: str) -> ET.Element:
        """Create absolute anchor element."""
        anchor_elem = ET.Element('xdr:absoluteAnchor')
        
        # Position
        pos_elem = ET.SubElement(anchor_elem, 'xdr:pos')
        abs_pos = getattr(image.anchor, 'absolute_position', (0, 0)) or (0, 0)
        ET.SubElement(pos_elem, 'xdr:x').text = str(int(abs_pos[0] * 9525))  # Convert to EMU
        ET.SubElement(pos_elem, 'xdr:y').text = str(int(abs_pos[1] * 9525))  # Convert to EMU
        
        # Extent (size)
        ext_elem = ET.SubElement(anchor_elem, 'xdr:ext')
        width_emu = int((image.width or 100) * 9525)
        height_emu = int((image.height or 100) * 9525)
        ET.SubElement(ext_elem, 'xdr:cx').text = str(width_emu)
        ET.SubElement(ext_elem, 'xdr:cy').text = str(height_emu)
        
        # Picture
        pic_elem = self._create_picture_element(image, rel_id)
        anchor_elem.append(pic_elem)
        
        # Client data
        client_data = ET.SubElement(anchor_elem, 'xdr:clientData')
        
        return anchor_elem
    
    def _create_picture_element(self, image: Image, rel_id: str) -> ET.Element:
        """Create picture element with all required sub-elements."""
        pic_elem = ET.Element('xdr:pic')
        
        # Non-visual picture properties
        nv_pic_pr = ET.SubElement(pic_elem, 'xdr:nvPicPr')
        
        # Non-visual drawing properties
        c_nv_pr = ET.SubElement(nv_pic_pr, 'xdr:cNvPr')
        c_nv_pr.set('id', str(self.image_counter))
        c_nv_pr.set('name', image.name or f'图片 {self.image_counter}')
        
        # CRITICAL: Add extension list with creation ID (required by Excel)
        ext_lst = ET.SubElement(c_nv_pr, 'a:extLst')
        ext = ET.SubElement(ext_lst, 'a:ext')
        ext.set('uri', '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}')
        
        # Add creation ID
        creation_id = ET.SubElement(ext, 'a16:creationId')
        creation_id.set('xmlns:a16', 'http://schemas.microsoft.com/office/drawing/2014/main')
        # Generate a unique GUID for this image
        import uuid
        creation_id.set('id', '{' + str(uuid.uuid4()).upper() + '}')
        
        # Non-visual picture drawing properties
        c_nv_pic_pr = ET.SubElement(nv_pic_pr, 'xdr:cNvPicPr')
        # Add picture locks for image protection - minimal attributes for Excel compatibility
        pic_locks = ET.SubElement(c_nv_pic_pr, 'a:picLocks')
        pic_locks.set('noChangeAspect', '1')  # Maintain aspect ratio
        
        # Blip fill
        blip_fill = ET.SubElement(pic_elem, 'xdr:blipFill')
        blip = ET.SubElement(blip_fill, 'a:blip')
        # CRITICAL: Use r:embed with local r namespace declaration like OpenCells
        blip.set('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        blip.set('r:embed', rel_id)  # Reference to image relationship
        blip.set('cstate', 'print')  # Image state for printing
        
        # CRITICAL: Add extension list with useLocalDpi (required by Excel)
        blip_ext_lst = ET.SubElement(blip, 'a:extLst')
        blip_ext = ET.SubElement(blip_ext_lst, 'a:ext')
        blip_ext.set('uri', '{28A0092B-C50C-407E-A947-70E740481C1C}')
        
        # Add useLocalDpi
        use_local_dpi = ET.SubElement(blip_ext, 'a14:useLocalDpi')
        use_local_dpi.set('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main')
        use_local_dpi.set('val', '0')
        
        # Stretch
        stretch = ET.SubElement(blip_fill, 'a:stretch')
        fill_rect = ET.SubElement(stretch, 'a:fillRect')
        
        # Shape properties
        sp_pr = ET.SubElement(pic_elem, 'xdr:spPr')
        
        # Transform
        xfrm = ET.SubElement(sp_pr, 'a:xfrm')
        off = ET.SubElement(xfrm, 'a:off')
        off.set('x', '0')
        off.set('y', '0')
        ext = ET.SubElement(xfrm, 'a:ext')
        width_emu = int((image.width or 100) * 9525)
        height_emu = int((image.height or 100) * 9525)
        ext.set('cx', str(width_emu))
        ext.set('cy', str(height_emu))
        
        # Preset geometry - rectangle shape for image
        prst_geom = ET.SubElement(sp_pr, 'a:prstGeom')
        prst_geom.set('prst', 'rect')
        av_lst = ET.SubElement(prst_geom, 'a:avLst')
        
        # Remove line element for images - can cause Excel validation issues
        
        return pic_elem
    
    def create_drawing_rels_xml(self) -> str:
        """Create drawing relationships XML."""
        if not self.image_relationships:
            return ""
        
        relationships = ET.Element('Relationships')
        relationships.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
        
        for rel in self.image_relationships:
            relationship = ET.SubElement(relationships, 'Relationship')
            relationship.set('Id', rel['id'])
            relationship.set('Type', rel['type'])
            relationship.set('Target', rel['target'])
        
        # Format XML with proper declaration and indentation
        self._indent_xml(relationships)
        xml_str = ET.tostring(relationships, encoding='unicode')
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n{xml_str}'
    
    def get_image_files(self) -> Dict[str, bytes]:
        """Get all image files that need to be written to the archive."""
        return self.image_files.copy()
    
    def get_content_types_entries(self) -> List[Tuple[str, str]]:
        """Get content types entries for images."""
        entries = []
        for path in self.image_files.keys():
            filename = Path(path).name
            if filename.endswith('.png'):
                entries.append(('png', 'image/png'))
            elif filename.endswith(('.jpg', '.jpeg')):
                entries.append(('jpeg', 'image/jpeg'))
            elif filename.endswith('.gif'):
                entries.append(('gif', 'image/gif'))
        
        # Remove duplicates
        return list(set(entries))
    
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