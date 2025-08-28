"""
XLSX Image Writer - Handles image embedding in Excel files
"""

import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple
import base64
from pathlib import Path

from ...drawing import Image, ImageFormat, ImageCollection, Anchor, AnchorType
from ...utils import tuple_to_coordinate
from .constants import XlsxConstants


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
            'type': f"{XlsxConstants.REL_TYPES['worksheet'].rsplit('/', 1)[0]}/image",
            'target': image_path
        })
        
        return rel_id
    
    def create_drawing_xml(self, images: List[Image]) -> str:
        """Create drawing XML for worksheet images."""
        if not images:
            return ""
        
        # Create drawing XML structure - match OpenCells/Excel standard exactly
        drawing = ET.Element('xdr:wsDr')
        drawing.set('xmlns:xdr', XlsxConstants.DRAWING_NAMESPACES['xdr'])
        drawing.set('xmlns:a', XlsxConstants.DRAWING_NAMESPACES['a'])
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
        ET.SubElement(from_elem, 'xdr:colOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[0] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        ET.SubElement(from_elem, 'xdr:row').text = str(from_row)
        ET.SubElement(from_elem, 'xdr:rowOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[1] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        
        # Extent (size) - EMU units (English Metric Units)
        ext_elem = ET.SubElement(anchor_elem, 'xdr:ext')
        width_emu = int((image.width or XlsxConstants.IMAGE_DEFAULTS['width']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        height_emu = int((image.height or XlsxConstants.IMAGE_DEFAULTS['height']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
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
        anchor_elem.set('editAs', XlsxConstants.XML_ATTRIBUTES['edit_as'])
        
        # From position
        from_elem = ET.SubElement(anchor_elem, 'xdr:from')
        from_row, from_col = image.anchor.from_position
        
        ET.SubElement(from_elem, 'xdr:col').text = str(from_col)
        ET.SubElement(from_elem, 'xdr:colOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[0] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        ET.SubElement(from_elem, 'xdr:row').text = str(from_row)
        ET.SubElement(from_elem, 'xdr:rowOff').text = str(getattr(image.anchor, 'from_offset', (0, 0))[1] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        
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
        ET.SubElement(to_elem, 'xdr:colOff').text = str(int(to_col_off * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel']))
        ET.SubElement(to_elem, 'xdr:row').text = str(to_row)
        ET.SubElement(to_elem, 'xdr:rowOff').text = str(int(to_row_off * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel']))
        
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
        ET.SubElement(pos_elem, 'xdr:x').text = str(int(abs_pos[0] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel']))
        ET.SubElement(pos_elem, 'xdr:y').text = str(int(abs_pos[1] * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel']))
        
        # Extent (size)
        ext_elem = ET.SubElement(anchor_elem, 'xdr:ext')
        width_emu = int((image.width or XlsxConstants.IMAGE_DEFAULTS['width']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        height_emu = int((image.height or XlsxConstants.IMAGE_DEFAULTS['height']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
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
        c_nv_pr.set('name', image.name or f'{XlsxConstants.IMAGE_DEFAULTS["name_prefix"]} {self.image_counter}')
        
        # CRITICAL: Add extension list with creation ID (required by Excel)
        ext_lst = ET.SubElement(c_nv_pr, 'a:extLst')
        ext = ET.SubElement(ext_lst, 'a:ext')
        ext.set('uri', XlsxConstants.IMAGE_EXTENSIONS['creation_id'])
        
        # Add creation ID
        creation_id = ET.SubElement(ext, 'a16:creationId')
        creation_id.set('xmlns:a16', XlsxConstants.DRAWING_NAMESPACES['a16'])
        # Generate a unique GUID for this image
        import uuid
        creation_id.set('id', '{' + str(uuid.uuid4()).upper() + '}')
        
        # Non-visual picture drawing properties
        c_nv_pic_pr = ET.SubElement(nv_pic_pr, 'xdr:cNvPicPr')
        # Add picture locks for image protection - minimal attributes for Excel compatibility
        pic_locks = ET.SubElement(c_nv_pic_pr, 'a:picLocks')
        pic_locks.set('noChangeAspect', XlsxConstants.XML_ATTRIBUTES['no_change_aspect'])
        
        # Blip fill
        blip_fill = ET.SubElement(pic_elem, 'xdr:blipFill')
        blip = ET.SubElement(blip_fill, 'a:blip')
        # CRITICAL: Use r:embed with local r namespace declaration like OpenCells
        blip.set('xmlns:r', XlsxConstants.NAMESPACES['r'])
        blip.set('r:embed', rel_id)  # Reference to image relationship
        blip.set('cstate', XlsxConstants.XML_ATTRIBUTES['cstate'])
        
        # CRITICAL: Add extension list with useLocalDpi (required by Excel)
        blip_ext_lst = ET.SubElement(blip, 'a:extLst')
        blip_ext = ET.SubElement(blip_ext_lst, 'a:ext')
        blip_ext.set('uri', XlsxConstants.IMAGE_EXTENSIONS['local_dpi'])
        
        # Add useLocalDpi
        use_local_dpi = ET.SubElement(blip_ext, 'a14:useLocalDpi')
        use_local_dpi.set('xmlns:a14', XlsxConstants.DRAWING_NAMESPACES['a14'])
        use_local_dpi.set('val', XlsxConstants.XML_ATTRIBUTES['local_dpi_val'])
        
        # Stretch
        stretch = ET.SubElement(blip_fill, 'a:stretch')
        fill_rect = ET.SubElement(stretch, 'a:fillRect')
        
        # Shape properties
        sp_pr = ET.SubElement(pic_elem, 'xdr:spPr')
        
        # Transform
        xfrm = ET.SubElement(sp_pr, 'a:xfrm')
        off = ET.SubElement(xfrm, 'a:off')
        off.set('x', XlsxConstants.XML_ATTRIBUTES['transform_xy'])
        off.set('y', XlsxConstants.XML_ATTRIBUTES['transform_xy'])
        ext = ET.SubElement(xfrm, 'a:ext')
        width_emu = int((image.width or XlsxConstants.IMAGE_DEFAULTS['width']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        height_emu = int((image.height or XlsxConstants.IMAGE_DEFAULTS['height']) * XlsxConstants.IMAGE_DEFAULTS['emu_per_pixel'])
        ext.set('cx', str(width_emu))
        ext.set('cy', str(height_emu))
        
        # Preset geometry - rectangle shape for image
        prst_geom = ET.SubElement(sp_pr, 'a:prstGeom')
        prst_geom.set('prst', XlsxConstants.XML_ATTRIBUTES['preset_geom'])
        av_lst = ET.SubElement(prst_geom, 'a:avLst')
        
        # Remove line element for images - can cause Excel validation issues
        
        return pic_elem
    
    def create_drawing_rels_xml(self) -> str:
        """Create drawing relationships XML."""
        if not self.image_relationships:
            return ""
        
        relationships = ET.Element('Relationships')
        relationships.set('xmlns', XlsxConstants.NAMESPACES['pkg'])
        
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
                entries.append(('png', XlsxConstants.IMAGE_CONTENT_TYPES['png']))
            elif filename.endswith(('.jpg', '.jpeg')):
                entries.append(('jpeg', XlsxConstants.IMAGE_CONTENT_TYPES['jpeg']))
            elif filename.endswith('.gif'):
                entries.append(('gif', XlsxConstants.IMAGE_CONTENT_TYPES['gif']))
        
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