"""
Aspose.Cells for Python - Picture XML Loader Module

Loads worksheet pictures from drawing parts and resolves picture content types.
"""

import posixpath
import xml.etree.ElementTree as ET


# Namespace prefixes used when serializing stored extLst XML back to string.
_SERIALIZE_NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
}


def _register_serialize_namespaces():
    for prefix, uri in _SERIALIZE_NS.items():
        ET.register_namespace(prefix, uri)


def _elem_to_xml_str(elem):
    """Serialize an ET element to a string, using known namespace prefixes."""
    _register_serialize_namespaces()
    return ET.tostring(elem, encoding='unicode')


class PictureXmlLoader:
    """Loads pictures from worksheet drawing parts."""

    def __init__(self):
        self._xdr_ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        self._r_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        self._hyperlink_rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

    def collect_drawing_image_parts(
        self,
        worksheet,
        drawing_path,
        drawing_rels,
        drawing_rel_types,
        zipf,
        content_type_overrides=None,
        content_type_defaults=None,
    ):
        worksheet._source_drawing_extra_parts = []
        for rel_id, target in drawing_rels.items():
            rel_type = drawing_rel_types.get(rel_id, "")
            if not rel_type.endswith('/image'):
                continue
            part_path = self._resolve_target(drawing_path, target)
            part_bytes = self._try_read_bytes(zipf, part_path)
            if part_bytes is None:
                continue
            content_type = self.resolve_content_type(part_path, content_type_overrides, content_type_defaults)
            worksheet._source_drawing_extra_parts.append((part_path, part_bytes, content_type))

    def load_pictures(
        self,
        worksheet,
        drawing_root,
        drawing_path,
        drawing_rels,
        zipf,
        get_anchor_int,
        content_type_overrides=None,
        content_type_defaults=None,
        drawing_rel_types=None,
    ):
        drawing_rel_types = drawing_rel_types or {}
        anchors = []
        anchors.extend(drawing_root.findall('xdr:twoCellAnchor', namespaces=self._xdr_ns))
        anchors.extend(drawing_root.findall('xdr:oneCellAnchor', namespaces=self._xdr_ns))
        for anchor in anchors:
            pics = anchor.findall('.//xdr:pic', namespaces=self._xdr_ns)
            if not pics:
                continue
            from_col = get_anchor_int(anchor, 'xdr:from/xdr:col', default=0)
            from_row = get_anchor_int(anchor, 'xdr:from/xdr:row', default=0)
            to_col = get_anchor_int(anchor, 'xdr:to/xdr:col', default=from_col + 1)
            to_row = get_anchor_int(anchor, 'xdr:to/xdr:row', default=from_row + 1)
            from_col_off = get_anchor_int(anchor, 'xdr:from/xdr:colOff', default=0)
            from_row_off = get_anchor_int(anchor, 'xdr:from/xdr:rowOff', default=0)
            to_col_off = get_anchor_int(anchor, 'xdr:to/xdr:colOff', default=0)
            to_row_off = get_anchor_int(anchor, 'xdr:to/xdr:rowOff', default=0)

            # Read editAs attribute from anchor element (e.g. "oneCell", "absolute")
            edit_as = anchor.get('editAs')

            for pic_elem in pics:
                blip = pic_elem.find('xdr:blipFill/a:blip', namespaces=self._xdr_ns)
                if blip is None:
                    continue
                embed_rel_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if not embed_rel_id:
                    continue
                target = drawing_rels.get(embed_rel_id)
                if not target:
                    continue
                part_path = self._resolve_target(drawing_path, target)
                part_bytes = self._try_read_bytes(zipf, part_path)
                if part_bytes is None:
                    continue
                extension = posixpath.splitext(part_path)[1].lstrip('.').lower() or 'png'
                c_nv_pr = pic_elem.find('xdr:nvPicPr/xdr:cNvPr', namespaces=self._xdr_ns)
                name = c_nv_pr.get('name') if c_nv_pr is not None else None
                pic = worksheet.pictures._add_loaded(
                    part_bytes,
                    extension,
                    from_row,
                    from_col,
                    to_row,
                    to_col,
                    name=name,
                )
                pic._upper_left_column_offset = from_col_off
                pic._upper_left_row_offset = from_row_off
                pic._lower_right_column_offset = to_col_off
                pic._lower_right_row_offset = to_row_off
                pic._source_part_path = part_path
                pic._source_content_type = self.resolve_content_type(part_path, content_type_overrides, content_type_defaults)

                # Load extended picture attributes
                pic._edit_as = edit_as  # e.g. "oneCell", "absolute", or None (= "twoCell")

                # Load hyperlink URL from <a:hlinkClick r:id="..."/> inside <xdr:cNvPr>
                if c_nv_pr is not None:
                    hlink = c_nv_pr.find('a:hlinkClick', namespaces=self._xdr_ns)
                    if hlink is not None:
                        hlink_rel_id = hlink.get(self._r_attr)
                        if hlink_rel_id:
                            rel_type = drawing_rel_types.get(hlink_rel_id, '')
                            if rel_type.endswith('/hyperlink'):
                                pic._hyperlink_url = drawing_rels.get(hlink_rel_id)

                    # Load extLst from cNvPr for round-trip preservation
                    cNvPr_extLst = c_nv_pr.find('a:extLst', namespaces=self._xdr_ns)
                    if cNvPr_extLst is not None:
                        pic._source_cNvPr_extLst_xml = _elem_to_xml_str(cNvPr_extLst)

                # Load noChangeAspect from <xdr:cNvPicPr>/<a:picLocks>
                c_nv_pic_pr = pic_elem.find('xdr:nvPicPr/xdr:cNvPicPr', namespaces=self._xdr_ns)
                if c_nv_pic_pr is not None:
                    pic_locks = c_nv_pic_pr.find('a:picLocks', namespaces=self._xdr_ns)
                    if pic_locks is not None:
                        no_change = pic_locks.get('noChangeAspect', '0')
                        pic._no_change_aspect = no_change == '1'
                    else:
                        pic._no_change_aspect = False

                # Load extLst from <a:blip> for round-trip preservation
                blip_extLst = blip.find('a:extLst', namespaces=self._xdr_ns)
                if blip_extLst is not None:
                    pic._source_blip_extLst_xml = _elem_to_xml_str(blip_extLst)

                # Load spPr inner content for round-trip preservation
                sp_pr = pic_elem.find('xdr:spPr', namespaces=self._xdr_ns)
                if sp_pr is not None and len(sp_pr) > 0:
                    inner_parts = [_elem_to_xml_str(child) for child in sp_pr]
                    pic._source_spPr_xml = ''.join(inner_parts)

    def resolve_content_type(self, part_path, content_type_overrides, content_type_defaults):
        normalized = f'/{part_path.lstrip("/")}'
        if content_type_overrides:
            ctype = content_type_overrides.get(normalized)
            if ctype:
                return ctype
        ext = posixpath.splitext(part_path)[1].lstrip('.').lower()
        if ext and content_type_defaults:
            return content_type_defaults.get(ext)
        return None

    def _try_read_bytes(self, zipf, zip_path):
        try:
            return zipf.read(zip_path)
        except KeyError:
            return None

    def _resolve_target(self, base_part_path, target):
        base_dir = posixpath.dirname(base_part_path)
        combined = posixpath.normpath(posixpath.join(base_dir, target))
        return combined.lstrip('/')
