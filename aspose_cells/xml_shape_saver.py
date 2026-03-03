"""
Aspose.Cells for Python - Shape XML Saver Module

Generates xdr:sp (shape) XML elements for SpreadsheetML drawing parts
and corresponding relationship entries.
"""

from .shape import (
    MsoDrawingType, FillType, MsoLineDashStyle,
    TextAlignmentType, TextAnchorType,
    _DRAWING_TYPE_TO_PRST, _DASH_STYLE_TO_VAL,
    _TEXT_ALIGN_TO_VAL, _ANCHOR_TO_VAL,
)

_HYPERLINK_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class ShapeXmlSaver:
    """Generates drawing XML and relationship XML for worksheet shapes."""

    def __init__(self, escape_xml_fn):
        """
        Args:
            escape_xml_fn: Callable that XML-escapes a string (provided by
                the parent saver so this module doesn't duplicate it).
        """
        self._esc = escape_xml_fn

    # -----------------------------------------------------------------------
    # Public API (mirrors PictureXmlSaver interface)
    # -----------------------------------------------------------------------

    def collect_shape_refs(self, worksheet):
        """
        Builds the list of shape reference dicts needed for drawing XML
        generation.

        Returns:
            list[dict]: Each dict has keys:
                - 'shape': Shape object
                - 'hyperlink_rel_id': str rel ID if shape has a hyperlink, else None
        """
        shapes = worksheet.shapes
        return [{'shape': s, 'hyperlink_rel_id': None} for s in shapes]

    def assign_hyperlink_rel_ids(self, shape_refs, start_rel_id):
        """
        Assigns relationship IDs to shapes that have hyperlink URLs.

        Args:
            shape_refs (list[dict]): From collect_shape_refs().
            start_rel_id (int): First available rel ID integer.

        Returns:
            int: Next available rel ID after shape assignments.
        """
        next_id = start_rel_id
        for ref in shape_refs:
            shape = ref['shape']
            if shape.hyperlink:
                ref['hyperlink_rel_id'] = f"rId{next_id}"
                next_id += 1
        return next_id

    def format_shape_anchors_xml(self, shape_refs, chart_count, picture_count):
        """
        Returns concatenated xdr:twoCellAnchor XML strings for all shapes.

        Args:
            shape_refs (list[dict]): From collect_shape_refs() (after
                assign_hyperlink_rel_ids() has been called).
            chart_count (int): Number of chart anchors already written
                (used for object ID uniqueness only).
            picture_count (int): Number of picture anchors already written
                (used for object ID uniqueness only).

        Returns:
            str: XML fragments ready for insertion into the drawing XML body.
        """
        parts = []
        for i, ref in enumerate(shape_refs):
            object_id = 3000 + i
            parts.append(self._format_shape_anchor_xml(
                ref['shape'], object_id, ref['hyperlink_rel_id']))
        return "".join(parts)

    def format_shape_relationships_xml(self, shape_refs, chart_count, picture_count):
        """
        Returns Relationship XML entries for shapes that have hyperlinks.

        Returns:
            str: Zero or more <Relationship .../> elements.
        """
        parts = []
        for ref in shape_refs:
            rel_id = ref['hyperlink_rel_id']
            if rel_id and ref['shape'].hyperlink:
                escaped_url = self._esc(ref['shape'].hyperlink)
                parts.append(
                    f'    <Relationship Id="{rel_id}"'
                    f' Type="{_HYPERLINK_REL_TYPE}"'
                    f' Target="{escaped_url}" TargetMode="External"/>\n'
                )
        return "".join(parts)

    # -----------------------------------------------------------------------
    # Private helpers
    # -----------------------------------------------------------------------

    def _format_shape_anchor_xml(self, shape, object_id, hyperlink_rel_id):
        """Generates the full xdr:twoCellAnchor XML for one shape."""
        edit_as_attr = ""
        if shape.placement and shape.placement != "twoCell":
            edit_as_attr = f' editAs="{self._esc(shape.placement)}"'

        c = f'    <xdr:twoCellAnchor{edit_as_attr}>\n'

        # from/to anchors
        c += self._format_anchor_coords(
            "from",
            shape._upper_left_column, shape._upper_left_column_offset,
            shape._upper_left_row,    shape._upper_left_row_offset,
        )
        c += self._format_anchor_coords(
            "to",
            shape._lower_right_column, shape._lower_right_column_offset,
            shape._lower_right_row,    shape._lower_right_row_offset,
        )

        # xdr:sp — use round-trip source if available
        if shape._source_xml:
            c += f'      {shape._source_xml}\n'
        else:
            c += self._format_sp_xml(shape, object_id, hyperlink_rel_id)

        c += '      <xdr:clientData/>\n'
        c += '    </xdr:twoCellAnchor>\n'
        return c

    def _format_anchor_coords(self, tag, col, col_off, row, row_off):
        return (
            f'      <xdr:{tag}>'
            f'<xdr:col>{col}</xdr:col>'
            f'<xdr:colOff>{col_off}</xdr:colOff>'
            f'<xdr:row>{row}</xdr:row>'
            f'<xdr:rowOff>{row_off}</xdr:rowOff>'
            f'</xdr:{tag}>\n'
        )

    def _format_sp_xml(self, shape, object_id, hyperlink_rel_id):
        """Generates the full <xdr:sp> element XML from shape properties."""
        is_text_box = (shape.drawing_type == MsoDrawingType.TEXT_BOX)
        txbox_attr = ' txBox="1"' if is_text_box else ''
        prst = _DRAWING_TYPE_TO_PRST.get(shape.drawing_type, "rect")
        safe_name = self._esc(shape.name or "Shape")

        c = '      <xdr:sp macro="" textlink="">\n'

        # nvSpPr
        c += '        <xdr:nvSpPr>\n'
        c += f'          <xdr:cNvPr id="{object_id}" name="{safe_name}"'
        if hyperlink_rel_id:
            c += '>\n'
            c += (
                f'            <a:hlinkClick xmlns:r="{_REL_NS}" '
                f'r:id="{hyperlink_rel_id}"/>\n'
            )
            c += '          </xdr:cNvPr>\n'
        else:
            c += '/>\n'
        c += f'          <xdr:cNvSpPr{txbox_attr}/>\n'
        c += '        </xdr:nvSpPr>\n'

        # spPr
        c += '        <xdr:spPr>\n'
        c += '          <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>\n'
        c += f'          <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>\n'
        c += self._format_fill_xml(shape.fill)
        c += self._format_line_xml(shape.line)
        c += '        </xdr:spPr>\n'

        # txBody — always write for TEXT_BOX; write for other shapes with text
        if shape.text or is_text_box:
            c += self._format_txbody_xml(shape)

        c += '      </xdr:sp>\n'
        return c

    def _format_fill_xml(self, fill):
        """Returns fill XML fragment (indented for spPr context)."""
        ft = fill.fill_type
        if ft == FillType.NONE:
            return '          <a:noFill/>\n'
        if ft == FillType.SOLID:
            color = (fill.fore_color or "FFFFFF").upper()
            return (
                '          <a:solidFill>'
                f'<a:srgbClr val="{color}"/>'
                '</a:solidFill>\n'
            )
        # GRADIENT / PATTERN / AUTOMATIC — omit explicit fill (inherit from theme)
        return ''

    def _format_line_xml(self, line):
        """Returns line/border XML fragment."""
        if not line.is_visible:
            return '          <a:ln><a:noFill/></a:ln>\n'

        w = line.weight
        color = (line.color or "000000").upper()

        c = f'          <a:ln w="{w}">\n'
        c += f'            <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>\n'

        # Dash style
        if line.dash_style != MsoLineDashStyle.SOLID:
            val = _DASH_STYLE_TO_VAL.get(line.dash_style, "dash")
            c += f'            <a:prstDash val="{val}"/>\n'

        c += '          </a:ln>\n'
        return c

    def _format_txbody_xml(self, shape):
        """Returns <xdr:txBody> XML for a shape."""
        anchor_val = _ANCHOR_TO_VAL.get(shape.text_vertical_alignment, "t")
        wrap_val   = "square" if shape.is_text_wrapped else "none"
        algn_val   = _TEXT_ALIGN_TO_VAL.get(shape.text_horizontal_alignment, "l")

        c = '        <xdr:txBody>\n'
        c += (
            f'          <a:bodyPr vertOverflow="clip" wrap="{wrap_val}"'
            f' lIns="91440" tIns="45720" rIns="91440" bIns="45720"'
            f' anchor="{anchor_val}" anchorCtr="0"/>\n'
        )
        c += '          <a:lstStyle/>\n'

        font = shape.font
        sz   = int(font.size * 100)
        b    = "1" if font.bold    else "0"
        i    = "1" if font.italic  else "0"
        u    = "sng" if font.underline else "none"
        color = (font.color or "000000").upper()
        fname = self._esc(font.name or "Calibri")

        # Split text into paragraphs; empty text → one empty paragraph
        paragraphs = shape.text.split("\n") if shape.text else [""]

        for para_text in paragraphs:
            c += '          <a:p>\n'
            if algn_val != "l":
                c += f'            <a:pPr algn="{algn_val}"/>\n'
            if para_text:
                c += '            <a:r>\n'
                c += (
                    f'              <a:rPr lang="en-US" sz="{sz}"'
                    f' b="{b}" i="{i}" u="{u}" dirty="0">\n'
                    f'                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>\n'
                    f'                <a:latin typeface="{fname}"/>\n'
                    f'              </a:rPr>\n'
                )
                c += f'              <a:t>{self._esc(para_text)}</a:t>\n'
                c += '            </a:r>\n'
            c += '          </a:p>\n'

        c += '        </xdr:txBody>\n'
        return c
