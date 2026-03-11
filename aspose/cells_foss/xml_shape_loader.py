"""
Aspose.Cells for Python - Shape XML Loader Module

Parses xdr:sp (shape) elements from SpreadsheetML drawing XML parts and
populates a worksheet's ShapeCollection.
"""

import xml.etree.ElementTree as ET

from .shape import (
    Shape, MsoDrawingType, FillType, MsoLineDashStyle,
    TextAlignmentType, TextAnchorType,
    _PRST_TO_DRAWING_TYPE, _DASH_VAL_TO_STYLE,
    _TEXT_ALIGN_VAL_TO_TYPE, _ANCHOR_VAL_TO_TYPE,
)

_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_XDR_NS = {"xdr": _XDR, "a": _A, "r": _R}

_HLINKCLICK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def _tag(ns, local):
    return f"{{{ns}}}{local}"


class ShapeXmlLoader:
    """Loads xdr:sp shape elements from a drawing XML part."""

    @staticmethod
    def _get_child_int(parent, local_name, default=0):
        """Returns int value of an xdr child element, or default when missing/invalid."""
        if parent is None:
            return default
        child = parent.find(_tag(_XDR, local_name))
        if child is None or child.text is None:
            return default
        try:
            return int(float(child.text))
        except ValueError:
            return default

    def load_shapes(self, worksheet, drawing_root, drawing_path,
                    drawing_rels, zipf, get_anchor_int_fn,
                    drawing_rel_types=None):
        """
        Parses all xdr:sp elements in drawing_root and appends Shape objects
        to worksheet.shapes.

        Args:
            worksheet: The Worksheet being loaded into.
            drawing_root: xml.etree.ElementTree.Element — root of the drawing XML.
            drawing_path (str): ZIP path of the drawing file (unused, kept for
                signature parity with PictureXmlLoader).
            drawing_rels (dict): Mapping of rel_id → target path.
            zipf: zipfile.ZipFile object (unused for shapes without embeds).
            get_anchor_int_fn: Callable that converts an xdr:col / xdr:row
                element to an int (handles missing elements gracefully).
            drawing_rel_types (dict|None): Mapping of rel_id → relationship type
                string, used to detect hyperlinks.
        """
        if drawing_rel_types is None:
            drawing_rel_types = {}

        anchor_tags = (
            _tag(_XDR, "twoCellAnchor"),
            _tag(_XDR, "oneCellAnchor"),
        )

        for anchor in drawing_root:
            if anchor.tag not in anchor_tags:
                continue

            sp_elem = anchor.find(_tag(_XDR, "sp"))
            if sp_elem is None:
                # This anchor holds a chart (graphicFrame) or picture (pic) — skip.
                continue

            shape = self._parse_sp(sp_elem, anchor, get_anchor_int_fn,
                                   drawing_rels, drawing_rel_types)
            if shape is not None:
                worksheet.shapes._add_loaded(shape)

    # -----------------------------------------------------------------------
    # Private helpers
    # -----------------------------------------------------------------------

    def _parse_sp(self, sp_elem, anchor, get_anchor_int_fn,
                  drawing_rels, drawing_rel_types):
        """Converts a single xdr:sp element + its anchor into a Shape."""
        try:
            # ---- anchor coords ----
            from_elem = anchor.find(_tag(_XDR, "from"))
            to_elem   = anchor.find(_tag(_XDR, "to"))

            ul_col = self._get_child_int(from_elem, "col", default=0)
            ul_row = self._get_child_int(from_elem, "row", default=0)
            lr_col = self._get_child_int(to_elem, "col", default=0)
            lr_row = self._get_child_int(to_elem, "row", default=0)

            ul_col_off = self._get_child_int(from_elem, "colOff", default=0)
            ul_row_off = self._get_child_int(from_elem, "rowOff", default=0)
            lr_col_off = self._get_child_int(to_elem, "colOff", default=0)
            lr_row_off = self._get_child_int(to_elem, "rowOff", default=0)

            # ---- nvSpPr ----
            nvsppr = sp_elem.find(_tag(_XDR, "nvSpPr"))
            name = ""
            is_text_box = False
            hyperlink_url = None

            if nvsppr is not None:
                cnvpr = nvsppr.find(_tag(_XDR, "cNvPr"))
                if cnvpr is not None:
                    name = cnvpr.get("name", "")
                    # Hyperlink
                    hlinkclick = cnvpr.find(_tag(_A, "hlinkClick"))
                    if hlinkclick is not None:
                        rel_id = hlinkclick.get(_tag(_R, "id"))
                        if rel_id and drawing_rel_types.get(rel_id) == _HLINKCLICK_REL_TYPE:
                            target = drawing_rels.get(rel_id, "")
                            if target:
                                hyperlink_url = target

                cnvsppr = nvsppr.find(_tag(_XDR, "cNvSpPr"))
                if cnvsppr is not None and cnvsppr.get("txBox") == "1":
                    is_text_box = True

            # ---- shape type from prstGeom ----
            sppr = sp_elem.find(_tag(_XDR, "spPr"))
            drawing_type = MsoDrawingType.UNKNOWN

            if sppr is not None:
                prstgeom = sppr.find(_tag(_A, "prstGeom"))
                if prstgeom is not None:
                    prst = prstgeom.get("prst", "")
                    if is_text_box:
                        drawing_type = MsoDrawingType.TEXT_BOX
                    else:
                        drawing_type = _PRST_TO_DRAWING_TYPE.get(prst, MsoDrawingType.UNKNOWN)
                elif is_text_box:
                    drawing_type = MsoDrawingType.TEXT_BOX
            elif is_text_box:
                drawing_type = MsoDrawingType.TEXT_BOX

            shape = Shape(drawing_type, ul_row, ul_col, lr_row, lr_col)
            shape.name = name
            shape._upper_left_column_offset = ul_col_off
            shape._upper_left_row_offset    = ul_row_off
            shape._lower_right_column_offset = lr_col_off
            shape._lower_right_row_offset    = lr_row_off
            shape.hyperlink = hyperlink_url

            # editAs
            edit_as = anchor.get("editAs", "twoCell")
            shape.placement = edit_as

            # ---- fill ----
            if sppr is not None:
                self._parse_fill(sppr, shape)
                self._parse_line(sppr, shape)

            # ---- txBody ----
            txbody = sp_elem.find(_tag(_XDR, "txBody"))
            if txbody is not None:
                self._parse_txbody(txbody, shape)

            # ---- preserve raw xdr:sp XML for round-trip ----
            shape._source_xml = ET.tostring(sp_elem, encoding="unicode")

            return shape

        except Exception:
            return None

    def _parse_fill(self, sppr, shape):
        """Reads fill properties from xdr:spPr into shape.fill."""
        # <a:solidFill>
        solid = sppr.find(_tag(_A, "solidFill"))
        if solid is not None:
            shape.fill.fill_type = FillType.SOLID
            srgb = solid.find(_tag(_A, "srgbClr"))
            if srgb is not None:
                shape.fill.fore_color = srgb.get("val", "FFFFFF").upper()
                # lumMod / alpha not extracted for simplicity
            return

        # <a:noFill>
        nofill = sppr.find(_tag(_A, "noFill"))
        if nofill is not None:
            shape.fill.fill_type = FillType.NONE
            return

        # <a:gradFill> / <a:pattFill> — mark type but don't parse details
        if sppr.find(_tag(_A, "gradFill")) is not None:
            shape.fill.fill_type = FillType.GRADIENT
        elif sppr.find(_tag(_A, "pattFill")) is not None:
            shape.fill.fill_type = FillType.PATTERN

    def _parse_line(self, sppr, shape):
        """Reads border/outline properties from xdr:spPr into shape.line."""
        ln = sppr.find(_tag(_A, "ln"))
        if ln is None:
            return

        w = ln.get("w")
        if w is not None:
            try:
                shape.line.weight = int(w)
            except ValueError:
                pass

        # <a:noFill> child → invisible border
        if ln.find(_tag(_A, "noFill")) is not None:
            shape.line.is_visible = False
            return

        shape.line.is_visible = True

        solid = ln.find(_tag(_A, "solidFill"))
        if solid is not None:
            srgb = solid.find(_tag(_A, "srgbClr"))
            if srgb is not None:
                shape.line.color = srgb.get("val", "000000").upper()

        prstdash = ln.find(_tag(_A, "prstDash"))
        if prstdash is not None:
            val = prstdash.get("val", "solid")
            shape.line.dash_style = _DASH_VAL_TO_STYLE.get(val, MsoLineDashStyle.SOLID)

    def _parse_txbody(self, txbody, shape):
        """Reads text body properties from xdr:txBody into shape."""
        bodypr = txbody.find(_tag(_A, "bodyPr"))
        if bodypr is not None:
            anchor_val = bodypr.get("anchor", "t")
            shape.text_vertical_alignment = _ANCHOR_VAL_TO_TYPE.get(
                anchor_val, TextAnchorType.TOP)
            wrap_val = bodypr.get("wrap", "square")
            shape.is_text_wrapped = (wrap_val != "none")

        # Collect text from all paragraphs / runs
        text_parts = []
        para_aligns = []

        for para in txbody.findall(_tag(_A, "p")):
            ppr = para.find(_tag(_A, "pPr"))
            if ppr is not None:
                algn = ppr.get("algn", "l")
                para_aligns.append(_TEXT_ALIGN_VAL_TO_TYPE.get(
                    algn, TextAlignmentType.LEFT))

            para_texts = []
            for run in para.findall(_tag(_A, "r")):
                t_elem = run.find(_tag(_A, "t"))
                if t_elem is not None and t_elem.text:
                    para_texts.append(t_elem.text)

                # Font from the first run's rPr
                if not shape.font.name or shape.font.name == "Calibri":
                    rpr = run.find(_tag(_A, "rPr"))
                    if rpr is not None:
                        sz = rpr.get("sz")
                        if sz is not None:
                            try:
                                shape.font.size = int(sz) / 100.0
                            except ValueError:
                                pass
                        shape.font.bold   = rpr.get("b", "0") == "1"
                        shape.font.italic = rpr.get("i", "0") == "1"
                        shape.font.underline = rpr.get("u") not in (None, "none")
                        solid = rpr.find(_tag(_A, "solidFill"))
                        if solid is not None:
                            srgb = solid.find(_tag(_A, "srgbClr"))
                            if srgb is not None:
                                shape.font.color = srgb.get("val", "000000").upper()
                        latin = rpr.find(_tag(_A, "latin"))
                        if latin is not None:
                            shape.font.name = latin.get("typeface", "Calibri")

            text_parts.append("".join(para_texts))

        shape.text = "\n".join(text_parts) if text_parts else ""

        # Use the alignment from the first paragraph that has one
        for a in para_aligns:
            shape.text_horizontal_alignment = a
            break
