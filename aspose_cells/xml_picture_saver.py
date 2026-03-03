"""
Aspose.Cells for Python - Picture XML Saver Module

Writes worksheet picture parts and drawing picture anchors/relationships.
"""

import posixpath

_HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
_IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'


class PictureXmlSaver:
    """Handles writing picture-related drawing/media XML payloads."""

    def __init__(self, escape_xml):
        self._escape_xml = escape_xml

    def collect_picture_refs(self, zipf, worksheet, sheet_num, write_part_once):
        refs = []
        if getattr(worksheet, "pictures", None) is None or worksheet.pictures.count == 0:
            return refs
        for i, picture in enumerate(worksheet.pictures):
            ext = (getattr(picture, "_image_extension", "png") or "png").lower().lstrip(".")
            part_path = getattr(picture, "_source_part_path", None) or f"xl/media/image_s{sheet_num}_{i+1}.{ext}"
            refs.append(
                {
                    "path": part_path,
                    "extension": ext,
                    "content_type": getattr(picture, "_source_content_type", None) or self.image_content_type_for_extension(ext),
                    "picture": picture,
                }
            )
            write_part_once(zipf, part_path, picture.image_bytes)
        return refs

    def _compute_picture_rels(self, picture_refs, chart_rel_count):
        """
        Compute rel ID assignments for pictures (image + optional hyperlink).

        Returns a list of dicts: {ref, image_rel_id, hyperlink_rel_id (or None)}.
        Pictures with a hyperlink_url get two rel IDs (hyperlink first, then image);
        pictures without get one rel ID (image only).
        """
        result = []
        counter = chart_rel_count + 1
        for ref in picture_refs:
            pic = ref["picture"]
            hyperlink_url = getattr(pic, "_hyperlink_url", None)
            if hyperlink_url:
                hyperlink_rel_id = f"rId{counter}"
                counter += 1
                image_rel_id = f"rId{counter}"
                counter += 1
            else:
                hyperlink_rel_id = None
                image_rel_id = f"rId{counter}"
                counter += 1
            result.append({
                "ref": ref,
                "image_rel_id": image_rel_id,
                "hyperlink_rel_id": hyperlink_rel_id,
            })
        return result

    def format_picture_anchors_xml(self, picture_refs, chart_rel_count):
        computed = self._compute_picture_rels(picture_refs, chart_rel_count)
        content = ""
        for i, item in enumerate(computed):
            content += self._format_picture_anchor_xml(
                item["ref"]["picture"],
                item["image_rel_id"],
                item["hyperlink_rel_id"],
                i + 1,
            )
        return content

    def format_picture_relationships_xml(self, picture_refs, chart_rel_count):
        computed = self._compute_picture_rels(picture_refs, chart_rel_count)
        content = ""
        for item in computed:
            pic = item["ref"]["picture"]
            part_name = posixpath.basename(item["ref"]["path"])
            # Hyperlink rel (if present) — must come before image rel to match assigned IDs
            if item["hyperlink_rel_id"] and getattr(pic, "_hyperlink_url", None):
                content += (
                    f'    <Relationship Id="{item["hyperlink_rel_id"]}" '
                    f'Type="{_HYPERLINK_TYPE}" '
                    f'Target="{self._escape_xml(pic._hyperlink_url)}" TargetMode="External"/>\n'
                )
            # Image rel
            content += (
                f'    <Relationship Id="{item["image_rel_id"]}" '
                f'Type="{_IMAGE_TYPE}" '
                f'Target="../media/{part_name}"/>\n'
            )
        return content

    def get_extra_content_type_overrides(self, workbook):
        seen = set()
        overrides = []
        for worksheet in workbook.worksheets:
            for part_path, _bytes, content_type in getattr(worksheet, "_source_drawing_extra_parts", []):
                if not content_type:
                    continue
                part_name = f'/{part_path.lstrip("/")}'
                key = (part_name, content_type)
                if key in seen:
                    continue
                seen.add(key)
                overrides.append(key)
            if getattr(worksheet, "pictures", None) is None:
                continue
            for pic in worksheet.pictures:
                part_path = getattr(pic, "_source_part_path", None)
                content_type = getattr(pic, "_source_content_type", None)
                if not (part_path and content_type):
                    continue
                part_name = f'/{part_path.lstrip("/")}'
                key = (part_name, content_type)
                if key in seen:
                    continue
                seen.add(key)
                overrides.append(key)
        return overrides

    @staticmethod
    def image_content_type_for_extension(extension):
        ext = (extension or "").lower().lstrip(".")
        mapping = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
            "tif": "image/tiff",
            "tiff": "image/tiff",
            "webp": "image/webp",
        }
        return mapping.get(ext, "image/png")

    def _format_picture_anchor_xml(self, picture, image_rel_id, hyperlink_rel_id, object_index):
        from_col_off = int(getattr(picture, '_upper_left_column_offset', 0))
        from_row_off = int(getattr(picture, '_upper_left_row_offset', 0))
        to_col_off = int(getattr(picture, '_lower_right_column_offset', 0))
        to_row_off = int(getattr(picture, '_lower_right_row_offset', 0))
        name = getattr(picture, "name", None) or f"Picture {object_index}"
        object_id = 2000 + object_index

        # editAs attribute on twoCellAnchor
        edit_as = getattr(picture, '_edit_as', None)
        edit_as_attr = f' editAs="{edit_as}"' if edit_as else ''

        content = f'    <xdr:twoCellAnchor{edit_as_attr}>\n'
        content += '        <xdr:from>\n'
        content += f'            <xdr:col>{picture._upper_left_column}</xdr:col>\n'
        content += f'            <xdr:colOff>{from_col_off}</xdr:colOff>\n'
        content += f'            <xdr:row>{picture._upper_left_row}</xdr:row>\n'
        content += f'            <xdr:rowOff>{from_row_off}</xdr:rowOff>\n'
        content += '        </xdr:from>\n'
        content += '        <xdr:to>\n'
        content += f'            <xdr:col>{picture._lower_right_column}</xdr:col>\n'
        content += f'            <xdr:colOff>{to_col_off}</xdr:colOff>\n'
        content += f'            <xdr:row>{picture._lower_right_row}</xdr:row>\n'
        content += f'            <xdr:rowOff>{to_row_off}</xdr:rowOff>\n'
        content += '        </xdr:to>\n'
        content += '        <xdr:pic>\n'
        content += '            <xdr:nvPicPr>\n'

        # cNvPr: include hlinkClick if picture has a hyperlink
        cNvPr_children = ''
        if hyperlink_rel_id and getattr(picture, '_hyperlink_url', None):
            cNvPr_children += (
                f'<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                f' r:id="{hyperlink_rel_id}"/>'
            )
        # Preserve extLst from source (e.g. creationId)
        source_cNvPr_extLst = getattr(picture, '_source_cNvPr_extLst_xml', None)
        if source_cNvPr_extLst:
            cNvPr_children += source_cNvPr_extLst

        if cNvPr_children:
            content += f'                <xdr:cNvPr id="{object_id}" name="{self._escape_xml(name)}">{cNvPr_children}</xdr:cNvPr>\n'
        else:
            content += f'                <xdr:cNvPr id="{object_id}" name="{self._escape_xml(name)}"/>\n'

        # cNvPicPr: picLocks based on no_change_aspect
        no_change_aspect = getattr(picture, '_no_change_aspect', True)
        if no_change_aspect:
            content += '                <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>\n'
        else:
            content += '                <xdr:cNvPicPr/>\n'

        content += '            </xdr:nvPicPr>\n'
        content += '            <xdr:blipFill>\n'

        # blip: include extLst if present
        source_blip_extLst = getattr(picture, '_source_blip_extLst_xml', None)
        if source_blip_extLst:
            content += (
                f'                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                f' r:embed="{image_rel_id}">{source_blip_extLst}</a:blip>\n'
            )
        else:
            content += (
                f'                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                f' r:embed="{image_rel_id}"/>\n'
            )

        content += '                <a:stretch><a:fillRect/></a:stretch>\n'
        content += '            </xdr:blipFill>\n'

        # spPr: use preserved source XML or default
        source_spPr_xml = getattr(picture, '_source_spPr_xml', None)
        if source_spPr_xml:
            content += f'            <xdr:spPr>{source_spPr_xml}</xdr:spPr>\n'
        else:
            content += '            <xdr:spPr>\n'
            content += '                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n'
            content += '            </xdr:spPr>\n'

        content += '        </xdr:pic>\n'
        content += '        <xdr:clientData/>\n'
        content += '    </xdr:twoCellAnchor>\n'
        return content
