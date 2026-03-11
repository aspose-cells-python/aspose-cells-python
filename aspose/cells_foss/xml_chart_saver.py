"""
Aspose.Cells for Python - Chart XML Saver Module

This module provides chart and drawing XML writers for XLSX output.
"""

import posixpath
import re
import uuid
from .chart import ChartType
from .xml_picture_saver import PictureXmlSaver
from .xml_shape_saver import ShapeXmlSaver

# Standard chart style and color XML templates for chartEx charts (treemap, waterfall).
# Excel always writes these companion files alongside chartEx XML.
_CHART_EX_STYLE_XML = (
    '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="102"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000" b="1" kern="1200"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:categoryAxis><cs:chartArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="1"><a:schemeClr val="bg1"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:chartArea><cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000" kern="1200"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef><cs:lineWidthScale>3</cs:lineWidthScale><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln cap="rnd"><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout/><cs:dataPointWireframe><cs:lnRef idx="1"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:dataTable><cs:downBar><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="1"><a:schemeClr val="dk1"><a:tint val="95000"/></a:schemeClr></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:downBar><cs:dropLine><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorBar><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="1"><a:schemeClr val="tx1"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:errorBar><cs:floor><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:floor><cs:gridlineMajor><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="50000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000" kern="1200"/></cs:legend><cs:plotArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="1"><a:schemeClr val="bg1"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea><cs:plotArea3D><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1800" b="1" kern="1200"/></cs:title><cs:trendline><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln cap="rnd"><a:round/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000" kern="1200"/></cs:trendlineLabel><cs:upBar><cs:lnRef idx="1"><a:schemeClr val="tx1"/></cs:lnRef><cs:fillRef idx="1"><a:schemeClr val="dk1"><a:tint val="5000"/></a:schemeClr></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr></cs:upBar><cs:valueAxis><cs:lnRef idx="1"><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:valueAxis><cs:wall><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:wall></cs:chartStyle>'
)

_CHART_EX_COLORS_XML = (
    '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="acrossLinear" id="2"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/></cs:colorStyle>'
)


class ChartXmlSaver:
    """
    Handles writing chart-related XLSX parts:
    - xl/charts/chartN.xml
    - xl/drawings/drawingN.xml
    - xl/drawings/_rels/drawingN.xml.rels
    """

    def __init__(self, escape_xml):
        self._escape_xml = escape_xml
        self._written_parts = set()
        self._picture_writer = PictureXmlSaver(escape_xml)
        self._shape_writer = ShapeXmlSaver(escape_xml)
        # Tracks (part_name, content_type) pairs for chartEx style/colors companion files
        # generated for new (non-source) chartEx charts.
        self._new_chartex_overrides = []

    def _write_part_once(self, zipf, part_path, content_bytes):
        normalized = part_path.lstrip("/")
        if normalized in self._written_parts:
            return
        zipf.writestr(normalized, content_bytes)
        self._written_parts.add(normalized)

    def write_chart_parts(self, zipf, worksheet, sheet_num, next_chart_index):
        """
        Writes drawing/chart XML parts for worksheet charts.

        Args:
            zipf: ZipFile object.
            worksheet: Worksheet object.
            sheet_num (int): 1-based worksheet index.
            next_chart_index (int): Next global chart part index.

        Returns:
            int: Updated next global chart part index.
        """
        has_charts = getattr(worksheet, 'charts', None) is not None and worksheet.charts.count > 0
        has_pictures = getattr(worksheet, 'pictures', None) is not None and worksheet.pictures.count > 0
        has_shapes = getattr(worksheet, 'shapes', None) is not None and worksheet.shapes.count > 0
        has_source_drawing = getattr(worksheet, '_source_drawing_xml', None) is not None
        if not has_charts and not has_pictures and not has_shapes and not has_source_drawing:
            return next_chart_index

        chart_refs = []
        picture_refs = []
        used_indices = set()
        for chart in worksheet.charts if has_charts else []:
            chart_part_path = None
            chart_part_index = None
            source_part_index = getattr(chart, '_source_chart_part_index', None)
            source_part_path = getattr(chart, '_source_chart_part_path', None)
            source_chart_xml = getattr(chart, '_source_chart_xml', None)
            source_is_chart_ex = bool(getattr(chart, '_source_is_chart_ex', False))
            is_chart_ex = source_is_chart_ex or chart.type in (ChartType.WATERFALL, ChartType.TREEMAP, ChartType.BOX_WHISKER, ChartType.SUNBURST, ChartType.HISTOGRAM, ChartType.FUNNEL, ChartType.MAP)
            if source_part_index is not None:
                source_part_index = int(source_part_index)
            elif source_part_path:
                source_chart_match = re.search(r'chart(?:Ex)?(\d+)\.xml$', source_part_path)
                if source_chart_match:
                    source_part_index = int(source_chart_match.group(1))
            if source_chart_xml is not None and source_part_path:
                chart_part_path = source_part_path
                chart_part_index = source_part_index
                if chart_part_index is not None:
                    used_indices.add(chart_part_index)
                    if chart_part_index >= next_chart_index:
                        next_chart_index = chart_part_index + 1
            elif source_part_index is not None and source_part_index not in used_indices:
                chart_part_index = source_part_index
            else:
                while next_chart_index in used_indices:
                    next_chart_index += 1
                chart_part_index = next_chart_index
                next_chart_index += 1
            if chart_part_path is None:
                used_indices.add(chart_part_index)
                if chart_part_index >= next_chart_index:
                    next_chart_index = chart_part_index + 1
                chart_name = f'chartEx{chart_part_index}.xml' if is_chart_ex else f'chart{chart_part_index}.xml'
                chart_part_path = f'xl/charts/{chart_name}'

            if source_chart_xml is not None:
                self._write_part_once(zipf, chart_part_path, source_chart_xml)
                source_chart_rels_xml = getattr(chart, '_source_chart_rels_xml', None)
                if source_chart_rels_xml is not None:
                    chart_rels_path = self._rels_path_for_part(chart_part_path)
                    self._write_part_once(zipf, chart_rels_path, source_chart_rels_xml)
                # Determine the actual drawing path for this sheet. Prefer the stored
                # original path (e.g. "xl/drawings/drawing5.xml" for sheet4) so we
                # don't collide with chart extra parts that occupy the sheet-number slot.
                ws_drawing_path = (
                    getattr(worksheet, '_source_drawing_part_path', None)
                    or f'xl/drawings/drawing{sheet_num}.xml'
                )
                for part_path, part_bytes, _ in getattr(chart, '_source_chart_extra_parts', []):
                    if str(part_path).replace("\\", "/") == ws_drawing_path:
                        continue
                    self._write_part_once(zipf, part_path, part_bytes)
            else:
                if chart.type == ChartType.TREEMAP:
                    chart_xml = self._format_treemap_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.WATERFALL:
                    chart_xml = self._format_waterfall_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.BOX_WHISKER:
                    chart_xml = self._format_box_whisker_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.SUNBURST:
                    chart_xml = self._format_sunburst_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.HISTOGRAM:
                    chart_xml = self._format_histogram_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.FUNNEL:
                    chart_xml = self._format_funnel_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.MAP:
                    chart_xml = self._format_map_chart_ex_xml(chart, worksheet)
                elif chart.type == ChartType.COMBO:
                    chart_xml = self._format_combo_chart_xml(chart, worksheet, chart_part_index)
                elif chart.type == ChartType.SCATTER:
                    chart_xml = self._format_scatter_chart_xml(chart, worksheet, chart_part_index)
                elif chart.type == ChartType.SURFACE:
                    chart_xml = self._format_surface_chart_xml(chart, worksheet, chart_part_index)
                elif chart.type == ChartType.RADAR:
                    chart_xml = self._format_radar_chart_xml(chart, worksheet, chart_part_index)
                else:
                    chart_xml = self._format_chart_xml(chart, worksheet, chart_part_index)
                self._write_part_once(zipf, chart_part_path, chart_xml)
                # For new chartEx charts, write required companion style/colors files and rels.
                if is_chart_ex:
                    self._write_chartex_companion_parts(zipf, chart_part_index, chart_name)
            chart_refs.append({
                "index": chart_part_index,
                "path": chart_part_path,
                "is_chart_ex": is_chart_ex,
            })

        picture_refs = self._picture_writer.collect_picture_refs(
            zipf,
            worksheet,
            sheet_num,
            self._write_part_once,
        )
        shape_refs = self._shape_writer.collect_shape_refs(worksheet)

        # Use the original drawing path when available to avoid collisions.
        drawing_path = (
            getattr(worksheet, '_source_drawing_part_path', None)
            or f'xl/drawings/drawing{sheet_num}.xml'
        )
        drawing_rels_path = self._rels_path_for_part(drawing_path)

        preserve_source_drawing = (
            getattr(worksheet, '_source_drawing_xml', None) is not None
            and all(getattr(chart, '_source_chart_xml', None) is not None for chart in worksheet.charts)
            and not getattr(worksheet, "_drawing_dirty", False)
        )
        if preserve_source_drawing:
            self._write_part_once(zipf, drawing_path, worksheet._source_drawing_xml)
            source_drawing_rels_xml = getattr(worksheet, '_source_drawing_rels_xml', None)
            if source_drawing_rels_xml is not None:
                self._write_part_once(zipf, drawing_rels_path, source_drawing_rels_xml)
            else:
                drawing_rels_xml = self._format_drawing_relationships_xml(chart_refs, picture_refs, shape_refs)
                self._write_part_once(zipf, drawing_rels_path, drawing_rels_xml)
            for part_path, part_bytes, _ in getattr(worksheet, "_source_drawing_extra_parts", []):
                self._write_part_once(zipf, part_path, part_bytes)
        else:
            drawing_xml = self._format_drawing_xml(worksheet, chart_refs, picture_refs, shape_refs)
            self._write_part_once(zipf, drawing_path, drawing_xml)

            drawing_rels_xml = self._format_drawing_relationships_xml(chart_refs, picture_refs, shape_refs)
            self._write_part_once(zipf, drawing_rels_path, drawing_rels_xml)

        return next_chart_index

    def _write_chartex_companion_parts(self, zipf, chart_part_index, chart_name):
        """
        Writes style{n}.xml, colors{n}.xml and their rels for a new chartEx chart.

        Excel always includes these companion files; without them the file fails
        to open.  chart_name is the basename of the chartEx part (e.g. 'chartEx1.xml').
        """
        style_path = f'xl/charts/style{chart_part_index}.xml'
        colors_path = f'xl/charts/colors{chart_part_index}.xml'
        rels_path = f'xl/charts/_rels/{chart_name}.rels'

        self._write_part_once(zipf, style_path, _CHART_EX_STYLE_XML.encode('utf-8'))
        self._write_part_once(zipf, colors_path, _CHART_EX_COLORS_XML.encode('utf-8'))

        chart_ex_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId2"'
            f' Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"'
            f' Target="colors{chart_part_index}.xml"/>'
            f'<Relationship Id="rId1"'
            f' Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle"'
            f' Target="style{chart_part_index}.xml"/>'
            '</Relationships>'
        )
        self._write_part_once(zipf, rels_path, chart_ex_rels.encode('utf-8'))

        style_part_name = f'/{style_path}'
        colors_part_name = f'/{colors_path}'
        if style_part_name not in {p for p, _ in self._new_chartex_overrides}:
            self._new_chartex_overrides.append(
                (style_part_name, 'application/vnd.ms-office.chartstyle+xml')
            )
        if colors_part_name not in {p for p, _ in self._new_chartex_overrides}:
            self._new_chartex_overrides.append(
                (colors_part_name, 'application/vnd.ms-office.chartcolorstyle+xml')
            )

    def get_extra_content_type_overrides(self, workbook):
        """
        Collects extra chart-related content type overrides from loaded chart parts.

        Returns:
            list[tuple[str, str]]: [(part_name_with_leading_slash, content_type), ...]
        """
        # Build the set of worksheet drawing paths that the saver will write.
        # These use the original stored path when available to match write_chart_parts.
        # Chart extra parts whose path matches a worksheet drawing must be skipped
        # to avoid content-type entries for files that won't actually be written as
        # chart extra parts.
        worksheet_drawing_paths = set()
        for i, ws in enumerate(workbook.worksheets):
            has_drawing = (
                (getattr(ws, 'charts', None) and ws.charts.count > 0)
                or getattr(ws, '_source_drawing_xml', None) is not None
                or (getattr(ws, 'pictures', None) and ws.pictures.count > 0)
            )
            if has_drawing:
                stored = getattr(ws, '_source_drawing_part_path', None)
                p = stored or f'xl/drawings/drawing{i+1}.xml'
                worksheet_drawing_paths.add(f'/{p}')

        seen = set()
        overrides = []
        for worksheet in workbook.worksheets:
            if getattr(worksheet, 'charts', None) is None:
                continue
            for chart in worksheet.charts:
                for part_path, _bytes, content_type in getattr(chart, '_source_chart_extra_parts', []):
                    if not content_type:
                        continue
                    part_name = f'/{part_path.lstrip("/")}'
                    if part_name in worksheet_drawing_paths:
                        continue
                    key = (part_name, content_type)
                    if key in seen:
                        continue
                    seen.add(key)
                    overrides.append(key)
                main_content_type = getattr(chart, '_source_chart_content_type', None)
                main_part_path = getattr(chart, '_source_chart_part_path', None)
                if main_content_type and main_part_path:
                    main_part_name = f'/{main_part_path.lstrip("/")}'
                    key = (main_part_name, main_content_type)
                    if key not in seen:
                        seen.add(key)
                        overrides.append(key)
        for part_name, content_type in self._picture_writer.get_extra_content_type_overrides(workbook):
            key = (part_name, content_type)
            if key in seen:
                continue
            seen.add(key)
            overrides.append(key)
        # Include style/colors companion files for newly generated chartEx charts.
        for part_name, content_type in self._new_chartex_overrides:
            key = (part_name, content_type)
            if key in seen:
                continue
            seen.add(key)
            overrides.append(key)
        return overrides

    def _rels_path_for_part(self, part_path):
        part_dir = posixpath.dirname(part_path)
        part_name = posixpath.basename(part_path)
        return posixpath.join(part_dir, '_rels', f'{part_name}.rels')

    def _format_drawing_xml(self, worksheet, chart_refs, picture_refs, shape_refs=None):
        """Formats xl/drawings/drawingN.xml for worksheet charts.

        Root element only declares xdr and a namespaces, matching Excel-generated files.
        Namespace declarations for c, r, mc, cx are placed inline where needed.
        """
        if shape_refs is None:
            shape_refs = []
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'

        c_ns = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

        for i, chart in enumerate(worksheet.charts):
            rel_id = f"rId{i+1}"
            object_id = i + 2
            chart_ref = chart_refs[i]
            chart_display_index = chart_ref['index'] if chart_ref['index'] is not None else (i + 1)
            chart_name = f"Chart {chart_display_index}"
            from_col_off = int(getattr(chart, '_upper_left_column_offset', 0))
            from_row_off = int(getattr(chart, '_upper_left_row_offset', 0))
            to_col_off = int(getattr(chart, '_lower_right_column_offset', 0))
            to_row_off = int(getattr(chart, '_lower_right_row_offset', 0))
            content += '    <xdr:twoCellAnchor>\n'
            content += '        <xdr:from>\n'
            content += f'            <xdr:col>{chart._upper_left_column}</xdr:col>\n'
            content += f'            <xdr:colOff>{from_col_off}</xdr:colOff>\n'
            content += f'            <xdr:row>{chart._upper_left_row}</xdr:row>\n'
            content += f'            <xdr:rowOff>{from_row_off}</xdr:rowOff>\n'
            content += '        </xdr:from>\n'
            content += '        <xdr:to>\n'
            content += f'            <xdr:col>{chart._lower_right_column}</xdr:col>\n'
            content += f'            <xdr:colOff>{to_col_off}</xdr:colOff>\n'
            content += f'            <xdr:row>{chart._lower_right_row}</xdr:row>\n'
            content += f'            <xdr:rowOff>{to_row_off}</xdr:rowOff>\n'
            content += '        </xdr:to>\n'
            if chart_ref["is_chart_ex"]:
                content += self._format_chart_ex_graphic_frame(
                    object_id, chart_name, rel_id,
                    chart._upper_left_column, chart._upper_left_row,
                    chart._lower_right_column, chart._lower_right_row,
                )
            else:
                graphic_uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                chart_tag = f'<c:chart {c_ns} {r_ns} r:id="{rel_id}"/>'
                content += '        <xdr:graphicFrame macro="">\n'
                content += '            <xdr:nvGraphicFramePr>\n'
                content += f'                <xdr:cNvPr id="{object_id}" name="{self._escape_xml(chart_name)}"/>\n'
                content += '                <xdr:cNvGraphicFramePr/>\n'
                content += '            </xdr:nvGraphicFramePr>\n'
                content += '            <xdr:xfrm>\n'
                content += '                <a:off x="0" y="0"/>\n'
                content += '                <a:ext cx="0" cy="0"/>\n'
                content += '            </xdr:xfrm>\n'
                content += '            <a:graphic>\n'
                content += f'                <a:graphicData uri="{graphic_uri}">\n'
                content += f'                    {chart_tag}\n'
                content += '                </a:graphicData>\n'
                content += '            </a:graphic>\n'
                content += '        </xdr:graphicFrame>\n'
            content += '        <xdr:clientData/>\n'
            content += '    </xdr:twoCellAnchor>\n'

        content += self._picture_writer.format_picture_anchors_xml(picture_refs, len(chart_refs))

        # Assign shape hyperlink rel IDs after chart + picture rels, then write anchors.
        picture_rel_count = sum(
            2 if getattr(ref["picture"], "_hyperlink_url", None) else 1
            for ref in picture_refs
        )
        self._shape_writer.assign_hyperlink_rel_ids(
            shape_refs, len(chart_refs) + picture_rel_count + 1)
        content += self._shape_writer.format_shape_anchors_xml(
            shape_refs, len(chart_refs), len(picture_refs))

        content += '</xdr:wsDr>\n'
        return content

    def _format_chart_ex_graphic_frame(
        self, object_id, chart_name, rel_id,
        from_col=0, from_row=0, to_col=10, to_row=20
    ):
        """
        Formats the mc:AlternateContent wrapper required for chartEx (treemap/waterfall) charts.

        Excel requires chartEx charts to be wrapped in mc:AlternateContent so that older
        Excel versions display a fallback placeholder shape instead of crashing.

        The mc namespace is declared locally on mc:AlternateContent (not on the root wsDr
        element), matching exactly how Excel generates these files.
        """
        mc_ns = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        cx_uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
        cx_ns = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
        r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        cx1_ns = 'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"'

        # Approximate EMU coordinates for the fallback shape.
        # Standard defaults: ~609600 EMU per column, ~180000 EMU per row.
        emu_per_col = 609600
        emu_per_row = 180000
        fb_x = from_col * emu_per_col
        fb_y = from_row * emu_per_row
        fb_cx = max((to_col - from_col) * emu_per_col, emu_per_col)
        fb_cy = max((to_row - from_row) * emu_per_row, emu_per_row)

        # Generate a unique creation ID GUID, matching what Excel writes for chartEx charts.
        creation_id = '{' + str(uuid.uuid4()).upper() + '}'
        a16_ns = 'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"'
        ext_uri = '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'

        content = f'        <mc:AlternateContent {mc_ns}>\n'
        content += f'            <mc:Choice {cx1_ns} Requires="cx1">\n'
        content += '                <xdr:graphicFrame macro="">\n'
        content += '                    <xdr:nvGraphicFramePr>\n'
        content += f'                        <xdr:cNvPr id="{object_id}" name="{self._escape_xml(chart_name)}">\n'
        content += '                            <a:extLst>\n'
        content += f'                                <a:ext uri="{ext_uri}">\n'
        content += f'                                    <a16:creationId {a16_ns} id="{creation_id}"/>\n'
        content += '                                </a:ext>\n'
        content += '                            </a:extLst>\n'
        content += '                        </xdr:cNvPr>\n'
        content += '                        <xdr:cNvGraphicFramePr/>\n'
        content += '                    </xdr:nvGraphicFramePr>\n'
        content += '                    <xdr:xfrm>\n'
        content += '                        <a:off x="0" y="0"/>\n'
        content += '                        <a:ext cx="0" cy="0"/>\n'
        content += '                    </xdr:xfrm>\n'
        content += '                    <a:graphic>\n'
        content += f'                        <a:graphicData uri="{cx_uri}">\n'
        content += f'                            <cx:chart {cx_ns} {r_ns} r:id="{rel_id}"/>\n'
        content += '                        </a:graphicData>\n'
        content += '                    </a:graphic>\n'
        content += '                </xdr:graphicFrame>\n'
        content += '            </mc:Choice>\n'
        content += '            <mc:Fallback>\n'
        content += '                <xdr:sp macro="" textlink="">\n'
        content += '                    <xdr:nvSpPr>\n'
        content += '                        <xdr:cNvPr id="0" name=""/>\n'
        content += '                        <xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr>\n'
        content += '                    </xdr:nvSpPr>\n'
        content += '                    <xdr:spPr>\n'
        content += f'                        <a:xfrm><a:off x="{fb_x}" y="{fb_y}"/><a:ext cx="{fb_cx}" cy="{fb_cy}"/></a:xfrm>\n'
        content += '                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n'
        content += '                        <a:solidFill><a:prstClr val="white"/></a:solidFill>\n'
        content += '                        <a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln>\n'
        content += '                    </xdr:spPr>\n'
        content += '                    <xdr:txBody>\n'
        content += '                        <a:bodyPr vertOverflow="clip" horzOverflow="clip"/>\n'
        content += '                        <a:lstStyle/>\n'
        content += '                        <a:p><a:r>'
        content += '<a:rPr lang="zh-CN" altLang="en-US" sz="1100"/>'
        content += '<a:t>This chart isn\'t available in your version of Excel.\r\n\r\n'
        content += 'Editing this shape or saving this workbook into a different file format will permanently break the chart.</a:t>'
        content += '</a:r></a:p>\n'
        content += '                    </xdr:txBody>\n'
        content += '                </xdr:sp>\n'
        content += '            </mc:Fallback>\n'
        content += '        </mc:AlternateContent>\n'
        return content

    def _format_drawing_relationships_xml(self, chart_refs, picture_refs, shape_refs=None):
        """Formats xl/drawings/_rels/drawingN.xml.rels for chart parts."""
        if shape_refs is None:
            shape_refs = []
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        for i, chart_ref in enumerate(chart_refs):
            rel_type = (
                "http://schemas.microsoft.com/office/2014/relationships/chartEx"
                if chart_ref["is_chart_ex"]
                else "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
            )
            part_name = posixpath.basename(chart_ref["path"])
            content += f'    <Relationship Id="rId{i+1}" Type="{rel_type}" Target="../charts/{part_name}"/>\n'
        content += self._picture_writer.format_picture_relationships_xml(picture_refs, len(chart_refs))
        content += self._shape_writer.format_shape_relationships_xml(
            shape_refs, len(chart_refs), len(picture_refs))
        content += '</Relationships>\n'
        return content

    # -----------------------------------------------------------------------
    # Standard chart XML (line/bar/pie/area/stock)
    # -----------------------------------------------------------------------

    def _format_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for supported single-type chart types."""
        if chart.type == ChartType.STOCK:
            return self._format_stock_chart_xml(chart, worksheet, chart_index)

        if chart.type not in (ChartType.LINE, ChartType.BAR, ChartType.PIE, ChartType.AREA):
            raise NotImplementedError("Only line, bar, pie, area and stock charts are currently supported")

        cat_ax_id = 70000000 + (chart_index * 2)
        val_ax_id = cat_ax_id + 1
        ser_ax_id = val_ax_id + 1

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'

        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)

        if chart.is_3d and chart.type in (ChartType.LINE, ChartType.BAR, ChartType.AREA):
            content += self._format_view_3d_xml(chart)

        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'
        if chart.type == ChartType.LINE:
            chart_element = 'line3DChart' if chart.is_3d else 'lineChart'
        elif chart.type == ChartType.BAR:
            chart_element = 'bar3DChart' if chart.is_3d else 'barChart'
        elif chart.type == ChartType.AREA:
            chart_element = 'area3DChart' if chart.is_3d else 'areaChart'
        else:
            if chart.is_of_pie:
                chart_element = 'ofPieChart'
            else:
                chart_element = 'pie3DChart' if chart.is_3d else 'pieChart'
        content += f'            <c:{chart_element}>\n'
        # Element order follows CT_* chart schema sequences.
        if chart.type == ChartType.BAR:
            content += f'                <c:barDir val="{self._escape_xml(chart.bar_direction)}"/>\n'
            content += f'                <c:grouping val="{self._escape_xml(chart.grouping)}"/>\n'
            content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
        elif chart.type in (ChartType.LINE, ChartType.AREA):
            content += f'                <c:grouping val="{self._escape_xml(chart.grouping)}"/>\n'
            content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
        else:
            content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'

        for i, series in enumerate(chart.n_series):
            content += self._format_line_series_xml(worksheet, chart, series, i)

        if chart.type == ChartType.LINE and not chart.is_3d:
            content += f'                <c:smooth val="{1 if chart.smooth else 0}"/>\n'
        if chart.type == ChartType.BAR:
            content += f'                <c:gapWidth val="{int(chart.gap_width)}"/>\n'
            content += f'                <c:overlap val="{int(chart.overlap)}"/>\n'
        if chart.type == ChartType.PIE:
            if chart.is_of_pie:
                content += f'                <c:ofPieType val="{self._escape_xml(chart.of_pie_type)}"/>\n'
                content += f'                <c:gapWidth val="{int(chart.gap_width)}"/>\n'
                content += f'                <c:secondPieSize val="{int(chart.second_pie_size)}"/>\n'
            content += f'                <c:firstSliceAng val="{int(chart.first_slice_angle)}"/>\n'
        if chart.is_3d and chart.type in (ChartType.LINE, ChartType.BAR, ChartType.AREA):
            content += f'                <c:gapDepth val="{int(chart.gap_depth)}"/>\n'
        if chart.type != ChartType.PIE:
            content += f'                <c:axId val="{cat_ax_id}"/>\n'
            content += f'                <c:axId val="{val_ax_id}"/>\n'
            if chart.is_3d:
                content += f'                <c:axId val="{ser_ax_id}"/>\n'
        content += f'            </c:{chart_element}>\n'

        if chart.type != ChartType.PIE:
            if chart.axes:
                for ax in chart.axes:
                    content += self._format_axis_xml(ax)
            else:
                content += self._format_cat_axis_xml(cat_ax_id, val_ax_id, is_3d=chart.is_3d)
                content += self._format_val_axis_xml(val_ax_id, cat_ax_id)
                if chart.is_3d:
                    content += self._format_ser_axis_xml(ser_ax_id, val_ax_id)
        content += '        </c:plotArea>\n'

        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    def _format_stock_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for stock chart variants."""
        style = getattr(chart, "_stock_style", "high_low_close")
        allowed_styles = {
            "high_low_close",
            "open_high_low_close",
            "volume_high_low_close",
            "volume_open_high_low_close",
        }
        if style not in allowed_styles:
            raise ValueError(f"Unsupported stock_style: {style}")

        expected_series = {
            "high_low_close": 3,
            "open_high_low_close": 4,
            "volume_high_low_close": 4,
            "volume_open_high_low_close": 5,
        }[style]
        if chart.n_series.count < expected_series:
            raise ValueError(
                f"stock_style '{style}' requires at least {expected_series} series, got {chart.n_series.count}"
            )

        # Axis layout:
        # - non-volume: cat + val
        # - volume variants: bar(cat+val) + stock(cat+val-secondary)
        cat_ax_id = 70000000 + (chart_index * 4)
        val_ax_id = cat_ax_id + 1
        sec_cat_ax_id = cat_ax_id + 2
        sec_val_ax_id = cat_ax_id + 3

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'

        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)

        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'

        series_offset = 0
        uses_volume = style.startswith("volume_")

        if uses_volume:
            # First series: volume as clustered columns.
            volume_series = chart.n_series[0]
            content += '            <c:barChart>\n'
            content += '                <c:barDir val="col"/>\n'
            content += '                <c:grouping val="clustered"/>\n'
            content += '                <c:varyColors val="0"/>\n'
            content += self._format_line_series_xml(worksheet, chart, volume_series, 0)
            content += '                <c:gapWidth val="150"/>\n'
            content += f'                <c:axId val="{cat_ax_id}"/>\n'
            content += f'                <c:axId val="{val_ax_id}"/>\n'
            content += '            </c:barChart>\n'
            series_offset = 1

        # Remaining OHLC/HLC data in stockChart.
        content += '            <c:stockChart>\n'
        content += '                <c:varyColors val="0"/>\n'

        stock_series_count = expected_series - series_offset
        for i in range(stock_series_count):
            series = chart.n_series[series_offset + i]
            content += self._format_line_series_xml(worksheet, chart, series, series_offset + i)

        content += '                <c:hiLowLines>\n'
        content += '                    <c:spPr><a:ln/></c:spPr>\n'
        content += '                </c:hiLowLines>\n'

        if style in ("open_high_low_close", "volume_open_high_low_close"):
            content += '                <c:upDownBars>\n'
            content += '                    <c:gapWidth val="150"/>\n'
            content += '                    <c:upBars><c:spPr/></c:upBars>\n'
            content += '                    <c:downBars><c:spPr/></c:downBars>\n'
            content += '                </c:upDownBars>\n'

        content += f'                <c:axId val="{sec_cat_ax_id if uses_volume else cat_ax_id}"/>\n'
        content += f'                <c:axId val="{sec_val_ax_id if uses_volume else val_ax_id}"/>\n'
        content += '            </c:stockChart>\n'

        if uses_volume:
            content += self._format_cat_axis_xml(cat_ax_id, val_ax_id)
            content += self._format_val_axis_xml(val_ax_id, cat_ax_id)

            # Secondary value axis for price lines.
            content += self._format_cat_axis_xml(sec_cat_ax_id, sec_val_ax_id)
            content += self._format_val_axis_xml(sec_val_ax_id, sec_cat_ax_id)
        else:
            content += self._format_cat_axis_xml(cat_ax_id, val_ax_id)
            content += self._format_val_axis_xml(val_ax_id, cat_ax_id)

        content += '        </c:plotArea>\n'
        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Scatter chart XML
    # -----------------------------------------------------------------------

    def _format_scatter_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for a pure scatter chart."""
        cat_ax_id = 70000000 + (chart_index * 2)
        val_ax_id = cat_ax_id + 1

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'
        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)
        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'
        content += '            <c:scatterChart>\n'
        content += f'                <c:scatterStyle val="{self._escape_xml(chart._scatter_style)}"/>\n'
        content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
        for i, series in enumerate(chart.n_series):
            content += self._format_scatter_series_xml(worksheet, chart, series, i)
        content += f'                <c:axId val="{cat_ax_id}"/>\n'
        content += f'                <c:axId val="{val_ax_id}"/>\n'
        content += '            </c:scatterChart>\n'
        if chart.axes:
            for ax in chart.axes:
                content += self._format_axis_xml(ax)
        else:
            content += self._format_val_axis_xml(cat_ax_id, val_ax_id)
            content += self._format_val_axis_xml(val_ax_id, cat_ax_id)
        content += '        </c:plotArea>\n'
        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Surface chart XML
    # -----------------------------------------------------------------------

    def _format_surface_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for a surface or surface3D chart."""
        # Use three axis IDs: cat (X), ser (Y/depth), val (Z/value)
        cat_ax_id = 80000000 + (chart_index * 3)
        ser_ax_id = cat_ax_id + 1
        val_ax_id = cat_ax_id + 2

        # surface3DChart when is_3d (default); surfaceChart for contour/flat
        elem_name = 'surface3DChart' if chart._is_3d else 'surfaceChart'
        wireframe_val = '1' if chart._wireframe else '0'

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'
        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)
        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'
        content += f'            <c:{elem_name}>\n'
        content += f'                <c:wireframe val="{wireframe_val}"/>\n'
        content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
        surface_series_index = 0
        for series in chart.n_series:
            values_formulas = self._expand_surface_values_formulas(series.values, worksheet.name)
            for values_formula in values_formulas:
                content += self._format_line_series_xml(
                    worksheet,
                    chart,
                    series,
                    surface_series_index,
                    values_formula_override=values_formula,
                    order_override=surface_series_index,
                )
                surface_series_index += 1
        content += f'                <c:axId val="{cat_ax_id}"/>\n'
        content += f'                <c:axId val="{ser_ax_id}"/>\n'
        content += f'                <c:axId val="{val_ax_id}"/>\n'
        content += f'            </c:{elem_name}>\n'
        # Surface charts need catAx + serAx + valAx
        content += self._format_cat_axis_xml(cat_ax_id, val_ax_id)
        content += self._format_ser_axis_xml(ser_ax_id, val_ax_id)
        content += self._format_val_axis_xml(val_ax_id, cat_ax_id)
        content += '        </c:plotArea>\n'
        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Radar chart XML
    # -----------------------------------------------------------------------

    def _format_radar_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for a radar chart."""
        cat_ax_id = 90000000 + (chart_index * 2)
        val_ax_id = cat_ax_id + 1

        radar_style = getattr(chart, '_radar_style', 'marker')

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'
        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)
        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'
        content += '            <c:radarChart>\n'
        content += f'                <c:radarStyle val="{self._escape_xml(radar_style)}"/>\n'
        content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
        for i, series in enumerate(chart.n_series):
            content += self._format_line_series_xml(worksheet, chart, series, i)
        content += f'                <c:axId val="{cat_ax_id}"/>\n'
        content += f'                <c:axId val="{val_ax_id}"/>\n'
        content += '            </c:radarChart>\n'
        if chart.axes:
            for ax in chart.axes:
                content += self._format_axis_xml(ax)
        else:
            content += self._format_cat_axis_xml(cat_ax_id, val_ax_id)
            content += self._format_val_axis_xml(val_ax_id, cat_ax_id)
        content += '        </c:plotArea>\n'
        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Combo chart XML
    # -----------------------------------------------------------------------

    def _format_combo_chart_xml(self, chart, worksheet, chart_index):
        """Formats xl/charts/chartN.xml for a combo chart."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        content += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        content += '    <c:chart>\n'

        if chart.title is not None:
            content += self._format_chart_title_xml(chart.title)

        content += '        <c:plotArea>\n'
        content += '            <c:layout/>\n'

        # Render each sub-chart
        if chart._sub_charts:
            for sub in chart._sub_charts:
                content += self._format_sub_chart_xml(chart, worksheet, sub)
        else:
            # Fallback: group all series by chart_type
            content += self._format_combo_fallback_xml(chart, worksheet, chart_index)

        # Axes
        if chart.axes:
            axis_cross_map = self._build_combo_axis_cross_map(chart)
            for ax in chart.axes:
                content += self._format_axis_xml(ax, cross_ax_override=axis_cross_map.get(ax.axis_id))
        else:
            # Derive unique axis ID pairs from sub-charts (in first-seen order).
            # Each pair (ax_ids[0], ax_ids[1]) becomes (catAx, valAx).
            # If sub-charts declare additional secondary ax_id pairs, add those too.
            seen_pairs = []
            seen_ids = set()
            for sub in chart._sub_charts:
                ids = sub.get("ax_ids", [])
                if len(ids) >= 2:
                    pair = (ids[0], ids[1])
                    if pair not in seen_pairs:
                        seen_pairs.append(pair)
                        seen_ids.update(pair)

            if seen_pairs:
                for cat_id, val_id in seen_pairs:
                    content += self._format_cat_axis_xml(cat_id, val_id)
                    content += self._format_val_axis_xml(val_id, cat_id)
            else:
                ax_base = 70000000 + (chart_index * 4)
                content += self._format_cat_axis_xml(ax_base, ax_base + 1)
                content += self._format_val_axis_xml(ax_base + 1, ax_base)

        content += '        </c:plotArea>\n'
        content += self._format_legend_xml(chart)
        content += '        <c:plotVisOnly val="1"/>\n'
        if chart._disp_blanks_as != "gap":
            content += f'        <c:dispBlanksAs val="{self._escape_xml(chart._disp_blanks_as)}"/>\n'
        content += '    </c:chart>\n'
        content += '</c:chartSpace>\n'
        return content

    def _format_sub_chart_xml(self, chart, worksheet, sub):
        """Formats a single sub-chart element within a combo chart plotArea."""
        sub_type = sub["type"]
        ax_ids = sub.get("ax_ids", [])

        if sub_type == ChartType.BAR:
            elem = "barChart"
            bar_dir = sub.get("bar_direction", chart.bar_direction)
            grouping = sub.get("grouping", "clustered")
            vary = sub.get("vary_colors", False)
            gap_width = sub.get("gap_width", 150)
            content = f'            <c:{elem}>\n'
            content += f'                <c:barDir val="{self._escape_xml(bar_dir)}"/>\n'
            content += f'                <c:grouping val="{self._escape_xml(grouping)}"/>\n'
            content += f'                <c:varyColors val="{1 if vary else 0}"/>\n'
            for ser_idx in sub.get("series", []):
                if ser_idx < chart.n_series.count:
                    series = chart.n_series[ser_idx]
                    actual_i = series._series_idx if series._series_idx is not None else ser_idx
                    content += self._format_line_series_xml(worksheet, chart, series, actual_i)
            content += f'                <c:gapWidth val="{gap_width}"/>\n'
            for ax_id in ax_ids:
                content += f'                <c:axId val="{ax_id}"/>\n'
            content += f'            </c:{elem}>\n'
            return content

        elif sub_type == ChartType.SCATTER:
            scatter_style = sub.get("scatter_style", chart._scatter_style)
            vary = sub.get("vary_colors", False)
            content = '            <c:scatterChart>\n'
            content += f'                <c:scatterStyle val="{self._escape_xml(scatter_style)}"/>\n'
            content += f'                <c:varyColors val="{1 if vary else 0}"/>\n'
            for ser_idx in sub.get("series", []):
                if ser_idx < chart.n_series.count:
                    series = chart.n_series[ser_idx]
                    actual_i = series._series_idx if series._series_idx is not None else ser_idx
                    content += self._format_scatter_series_xml(worksheet, chart, series, actual_i)
            for ax_id in ax_ids:
                content += f'                <c:axId val="{ax_id}"/>\n'
            content += '            </c:scatterChart>\n'
            return content

        elif sub_type == ChartType.LINE:
            grouping = sub.get("grouping", "standard")
            vary = sub.get("vary_colors", False)
            content = '            <c:lineChart>\n'
            content += f'                <c:grouping val="{self._escape_xml(grouping)}"/>\n'
            content += f'                <c:varyColors val="{1 if vary else 0}"/>\n'
            for ser_idx in sub.get("series", []):
                if ser_idx < chart.n_series.count:
                    series = chart.n_series[ser_idx]
                    actual_i = series._series_idx if series._series_idx is not None else ser_idx
                    content += self._format_line_series_xml(worksheet, chart, series, actual_i)
            for ax_id in ax_ids:
                content += f'                <c:axId val="{ax_id}"/>\n'
            content += '            </c:lineChart>\n'
            return content

        else:
            return ""

    def _format_combo_fallback_xml(self, chart, worksheet, chart_index):
        """Generates sub-charts by grouping series by chart_type (no sub_charts defined)."""
        ax_base = 70000000 + (chart_index * 4)
        bar_series = [s for s in chart.n_series if (s.chart_type or chart.type) == ChartType.BAR]
        scatter_series = [s for s in chart.n_series if (s.chart_type or chart.type) == ChartType.SCATTER]
        line_series = [s for s in chart.n_series if (s.chart_type or chart.type) == ChartType.LINE]

        content = ""
        if bar_series:
            content += '            <c:barChart>\n'
            content += f'                <c:barDir val="{self._escape_xml(chart.bar_direction)}"/>\n'
            content += f'                <c:grouping val="{self._escape_xml(chart.grouping)}"/>\n'
            content += f'                <c:varyColors val="{1 if chart.vary_colors else 0}"/>\n'
            for i, series in enumerate(bar_series):
                content += self._format_line_series_xml(worksheet, chart, series, i)
            content += f'                <c:gapWidth val="{int(chart.gap_width)}"/>\n'
            content += f'                <c:axId val="{ax_base}"/>\n'
            content += f'                <c:axId val="{ax_base + 1}"/>\n'
            content += '            </c:barChart>\n'

        if line_series:
            content += '            <c:lineChart>\n'
            content += f'                <c:grouping val="{self._escape_xml(chart.grouping)}"/>\n'
            content += f'                <c:varyColors val="0"/>\n'
            for i, series in enumerate(line_series):
                content += self._format_line_series_xml(worksheet, chart, series, len(bar_series) + i)
            content += f'                <c:axId val="{ax_base}"/>\n'
            content += f'                <c:axId val="{ax_base + 1}"/>\n'
            content += '            </c:lineChart>\n'

        if scatter_series:
            content += '            <c:scatterChart>\n'
            content += f'                <c:scatterStyle val="{self._escape_xml(chart._scatter_style)}"/>\n'
            content += '                <c:varyColors val="0"/>\n'
            for i, series in enumerate(scatter_series):
                content += self._format_scatter_series_xml(worksheet, chart, series, len(bar_series) + len(line_series) + i)
            content += f'                <c:axId val="{ax_base + 2}"/>\n'
            content += f'                <c:axId val="{ax_base + 3}"/>\n'
            content += '            </c:scatterChart>\n'

        return content

    # -----------------------------------------------------------------------
    # Box & Whisker (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_box_whisker_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for box-and-whisker charts."""
        # Prefer _xlchart.v1.x defined names (injected by xml_saver) for chartEx refs.
        series_name_map = getattr(chart, '_chartex_series_name_map', None)
        quartile_method = chart.quartile_method  # 'inclusive' or 'exclusive'

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula
            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f dir="row">{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="val">\n'
            content += f'                <cx:f dir="row">{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="boxWhisker" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                    <cx:layoutPr>'
            content += (
                f'<cx:visibility meanLine="{1 if chart.box_show_mean_line else 0}" '
                f'meanMarker="{1 if chart.box_show_mean_marker else 0}" '
                f'nonoutliers="{1 if chart.box_show_inner_points else 0}" '
                f'outliers="{1 if chart.box_show_outlier_points else 0}" />'
            )
            content += f'<cx:statistics quartileMethod="{self._escape_xml(quartile_method)}" />'
            content += '</cx:layoutPr>\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        content += f'            <cx:axis id="0"><cx:catScaling gapWidth="{int(chart.box_gap_width)}" /><cx:tickLabels /></cx:axis>\n'
        content += '            <cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Map (chartEx regionMap) XML
    # -----------------------------------------------------------------------

    def _format_map_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for map (regionMap) charts.

        Note: Map charts rely on cx:geography geocache binary data (Bing tiles)
        that is embedded by Excel and cannot be generated programmatically.
        This formatter produces a structurally-valid chartEx XML that Excel will
        open, but the map will not render geography until Excel geocodes the data.
        Round-tripped files (loaded from existing .xlsx) preserve the full source
        XML via _source_chart_xml and bypass this formatter entirely.
        """
        series_name_map = getattr(chart, '_chartex_series_name_map', None)
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula
            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="colorVal">\n'
            content += f'                <cx:f>{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="regionMap" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        # Region map chartEx payloads emitted by Excel do not include cx:axis nodes.
        # Writing axis nodes here can trigger repair in strict Excel builds.
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Funnel (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_funnel_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for funnel charts."""
        # Prefer _xlchart.v1.x defined names (injected by xml_saver) for chartEx refs.
        series_name_map = getattr(chart, '_chartex_series_name_map', None)
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula
            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="val">\n'
            content += f'                <cx:f>{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="funnel" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        # Funnel uses a single catScaling axis at id="1" (no valScaling axis)
        content += '            <cx:axis id="1"><cx:catScaling /><cx:tickLabels /></cx:axis>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Histogram (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_histogram_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for histogram charts."""
        # Prefer _xlchart.v1.x defined names (injected by xml_saver) for chartEx refs.
        series_name_map = getattr(chart, '_chartex_series_name_map', None)
        interval_closed = getattr(chart, '_histogram_interval_closed', 'r')
        # Build the <cx:binning .../> attributes string.
        # Keep generated histogram binning conservative for compatibility.
        # Excel-generated files we round-trip reliably use only intervalClosed.
        binning_attrs = f' intervalClosed="{self._escape_xml(interval_closed)}"'

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula
            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="val">\n'
            content += f'                <cx:f>{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="clusteredColumn" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += f'                    <cx:layoutPr><cx:binning{binning_attrs} /></cx:layoutPr>\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        content += '            <cx:axis id="0"><cx:catScaling /><cx:tickLabels /></cx:axis>\n'
        content += '            <cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Sunburst (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_sunburst_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for sunburst charts."""
        # Prefer _xlchart.v1.x defined names (injected by xml_saver) for chartEx refs.
        series_name_map = getattr(chart, '_chartex_series_name_map', None)
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula
            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="size">\n'
            content += f'                <cx:f>{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="sunburst" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Waterfall (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_waterfall_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for waterfall charts."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None
            content += f'        <cx:data id="{i}">\n'
            if categories_formula:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(categories_formula)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="val">\n'
            content += f'                <cx:f>{self._escape_xml(values_formula)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        if chart.title is not None:
            content += '        <cx:title pos="t" align="ctr" overlay="0">\n'
            content += '            <cx:tx><cx:txData>'
            content += f'<cx:v>{self._escape_xml(chart.title)}</cx:v>'
            content += '</cx:txData></cx:tx>\n'
            content += '        </cx:title>\n'
        else:
            content += '        <cx:title pos="t" align="ctr" overlay="0" />\n'
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        has_subtotals = False
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            content += f'                <cx:series layoutId="waterfall"{hidden_attr}>\n'
            if series.name:
                content += '                    <cx:tx><cx:txData>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                    <cx:layoutPr>'
            if series.is_subtotal:
                content += '<cx:subtotals />'
                has_subtotals = True
            if chart.show_connector_lines:
                content += '<cx:connectorLines visible="1" />'
            content += '</cx:layoutPr>\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        content += '            <cx:axis id="0"><cx:catScaling /><cx:tickLabels /></cx:axis>\n'
        content += '            <cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        chart._has_subtotals = has_subtotals
        return content

    # -----------------------------------------------------------------------
    # Treemap (chartEx) XML
    # -----------------------------------------------------------------------

    def _format_treemap_chart_ex_xml(self, chart, worksheet):
        """Formats a minimal chartEx payload for treemap charts."""
        # Use _xlchart.v1.x defined names in cx:f when available (pre-computed by
        # xml_saver._inject_chartex_defined_names so they appear in workbook.xml).
        series_name_map = getattr(chart, '_chartex_series_name_map', None)

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        content += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        content += 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">\n'
        content += '    <cx:chartData>\n'
        for i, series in enumerate(chart.n_series):
            values_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
            categories_source = series.category_data if series.category_data else chart.category_data
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name) if categories_source else None

            # Prefer defined name references over raw formulas.
            if series_name_map and i < len(series_name_map):
                cat_ref = series_name_map[i].get('cat') or categories_formula
                val_ref = series_name_map[i].get('val') or values_formula
            else:
                cat_ref = categories_formula
                val_ref = values_formula

            content += f'        <cx:data id="{i}">\n'
            if cat_ref:
                content += '            <cx:strDim type="cat">\n'
                content += f'                <cx:f>{self._escape_xml(cat_ref)}</cx:f>\n'
                content += '            </cx:strDim>\n'
            content += '            <cx:numDim type="size">\n'
            content += f'                <cx:f>{self._escape_xml(val_ref)}</cx:f>\n'
            content += '            </cx:numDim>\n'
            content += '        </cx:data>\n'
        content += '    </cx:chartData>\n'
        content += '    <cx:chart>\n'
        # Conservative compatibility mode for generated treemap chartEx:
        # omit title node entirely to match known-good programmatic output.
        content += '        <cx:plotArea>\n'
        content += '            <cx:plotAreaRegion>\n'
        for i, series in enumerate(chart.n_series):
            hidden_attr = ' hidden="1"' if series.hidden else ''
            series_unique_id = '{' + str(uuid.uuid4()).upper() + '}'
            content += f'                <cx:series layoutId="treemap" uniqueId="{series_unique_id}"{hidden_attr}>\n'
            tx_ref = None
            if series_name_map and i < len(series_name_map):
                tx_ref = series_name_map[i].get('tx')
            # Preserve literal series names for multi-series treemap charts so
            # names round-trip even when no formula-backed tx reference exists.
            if series.name and (tx_ref or chart.n_series.count > 1):
                content += '                    <cx:tx><cx:txData>'
                if tx_ref:
                    content += f'<cx:f>{self._escape_xml(tx_ref)}</cx:f>'
                content += f'<cx:v>{self._escape_xml(series.name)}</cx:v>'
                content += '</cx:txData></cx:tx>\n'
            content += f'                    <cx:dataId val="{i}" />\n'
            content += '                    <cx:layoutPr/>\n'
            content += '                </cx:series>\n'
        content += '            </cx:plotAreaRegion>\n'
        content += '        </cx:plotArea>\n'
        if chart.show_legend:
            content += f'        <cx:legend pos="{self._escape_xml(chart.legend_position)}" align="ctr" overlay="0" />\n'
        content += '    </cx:chart>\n'
        content += '</cx:chartSpace>\n'
        return content

    # -----------------------------------------------------------------------
    # Series XML helpers
    # -----------------------------------------------------------------------

    def _format_chart_title_xml(self, title):
        """Formats chart title XML with rich text payload."""
        escaped_title = self._escape_xml(title)
        content = '        <c:title>\n'
        content += '            <c:tx>\n'
        content += '                <c:rich>\n'
        content += '                    <a:bodyPr/>\n'
        content += '                    <a:lstStyle/>\n'
        content += '                    <a:p>\n'
        content += '                        <a:r>\n'
        content += '                            <a:rPr lang="en-US"/>\n'
        content += '                            <a:t>'
        content += f'{escaped_title}'
        content += '</a:t>\n'
        content += '                        </a:r>\n'
        content += '                        <a:endParaRPr lang="en-US"/>\n'
        content += '                    </a:p>\n'
        content += '                </c:rich>\n'
        content += '            </c:tx>\n'
        content += '            <c:layout/>\n'
        content += '        </c:title>\n'
        return content

    def _format_line_series_xml(self, worksheet, chart, series, series_index, values_formula_override=None, order_override=None):
        """Formats a bar/line/area chart series XML."""
        values_formula = values_formula_override or self._normalize_chart_range_formula(series.values, worksheet.name)
        categories_source = series.category_data if series.category_data else chart.category_data
        categories_formula = None
        if categories_source:
            categories_formula = self._normalize_chart_range_formula(categories_source, worksheet.name)

        order = order_override if order_override is not None else (
            series._series_order if series._series_order is not None else series_index
        )

        content = '                <c:ser>\n'
        content += f'                    <c:idx val="{series_index}"/>\n'
        content += f'                    <c:order val="{order}"/>\n'

        if series.name:
            content += '                    <c:tx>\n'
            content += f'                        <c:v>{self._escape_xml(series.name)}</c:v>\n'
            content += '                    </c:tx>\n'

        if categories_formula:
            content += '                    <c:cat>\n'
            content += '                        <c:strRef>\n'
            content += f'                            <c:f>{self._escape_xml(categories_formula)}</c:f>\n'
            content += '                        </c:strRef>\n'
            content += '                    </c:cat>\n'

        content += '                    <c:val>\n'
        content += '                        <c:numRef>\n'
        content += f'                            <c:f>{self._escape_xml(values_formula)}</c:f>\n'
        content += '                        </c:numRef>\n'
        content += '                    </c:val>\n'

        for eb in series.error_bars:
            content += self._format_error_bars_xml(worksheet, eb)

        content += '                </c:ser>\n'
        return content

    def _format_scatter_series_xml(self, worksheet, chart, series, series_index):
        """Formats a scatter chart series XML using xVal/yVal."""
        y_formula = self._normalize_chart_range_formula(series.values, worksheet.name)
        x_source = series.x_values if series.x_values else series.category_data
        x_formula = self._normalize_chart_range_formula(x_source, worksheet.name) if x_source else None

        order = series._series_order if series._series_order is not None else series_index

        content = '                <c:ser>\n'
        content += f'                    <c:idx val="{series_index}"/>\n'
        content += f'                    <c:order val="{order}"/>\n'

        if series.name:
            content += '                    <c:tx>\n'
            content += f'                        <c:v>{self._escape_xml(series.name)}</c:v>\n'
            content += '                    </c:tx>\n'

        if x_formula:
            content += '                    <c:xVal>\n'
            content += '                        <c:numRef>\n'
            content += f'                            <c:f>{self._escape_xml(x_formula)}</c:f>\n'
            content += '                        </c:numRef>\n'
            content += '                    </c:xVal>\n'

        content += '                    <c:yVal>\n'
        content += '                        <c:numRef>\n'
        content += f'                            <c:f>{self._escape_xml(y_formula)}</c:f>\n'
        content += '                        </c:numRef>\n'
        content += '                    </c:yVal>\n'

        # Per-series smooth
        ser_smooth = getattr(series, '_smooth', False)
        content += f'                    <c:smooth val="{1 if ser_smooth else 0}"/>\n'

        for eb in series.error_bars:
            content += self._format_error_bars_xml(worksheet, eb)

        content += '                </c:ser>\n'
        return content

    def _format_error_bars_xml(self, worksheet, eb):
        """Formats a <c:errBars> element from a ChartErrorBars object."""
        content = '                    <c:errBars>\n'
        content += f'                        <c:errDir val="{self._escape_xml(eb.direction)}"/>\n'
        content += f'                        <c:errBarType val="{self._escape_xml(eb.bar_type)}"/>\n'
        content += f'                        <c:errValType val="{self._escape_xml(eb.val_type)}"/>\n'
        content += f'                        <c:noEndCap val="{1 if eb.no_end_cap else 0}"/>\n'

        if eb.val_type == 'cust':
            if eb.plus_formula:
                plus_formula = self._normalize_chart_range_formula(eb.plus_formula, worksheet.name)
                content += '                        <c:plus>\n'
                content += '                            <c:numRef>\n'
                content += f'                                <c:f>{self._escape_xml(plus_formula)}</c:f>\n'
                content += '                            </c:numRef>\n'
                content += '                        </c:plus>\n'
            if eb.minus_formula:
                minus_formula = self._normalize_chart_range_formula(eb.minus_formula, worksheet.name)
                content += '                        <c:minus>\n'
                content += '                            <c:numRef>\n'
                content += f'                                <c:f>{self._escape_xml(minus_formula)}</c:f>\n'
                content += '                            </c:numRef>\n'
                content += '                        </c:minus>\n'
        else:
            content += f'                        <c:val val="{eb.val}"/>\n'

        if eb.line_width is not None or eb.line_color is not None:
            content += '                        <c:spPr>\n'
            w_attr = f' w="{eb.line_width}"' if eb.line_width is not None else ''
            content += f'                            <a:ln{w_attr}>\n'
            if eb.line_color:
                content += f'                                <a:solidFill><a:srgbClr val="{self._escape_xml(eb.line_color)}"/></a:solidFill>\n'
            content += '                            </a:ln>\n'
            content += '                        </c:spPr>\n'

        content += '                    </c:errBars>\n'
        return content

    # -----------------------------------------------------------------------
    # Axis XML helpers
    # -----------------------------------------------------------------------

    def _format_axis_xml(self, ax, cross_ax_override=None):
        """Formats a ChartAxis object into XML."""
        if ax.axis_type == 'cat':
            tag = 'catAx'
        elif ax.axis_type == 'ser':
            tag = 'serAx'
        else:
            tag = 'valAx'

        content = f'            <c:{tag}>\n'
        content += f'                <c:axId val="{ax.axis_id}"/>\n'
        content += f'                <c:scaling><c:orientation val="{self._escape_xml(ax.orientation)}"/>'
        if ax.min_val is not None:
            content += f'<c:min val="{ax.min_val}"/>'
        if ax.max_val is not None:
            content += f'<c:max val="{ax.max_val}"/>'
        content += '</c:scaling>\n'
        content += f'                <c:delete val="{1 if ax.deleted else 0}"/>\n'
        content += f'                <c:axPos val="{self._escape_xml(ax.position)}"/>\n'
        src_linked = '1' if ax.source_linked else '0'
        content += f'                <c:numFmt formatCode="{self._escape_xml(ax.num_fmt)}" sourceLinked="{src_linked}"/>\n'
        content += f'                <c:majorTickMark val="{self._escape_xml(ax.major_tick_mark)}"/>\n'
        content += f'                <c:minorTickMark val="{self._escape_xml(ax.minor_tick_mark)}"/>\n'
        content += f'                <c:tickLblPos val="{self._escape_xml(ax.tick_lbl_pos)}"/>\n'
        cross_ax = ax.cross_ax
        if cross_ax_override is not None and (cross_ax in (None, 0, "0", "")):
            cross_ax = cross_ax_override
        content += f'                <c:crossAx val="{cross_ax}"/>\n'
        content += f'                <c:crosses val="{self._escape_xml(ax.crosses)}"/>\n'
        if ax.axis_type == 'cat':
            content += f'                <c:auto val="{1 if ax.auto else 0}"/>\n'
            content += f'                <c:lblAlgn val="{self._escape_xml(ax.lbl_algn)}"/>\n'
            content += f'                <c:lblOffset val="{ax.lbl_offset}"/>\n'
        elif ax.axis_type == 'val':
            content += f'                <c:crossBetween val="{self._escape_xml(ax.cross_between)}"/>\n'
        content += f'            </c:{tag}>\n'
        return content

    def _build_combo_axis_cross_map(self, chart):
        """Builds axis_id -> cross_axis_id mapping from combo sub-chart axis pairs."""
        pairs = []
        for sub in getattr(chart, "_sub_charts", []) or []:
            ids = sub.get("ax_ids", [])
            if len(ids) >= 2:
                pair = (int(ids[0]), int(ids[1]))
                if pair not in pairs:
                    pairs.append(pair)

        if not pairs:
            return {}

        mapping = {}
        for cat_id, val_id in pairs:
            if cat_id not in mapping:
                mapping[cat_id] = val_id
            mapping[val_id] = cat_id
        return mapping

    def _format_view_3d_xml(self, chart):
        """Formats chart-level 3D view XML."""
        content = '        <c:view3D>\n'
        content += f'            <c:rotX val="{int(chart.view_3d.rotation_x)}"/>\n'
        content += f'            <c:rotY val="{int(chart.view_3d.rotation_y)}"/>\n'
        content += f'            <c:rAngAx val="{1 if chart.view_3d.right_angle_axes else 0}"/>\n'
        content += f'            <c:perspective val="{int(chart.view_3d.perspective)}"/>\n'
        content += f'            <c:hPercent val="{int(chart.view_3d.height_percent)}"/>\n'
        content += f'            <c:depthPercent val="{int(chart.view_3d.depth_percent)}"/>\n'
        content += '        </c:view3D>\n'
        return content

    def _format_legend_xml(self, chart):
        """Formats chart legend XML from chart settings."""
        content = '        <c:legend>\n'
        content += f'            <c:legendPos val="{chart.legend_position}"/>\n'
        content += '            <c:layout/>\n'
        if not chart.show_legend:
            content += '            <c:delete val="1"/>\n'
        content += '        </c:legend>\n'
        return content

    def _format_cat_axis_xml(self, cat_ax_id, val_ax_id, is_3d=False):
        """Formats category axis XML."""
        content = '            <c:catAx>\n'
        content += f'                <c:axId val="{cat_ax_id}"/>\n'
        content += '                <c:scaling><c:orientation val="minMax"/></c:scaling>\n'
        content += '                <c:delete val="0"/>\n'
        content += '                <c:axPos val="b"/>\n'
        content += '                <c:majorTickMark val="none"/>\n'
        content += '                <c:minorTickMark val="none"/>\n'
        content += '                <c:tickLblPos val="nextTo"/>\n'
        content += f'                <c:crossAx val="{val_ax_id}"/>\n'
        content += '                <c:crosses val="autoZero"/>\n'
        content += '                <c:auto val="1"/>\n'
        content += '                <c:lblAlgn val="ctr"/>\n'
        content += '                <c:lblOffset val="100"/>\n'
        content += '            </c:catAx>\n'
        return content

    def _format_val_axis_xml(self, val_ax_id, cat_ax_id):
        """Formats value axis XML."""
        content = '            <c:valAx>\n'
        content += f'                <c:axId val="{val_ax_id}"/>\n'
        content += '                <c:scaling><c:orientation val="minMax"/></c:scaling>\n'
        content += '                <c:delete val="0"/>\n'
        content += '                <c:axPos val="l"/>\n'
        content += '                <c:majorGridlines/>\n'
        content += '                <c:numFmt formatCode="General" sourceLinked="1"/>\n'
        content += '                <c:majorTickMark val="none"/>\n'
        content += '                <c:minorTickMark val="none"/>\n'
        content += '                <c:tickLblPos val="nextTo"/>\n'
        content += f'                <c:crossAx val="{cat_ax_id}"/>\n'
        content += '                <c:crosses val="autoZero"/>\n'
        content += '                <c:crossBetween val="between"/>\n'
        content += '            </c:valAx>\n'
        return content

    def _format_ser_axis_xml(self, ser_ax_id, val_ax_id):
        """Formats series axis XML for 3D charts."""
        content = '            <c:serAx>\n'
        content += f'                <c:axId val="{ser_ax_id}"/>\n'
        content += '                <c:scaling><c:orientation val="minMax"/></c:scaling>\n'
        content += '                <c:delete val="0"/>\n'
        content += '                <c:axPos val="b"/>\n'
        content += '                <c:majorTickMark val="none"/>\n'
        content += '                <c:minorTickMark val="none"/>\n'
        content += '                <c:tickLblPos val="nextTo"/>\n'
        content += f'                <c:crossAx val="{val_ax_id}"/>\n'
        content += '                <c:crosses val="autoZero"/>\n'
        content += '            </c:serAx>\n'
        return content

    def _normalize_chart_range_formula(self, range_ref, default_sheet_name):
        """
        Normalizes chart formulas into Sheet!$A$1:$B$2 format.
        """
        text = str(range_ref).strip()
        if not text:
            raise ValueError("Chart range cannot be empty")

        if '!' in text:
            sheet_part, ref_part = text.rsplit('!', 1)
            sheet_name = sheet_part
        else:
            escaped_sheet = default_sheet_name.replace("'", "''")
            sheet_name = f"'{escaped_sheet}'"
            ref_part = text

        ref_part = ref_part.strip()
        if ':' in ref_part:
            start_ref, end_ref = ref_part.split(':', 1)
            abs_ref = f"{self._to_abs_a1(start_ref)}:{self._to_abs_a1(end_ref)}"
        else:
            abs_ref = self._to_abs_a1(ref_part)
        return f"{sheet_name}!{abs_ref}"

    def _expand_surface_values_formulas(self, range_ref, default_sheet_name):
        """
        Surface charts require each series value reference to be a single row/column.
        If a rectangular matrix is provided, split it into one-column series refs.
        """
        formula = self._normalize_chart_range_formula(range_ref, default_sheet_name)
        if '!' not in formula:
            return [formula]

        sheet_part, ref_part = formula.rsplit('!', 1)
        ref_text = ref_part.strip()
        if ':' not in ref_text:
            return [formula]

        start_ref, end_ref = ref_text.split(':', 1)
        start_cell = self._parse_abs_a1(start_ref)
        end_cell = self._parse_abs_a1(end_ref)
        if start_cell is None or end_cell is None:
            return [formula]

        r1, c1 = start_cell
        r2, c2 = end_cell
        min_row, max_row = sorted((r1, r2))
        min_col, max_col = sorted((c1, c2))

        # Already a single row/column reference.
        if min_row == max_row or min_col == max_col:
            return [formula]

        refs = []
        for col in range(min_col, max_col + 1):
            col_letter = self._column_letter_from_index(col)
            refs.append(f"{sheet_part}!${col_letter}${min_row}:${col_letter}${max_row}")
        return refs

    def _parse_abs_a1(self, ref):
        """Parses an A1 cell ref into (row, col), both 1-based."""
        match = re.match(r'^\$?([A-Za-z]+)\$?(\d+)$', str(ref).strip())
        if not match:
            return None
        col_letters = match.group(1).upper()
        row = int(match.group(2))
        col = 0
        for ch in col_letters:
            col = (col * 26) + (ord(ch) - ord('A') + 1)
        return (row, col)

    def _column_letter_from_index(self, column_index):
        """Converts a 1-based column index to letters."""
        if column_index < 1:
            return "A"
        letters = ""
        n = int(column_index)
        while n > 0:
            n -= 1
            letters = chr(ord('A') + (n % 26)) + letters
            n //= 26
        return letters

    def _to_abs_a1(self, ref):
        """Converts A1 cell reference to absolute form."""
        match = re.match(r'^\$?([A-Za-z]+)\$?(\d+)$', str(ref).strip())
        if not match:
            return str(ref).strip()
        col = match.group(1).upper()
        row = match.group(2)
        return f"${col}${row}"
