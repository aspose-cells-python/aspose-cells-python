"""
Aspose.Cells for Python - Sparkline XML Loader

Reads <extLst> sparkline data embedded in worksheet XML.
Sparklines use the x14 extension mechanism — they are NOT separate part files.

Extension URI: {05C60535-1F16-4fd2-B633-F4F36F0B64E0}
"""

import xml.etree.ElementTree as ET

from .sparkline import SparklineGroup, SparklineType, SparklineEmptyCells, Sparkline

_SPARKLINE_URI  = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
_NS_MAIN        = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_NS_X14         = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
_NS_XM          = 'http://schemas.microsoft.com/office/excel/2006/main'
_NS_XR2         = 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2'

_TYPE_MAP = {
    'line':     SparklineType.LINE,
    'column':   SparklineType.COLUMN,
    'win-loss': SparklineType.WIN_LOSS,
}
_EMPTY_CELLS_MAP = {
    'zero':      SparklineEmptyCells.ZERO,
    'gap':       SparklineEmptyCells.GAP,
    'connected': SparklineEmptyCells.CONNECTED,
}


class SparklineXmlLoader:
    """Loads sparkline group data from the <extLst> in a worksheet XML root."""

    def load_sparklines(self, worksheet, worksheet_root):
        """
        Find and parse sparkline data in worksheet_root's <extLst>.
        Populates worksheet.sparkline_groups and stores source XML for round-trip.

        Args:
            worksheet:      Worksheet object to populate.
            worksheet_root: Parsed ElementTree root of the sheet XML.
        """
        # Find <extLst> as a direct child of the worksheet root
        ext_lst = worksheet_root.find(f'{{{_NS_MAIN}}}extLst')
        if ext_lst is None:
            # Try without namespace (some files omit it on extLst)
            ext_lst = worksheet_root.find('extLst')
        if ext_lst is None:
            return

        # Find the sparkline ext by URI
        sparkline_ext = None
        for ext in ext_lst:
            uri = ext.get('uri', '')
            if uri == _SPARKLINE_URI:
                sparkline_ext = ext
                break
        if sparkline_ext is None:
            return

        # Store raw XML of the <ext> element for round-trip preservation
        worksheet._source_sparkline_xml = ET.tostring(sparkline_ext, encoding='unicode')

        # Find <x14:sparklineGroups>
        groups_elem = sparkline_ext.find(f'{{{_NS_X14}}}sparklineGroups')
        if groups_elem is None:
            return

        for group_elem in groups_elem.findall(f'{{{_NS_X14}}}sparklineGroup'):
            group = self._parse_group(group_elem)
            worksheet.sparkline_groups._add_loaded(group)

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _parse_group(self, group_elem) -> SparklineGroup:
        group = SparklineGroup()

        # Store raw group XML for individual round-trip (future use)
        group._source_group_xml = ET.tostring(group_elem, encoding='unicode')

        # uid
        uid_attr = f'{{{_NS_XR2}}}uid'
        group._uid = group_elem.get(uid_attr)

        # type
        type_str = group_elem.get('type', 'line').lower()
        group.type = _TYPE_MAP.get(type_str, SparklineType.LINE)

        # displayEmptyCellsAs
        empty_str = group_elem.get('displayEmptyCellsAs', 'gap').lower()
        group.display_empty_cells_as = _EMPTY_CELLS_MAP.get(empty_str, SparklineEmptyCells.GAP)

        # Colors (strip the leading "FF" alpha prefix → 6-char RRGGBB)
        _color_map = {
            'colorSeries':   'color_series',
            'colorNegative': 'color_negative',
            'colorAxis':     'color_axis',
            'colorMarkers':  'color_markers',
            'colorFirst':    'color_first',
            'colorLast':     'color_last',
            'colorHigh':     'color_high',
            'colorLow':      'color_low',
        }
        for xml_name, py_attr in _color_map.items():
            elem = group_elem.find(f'{{{_NS_X14}}}{xml_name}')
            if elem is not None:
                rgb = elem.get('rgb', '')
                if len(rgb) == 8:            # AARRGGBB
                    setattr(group, py_attr, rgb[2:].upper())
                elif len(rgb) == 6:          # RRGGBB (rare)
                    setattr(group, py_attr, rgb.upper())

        # Individual sparklines
        sparklines_elem = group_elem.find(f'{{{_NS_X14}}}sparklines')
        if sparklines_elem is not None:
            for sp_elem in sparklines_elem.findall(f'{{{_NS_X14}}}sparkline'):
                f_elem   = sp_elem.find(f'{{{_NS_XM}}}f')
                ref_elem = sp_elem.find(f'{{{_NS_XM}}}sqref')
                data_range = f_elem.text.strip()   if f_elem   is not None else ''
                cell_ref   = ref_elem.text.strip() if ref_elem is not None else ''
                group.sparklines.append(Sparkline(data_range, cell_ref))

        return group
