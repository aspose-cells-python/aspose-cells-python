"""
Aspose.Cells for Python - Sparkline XML Saver

Generates the <extLst> block containing x14:sparklineGroups for embedding
directly in worksheet XML. No separate parts or rels are needed.
"""

from .sparkline import SparklineType, SparklineEmptyCells

_SPARKLINE_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
_NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
_NS_XM  = 'http://schemas.microsoft.com/office/excel/2006/main'

_TYPE_TO_STR = {
    SparklineType.LINE:     None,          # attribute omitted for default
    SparklineType.COLUMN:   'column',
    SparklineType.WIN_LOSS: 'win-loss',
}
_EMPTY_CELLS_TO_STR = {
    SparklineEmptyCells.ZERO:      'zero',
    SparklineEmptyCells.GAP:       'gap',
    SparklineEmptyCells.CONNECTED: 'connected',
}


class SparklineXmlSaver:
    """Serialises SparklineGroupCollection to <extLst> XML."""

    def __init__(self, escape_xml_fn=None):
        if escape_xml_fn is None:
            def _default_escape(text):
                if not isinstance(text, str):
                    text = str(text)
                text = text.replace('&', '&amp;')
                text = text.replace('<', '&lt;')
                text = text.replace('>', '&gt;')
                text = text.replace('"', '&quot;')
                return text
            escape_xml_fn = _default_escape
        self._escape = escape_xml_fn

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def format_sparkline_extlst_xml(self, worksheet) -> str:
        """
        Return the complete <extLst>...</extLst> XML string for sparklines,
        or '' if the worksheet has no sparklines.

        Round-trip: if _source_sparkline_xml is set and sparklines are not
        dirty, write the source verbatim wrapped in <extLst>.
        """
        groups = getattr(worksheet, 'sparkline_groups', None)
        if groups is None or groups.count == 0:
            return ''

        # Round-trip: use source XML when available and unmodified
        source = getattr(worksheet, '_source_sparkline_xml', None)
        dirty  = getattr(worksheet, '_sparklines_dirty', False)
        if source and not dirty:
            return f'    <extLst>{source}</extLst>\n'

        # Generate fresh XML
        return self._generate_extlst_xml(worksheet)

    # ------------------------------------------------------------------
    # XML generation
    # ------------------------------------------------------------------

    def _generate_extlst_xml(self, worksheet) -> str:
        esc = self._escape
        lines = [
            '    <extLst>',
            f'      <ext uri="{_SPARKLINE_URI}"',
            f'           xmlns:x14="{_NS_X14}">',
            f'        <x14:sparklineGroups xmlns:xm="{_NS_XM}">',
        ]

        for group in worksheet.sparkline_groups:
            lines.extend(self._format_group(group, esc))

        lines += [
            '        </x14:sparklineGroups>',
            '      </ext>',
            '    </extLst>',
        ]
        return '\n'.join(lines) + '\n'

    def _format_group(self, group, esc) -> list:
        # Build opening tag attributes
        attrs = []
        type_str = _TYPE_TO_STR.get(group.type)
        if type_str:
            attrs.append(f'type="{type_str}"')
        empty_str = _EMPTY_CELLS_TO_STR.get(group.display_empty_cells_as, 'gap')
        attrs.append(f'displayEmptyCellsAs="{empty_str}"')
        if group._uid:
            attrs.append(f'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"')
            attrs.append(f'xr2:uid="{esc(group._uid)}"')

        attr_str = ' '.join(attrs)
        lines = [f'          <x14:sparklineGroup {attr_str}>']

        # Colors — always emit all 8 for Excel compatibility
        for xml_name, py_attr in [
            ('colorSeries',   'color_series'),
            ('colorNegative', 'color_negative'),
            ('colorAxis',     'color_axis'),
            ('colorMarkers',  'color_markers'),
            ('colorFirst',    'color_first'),
            ('colorLast',     'color_last'),
            ('colorHigh',     'color_high'),
            ('colorLow',      'color_low'),
        ]:
            color = getattr(group, py_attr, '000000').upper().lstrip('#')
            if len(color) == 6:
                rgb = f'FF{color}'
            elif len(color) == 8:
                rgb = color
            else:
                rgb = f'FF{color:>06}'
            lines.append(f'            <x14:{xml_name} rgb="{rgb}"/>')

        # Sparklines
        lines.append('            <x14:sparklines>')
        for sp in group.sparklines:
            lines.append('              <x14:sparkline>')
            lines.append(f'                <xm:f>{esc(sp.data_range)}</xm:f>')
            lines.append(f'                <xm:sqref>{esc(sp.cell_reference)}</xm:sqref>')
            lines.append('              </x14:sparkline>')
        lines.append('            </x14:sparklines>')

        lines.append('          </x14:sparklineGroup>')
        return lines
