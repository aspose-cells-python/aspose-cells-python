"""
Aspose.Cells for Python - Table XML Saver

Generates xl/tables/tableN.xml content for Excel structured tables.
Follows the same architecture as PictureXmlSaver / ShapeXmlSaver.
"""


class TableXmlSaver:
    """Serialises Table objects to ECMA-376 table XML."""

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
    # Public API
    # ------------------------------------------------------------------

    def format_table_xml(self, table, table_id: int) -> str:
        """
        Return the complete XML string for xl/tables/tableN.xml.

        If table._source_table_xml is set, decode and return it verbatim
        (round-trip preservation).  Otherwise generate fresh XML.

        Args:
            table:    Table object.
            table_id: Globally-unique integer id to use in the <table id=""> attribute.
        Returns:
            UTF-8 XML string (not bytes).
        """
        if getattr(table, '_source_table_xml', None) is not None:
            src = table._source_table_xml
            if isinstance(src, bytes):
                return src.decode('utf-8', errors='replace')
            return src

        return self._generate_table_xml(table, table_id)

    # ------------------------------------------------------------------
    # XML generation
    # ------------------------------------------------------------------

    def _generate_table_xml(self, table, table_id: int) -> str:
        esc = self._escape
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        totals_shown = '1' if table.show_totals_row else '0'
        header_count = '' if table.has_headers else ' headerRowCount="0"'

        name         = esc(table.name)
        display_name = esc(table.display_name or table.name)
        ref          = esc(table.ref)

        lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<table xmlns="{xmlns}"',
            f'       id="{table_id}" name="{name}" displayName="{display_name}"',
            f'       ref="{ref}" totalsRowShown="{totals_shown}"{header_count}>',
        ]

        # autoFilter
        if table.show_auto_filter:
            lines.append(f'  <autoFilter ref="{ref}"/>')

        # tableColumns
        col_count = len(table.columns)
        lines.append(f'  <tableColumns count="{col_count}">')
        for col in table.columns:
            attrs = [
                f'id="{col.id}"',
                f'name="{esc(col.name)}"',
            ]
            fn = (col.totals_row_function or 'none').strip()
            if fn and fn != 'none':
                attrs.append(f'totalsRowFunction="{esc(fn)}"')
            if col.totals_row_label:
                attrs.append(f'totalsRowLabel="{esc(col.totals_row_label)}"')
            lines.append(f'    <tableColumn {" ".join(attrs)}/>')
        lines.append('  </tableColumns>')

        # tableStyleInfo
        si = table.table_style_info
        si_attrs = [
            f'name="{esc(si.name)}"',
            f'showFirstColumn="{1 if si.show_first_column else 0}"',
            f'showLastColumn="{1 if si.show_last_column else 0}"',
            f'showRowStripes="{1 if si.show_row_stripes else 0}"',
            f'showColumnStripes="{1 if si.show_column_stripes else 0}"',
        ]
        lines.append(f'  <tableStyleInfo {" ".join(si_attrs)}/>')

        lines.append('</table>')
        return '\n'.join(lines) + '\n'
