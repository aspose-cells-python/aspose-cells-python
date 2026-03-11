"""
Aspose.Cells for Python - Table XML Loader

Parses xl/tables/tableN.xml parts referenced via <tableParts> in worksheet XML.
Follows the same architecture as PictureXmlLoader / ShapeXmlLoader.
"""

import xml.etree.ElementTree as ET

from .table import Table, TableColumn, TableStyleInfo

_NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'

_NS = {'main': _NS_MAIN}


class TableXmlLoader:
    """Loads table definitions from an XLSX ZIP archive into a worksheet."""

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def load_tables(self, worksheet, worksheet_root, zipf, sheet_num):
        """
        Find <tableParts> in the worksheet XML, resolve each table part path
        via the sheet rels file, parse xl/tables/tableN.xml, and add Table
        objects to worksheet.tables.

        Args:
            worksheet:      Worksheet object to populate.
            worksheet_root: Parsed ElementTree root of the sheet XML.
            zipf:           Open ZipFile of the XLSX archive.
            sheet_num:      1-based sheet index.
        """
        # Find <tableParts>
        table_parts_elem = worksheet_root.find('main:tableParts', _NS)
        if table_parts_elem is None:
            return

        # Load the sheet rels to resolve r:id → target
        rels = self._load_sheet_rels(zipf, sheet_num)

        r_attr = f'{{{_NS_R}}}id'

        for table_part_elem in table_parts_elem:
            r_id = table_part_elem.get(r_attr)
            if not r_id:
                continue
            target = rels.get(r_id, '')
            if not target:
                continue

            # Resolve target relative to xl/worksheets/
            table_path = self._resolve_path('xl/worksheets/', target)

            try:
                table_bytes = zipf.read(table_path)
            except KeyError:
                continue

            table = self._parse_table_xml(table_bytes)
            if table is None:
                continue

            table._source_table_xml = table_bytes
            table._source_part_path = table_path
            worksheet.tables._add_loaded(table)

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _load_sheet_rels(self, zipf, sheet_num) -> dict:
        """Return {rel_id: target_str} from xl/worksheets/_rels/sheetN.xml.rels."""
        rels_path = f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels'
        try:
            rels_bytes = zipf.read(rels_path)
        except KeyError:
            return {}

        rels_root = ET.fromstring(rels_bytes)
        ns = {'r': _NS_RELS}
        result = {}
        for rel in rels_root.findall('r:Relationship', ns):
            rel_id = rel.get('Id', '')
            target = rel.get('Target', '')
            result[rel_id] = target
        return result

    @staticmethod
    def _resolve_path(base: str, target: str) -> str:
        """
        Resolve a relative target path against a base directory.
        base   = 'xl/worksheets/'
        target = '../tables/table1.xml'
        result = 'xl/tables/table1.xml'
        """
        if target.startswith('/'):
            return target.lstrip('/')
        parts = base.rstrip('/').split('/')
        for seg in target.split('/'):
            if seg == '..':
                if parts:
                    parts.pop()
            elif seg and seg != '.':
                parts.append(seg)
        return '/'.join(parts)

    def _parse_table_xml(self, xml_bytes: bytes):
        """
        Parse a table XML part and return a Table object, or None on error.
        """
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return None

        # Root element is <table> in main namespace
        tag = root.tag
        if tag == f'{{{_NS_MAIN}}}table':
            table_elem = root
        else:
            table_elem = root.find(f'{{{_NS_MAIN}}}table')
        if table_elem is None:
            table_elem = root  # fallback

        name        = table_elem.get('name', 'Table1')
        display     = table_elem.get('displayName', name)
        ref         = table_elem.get('ref', 'A1:A1')
        table_id    = int(table_elem.get('id', '1'))
        totals_str  = table_elem.get('totalsRowShown', '0')
        header_str  = table_elem.get('headerRowCount', '1')

        has_headers   = (header_str != '0')
        show_totals   = (totals_str == '1')

        table = Table(name, ref, has_headers)
        table.display_name = display
        table.show_totals_row = show_totals
        table._id = table_id

        # autoFilter presence
        af_elem = table_elem.find(f'{{{_NS_MAIN}}}autoFilter')
        table.show_auto_filter = (af_elem is not None)

        # tableColumns
        cols_elem = table_elem.find(f'{{{_NS_MAIN}}}tableColumns')
        if cols_elem is not None:
            for col_elem in cols_elem.findall(f'{{{_NS_MAIN}}}tableColumn'):
                col_id    = int(col_elem.get('id', '1'))
                col_name  = col_elem.get('name', f'Column{col_id}')
                tc = TableColumn(col_id, col_name)
                tc.totals_row_function = col_elem.get('totalsRowFunction', 'none')
                tc.totals_row_label    = col_elem.get('totalsRowLabel', '')
                tc.totals_row_formula  = col_elem.get('totalsRowFormula', '')
                table.columns.append(tc)

        # tableStyleInfo
        si_elem = table_elem.find(f'{{{_NS_MAIN}}}tableStyleInfo')
        if si_elem is not None:
            si = TableStyleInfo()
            si.name                = si_elem.get('name', si.name)
            si.show_first_column   = si_elem.get('showFirstColumn', '0') == '1'
            si.show_last_column    = si_elem.get('showLastColumn', '0') == '1'
            si.show_row_stripes    = si_elem.get('showRowStripes', '1') == '1'
            si.show_column_stripes = si_elem.get('showColumnStripes', '0') == '1'
            table.table_style_info = si

        return table
