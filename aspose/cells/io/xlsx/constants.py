"""XLSX format constants and XML templates."""

from typing import Dict


class XlsxConstants:
    """Central place for all XLSX-related constants and XML templates."""
    
    # XML Namespaces
    NAMESPACES = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    
    # Content Types
    CONTENT_TYPES = {
        'defaults': [
            ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
            ("xml", "application/xml")
        ],
        'overrides': {
            'workbook': ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
            'styles': ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
            'theme': ("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
            'core_props': ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
            'app_props': ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
            'shared_strings': ("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
            'worksheet': "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        }
    }
    
    # Relationship Types
    REL_TYPES = {
        'office_document': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        'core_properties': "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        'extended_properties': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        'worksheet': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        'styles': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        'theme': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        'shared_strings': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        'hyperlink': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    }
    
    # Built-in Number Formats
    BUILTIN_NUMBER_FORMATS = {
        'General': 0,
        '0': 1,
        '0.00': 2,
        '#,##0': 3,
        '#,##0.00': 4,
        '0%': 9,
        '0.00%': 10,
        'mm/dd/yyyy': 14,
        '$#,##0': 164,  # Custom format starting from 164
        '$#,##0.00': 165,
        '0.0%': 166
    }
    
    # Default Style Properties
    DEFAULT_FONT = {
        'name': 'Calibri',
        'size': 11,
        'bold': False,
        'italic': False,
        'color': '000000'
    }
    
    DEFAULT_FILLS = [
        {'pattern': 'none', 'color': None},
        {'pattern': 'gray125', 'color': None}
    ]
    
    DEFAULT_BORDER = {
        'left': None, 'right': None, 'top': None, 'bottom': None
    }
    
    DEFAULT_CELL_FORMAT = {
        'font_id': 0,
        'fill_id': 0,
        'border_id': 0,
        'number_format_id': 0
    }
    
    # Color Mapping
    COLOR_MAP = {
        'black': '000000',
        'white': 'FFFFFF',
        'red': 'FF0000',
        'green': '00FF00',
        'blue': '0000FF',
        'yellow': 'FFFF00',
        'cyan': '00FFFF',
        'magenta': 'FF00FF',
        'darkblue': '000080',
        'darkgreen': '008000',
        'darkred': '800000',
        'purple': '800080',
        'orange': 'FFA500',
        'gray': '808080',
        'lightblue': 'ADD8E6',
        'lightgreen': '90EE90',
        'lightyellow': 'FFFFE0',
        'lightgray': 'D3D3D3',
        'lightcyan': 'E0FFFF',
        'lightcoral': 'F08080',
        'lightpink': 'FFB6C1',
        'gold': 'FFD700',
        'lavender': 'E6E6FA'
    }
    
    # Default Sheet Properties
    SHEET_DEFAULTS = {
        'default_row_height': "15",
        'default_col_width': "10",
        'window_x': "0",
        'window_y': "0",
        'window_width': "14980",
        'window_height': "8580"
    }
    
    # Drawing and Image Constants
    DRAWING_NAMESPACES = {
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
        'a14': 'http://schemas.microsoft.com/office/drawing/2010/main'
    }
    
    # Image-related constants
    IMAGE_DEFAULTS = {
        'width': 100,
        'height': 100,
        'name_prefix': 'Image',
        'emu_per_pixel': 9525  # English Metric Units per pixel
    }
    
    # Extension URIs used in images
    IMAGE_EXTENSIONS = {
        'creation_id': '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}',
        'local_dpi': '{28A0092B-C50C-407E-A947-70E740481C1C}'
    }
    
    # Image MIME types
    IMAGE_CONTENT_TYPES = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif'
    }
    
    # XML Attributes and values
    XML_ATTRIBUTES = {
        'edit_as': 'oneCell',
        'no_change_aspect': '1',
        'cstate': 'print',
        'local_dpi_val': '0',
        'transform_xy': '0',
        'preset_geom': 'rect'
    }


class XlsxTemplates:
    """XML templates for XLSX files."""
    
    @staticmethod
    def get_theme_xml() -> str:
        """Get the complete theme XML template."""
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera>
            <a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig>
          </a:scene3d>
          <a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>'''
    
    @staticmethod
    def get_app_properties_data() -> Dict[str, str]:
        """Get application properties data."""
        return {
            'application': 'Aspose.Cells.Python',
            'doc_security': '0',
            'links_up_to_date': 'false',
            'shared_doc': 'false',
            'hyperlinks_changed': 'false',
            'app_version': '16.0300'
        }
    
    @staticmethod
    def get_core_properties_data() -> Dict[str, str]:
        """Get core properties data."""
        return {
            'creator': 'Aspose.Cells.Python',
            'modified': '2024-07-20T14:00:00Z',
            'created': '2024-07-20T14:00:00Z'
        }
    
    @staticmethod
    def get_rels_data():
        """Get main relationships data."""
        return [
            ("rId1", XlsxConstants.REL_TYPES['office_document'], "xl/workbook.xml"),
            ("rId2", XlsxConstants.REL_TYPES['core_properties'], "docProps/core.xml"),
            ("rId3", XlsxConstants.REL_TYPES['extended_properties'], "docProps/app.xml")
        ]