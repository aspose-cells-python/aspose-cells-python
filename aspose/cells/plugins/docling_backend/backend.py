"""
Docling backend using Aspose.Cells for Excel processing.
"""

import logging

from io import BytesIO
from pathlib import Path
from typing import Any, Union, cast

from docling_core.types.doc import (
    BoundingBox,
    CoordOrigin,
    DocItem,
    DoclingDocument,
    DocumentOrigin,
    GroupLabel,
    ImageRef,
    ProvenanceItem,
    Size,
    TableCell,
    TableData,
)


class AsposeCellsDoclingDocument(DoclingDocument):
    """Extended DoclingDocument that uses Aspose.Cells MarkdownConverter for export."""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._aspose_markdown_content = None
    
    def export_to_markdown(self, **kwargs) -> str:
        """Export using Aspose MarkdownConverter if available, fallback to docling default."""
        if hasattr(self, '_aspose_markdown_content') and self._aspose_markdown_content:
            return self._aspose_markdown_content
        else:
            # Fallback to original docling export
            return super().export_to_markdown(**kwargs)


from PIL import Image as PILImage
from pydantic import BaseModel, NonNegativeInt, PositiveInt
from typing_extensions import override

from docling.backend.abstract_backend import (
    DeclarativeDocumentBackend,
    PaginatedDocumentBackend,
)
from docling.datamodel.base_models import InputFormat
from docling.datamodel.document import InputDocument

# Import our Aspose.Cells modules
from aspose.cells import Workbook, Worksheet

_log = logging.getLogger(__name__)


class ExcelCell(BaseModel):
    """Represents an Excel cell.

    Attributes:
        row: The row number of the cell.
        col: The column number of the cell.
        text: The text content of the cell.
        row_span: The number of rows the cell spans.
        col_span: The number of columns the cell spans.
    """

    row: int
    col: int
    text: str
    row_span: int
    col_span: int


class ExcelTable(BaseModel):
    """Represents an Excel table on a worksheet.

    Attributes:
        anchor: The column and row indices of the upper-left cell of the table
        (0-based index).
        num_rows: The number of rows in the table.
        num_cols: The number of columns in the table.
        data: The data in the table, represented as a list of ExcelCell objects.
    """

    anchor: tuple[NonNegativeInt, NonNegativeInt]
    num_rows: int
    num_cols: int
    data: list[ExcelCell]


class CellsDocumentBackend(DeclarativeDocumentBackend, PaginatedDocumentBackend):
    """Backend for parsing Excel workbooks using Aspose.Cells.

    The backend converts an Excel workbook into a DoclingDocument object.
    Each worksheet is converted into a separate page.
    The following elements are parsed:
    - Cell contents, parsed as tables. If two groups of cells are disconnected
    between each other, they will be parsed as two different tables.
    - Images, parsed as PictureItem objects.

    The DoclingDocument tables and pictures have their provenance information, including
    the position in their original Excel worksheet. The position is represented by a
    bounding box object with the cell indices as units (0-based index). The size of this
    bounding box is the number of columns and rows that the table or picture spans.
    """

    @override
    def __init__(
        self, in_doc: "InputDocument", path_or_stream: Union[BytesIO, Path], **kwargs
    ) -> None:
        """Initialize the CellsDocumentBackend object.

        Parameters:
            in_doc: The input document object.
            path_or_stream: The path or stream to the Excel file.

        Raises:
            RuntimeError: An error occurred parsing the file.
        """
        super().__init__(in_doc, path_or_stream)

        # Store conversion parameters
        self.conversion_kwargs = kwargs

        # Initialise the parents for the hierarchy
        self.max_levels = 10

        self.parents: dict[int, Any] = {}
        for i in range(-1, self.max_levels):
            self.parents[i] = None

        self.workbook = None
        try:
            if isinstance(self.path_or_stream, BytesIO):
                # For BytesIO, we need to write to a temporary file
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                    tmp.write(self.path_or_stream.getvalue())
                    tmp_path = tmp.name
                self.workbook = Workbook.load(tmp_path)
                import os
                os.unlink(tmp_path)  # Clean up temp file

            elif isinstance(self.path_or_stream, Path):
                self.workbook = Workbook.load(str(self.path_or_stream))

            self.valid = self.workbook is not None
        except Exception as e:
            self.valid = False
            raise RuntimeError(
                f"CellsDocumentBackend could not load document with hash {self.document_hash}"
            ) from e

    @override
    def is_valid(self) -> bool:
        _log.debug(f"valid: {self.valid}")
        return self.valid

    @classmethod
    @override
    def supports_pagination(cls) -> bool:
        return True

    @override
    def page_count(self) -> int:
        if self.is_valid() and self.workbook:
            return len(self.workbook.sheetnames)
        else:
            return 0

    @classmethod
    @override
    def supported_formats(cls) -> set[InputFormat]:
        return {InputFormat.XLSX}

    @override
    def convert(self, **kwargs) -> DoclingDocument:
        """Parse the Excel workbook into a DoclingDocument object.

        Raises:
            RuntimeError: Unable to run the conversion since the backend object failed to
            initialize.

        Returns:
            The DoclingDocument object representing the Excel workbook.
        """
        origin = DocumentOrigin(
            filename=self.file.name or "file.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            binary_hash=self.document_hash,
        )

        doc = AsposeCellsDoclingDocument(name=self.file.stem or "file.xlsx", origin=origin)

        if self.is_valid():
            doc = self._convert_workbook_with_markdown(doc)
        else:
            raise RuntimeError(
                f"Cannot convert doc with {self.document_hash} because the backend failed to init."
            )

        return doc

    def _convert_workbook_with_markdown(self, doc: AsposeCellsDoclingDocument) -> AsposeCellsDoclingDocument:
        """Convert workbook using our MarkdownConverter and embed result in DoclingDocument."""
        
        # Use our MarkdownConverter instead of custom docling logic
        from ...converters.markdown_converter import MarkdownConverter
        
        converter = MarkdownConverter()
        
        # Use same parameters as markitdown plugin
        convert_kwargs = {
            "sheet_name": self.conversion_kwargs.get("sheet_name", None),
            "include_metadata": self.conversion_kwargs.get("include_metadata", True),
            "value_mode": self.conversion_kwargs.get("value_mode", "value"),
            "include_hyperlinks": self.conversion_kwargs.get("include_hyperlinks", True),
        }
        
        # Convert workbook to markdown using our converter
        markdown_content = converter.convert_workbook(self.workbook, **convert_kwargs)
        
        # Store the markdown content in the document for export
        doc._aspose_markdown_content = markdown_content
        
        # Still add basic docling structure for compatibility
        doc = self._convert_workbook(doc)
        
        return doc

    def _convert_workbook(self, doc: DoclingDocument) -> DoclingDocument:
        """Parse the Excel workbook and attach its structure to a DoclingDocument.

        Args:
            doc: A DoclingDocument object.

        Returns:
            A DoclingDocument object with the parsed items.
        """

        if self.workbook is not None:
            # Iterate over all sheets
            for i, sheet_name in enumerate(self.workbook.sheetnames):
                _log.info(f"Processing sheet: {sheet_name}")

                sheet = self.workbook.worksheets[sheet_name]
                page_no = i + 1
                # Add page with initial size
                page = doc.add_page(page_no=page_no, size=Size(width=0, height=0))

                self.parents[0] = doc.add_group(
                    parent=None,
                    label=GroupLabel.SECTION,
                    name=f"sheet: {sheet_name}",
                )
                doc = self._convert_sheet(doc, sheet, page_no)
                width, height = self._find_page_size(doc, page_no)
                page.size = Size(width=width, height=height)
        else:
            _log.error("Workbook is not initialized.")

        return doc

    def _convert_sheet(self, doc: DoclingDocument, sheet: Worksheet, page_no: int) -> DoclingDocument:
        """Parse an Excel worksheet and attach its structure to a DoclingDocument

        Args:
            doc: The DoclingDocument to be updated.
            sheet: The Excel worksheet to be parsed.
            page_no: The page number for this sheet.

        Returns:
            The updated DoclingDocument.
        """

        doc = self._find_tables_in_sheet(doc, sheet, page_no)
        doc = self._find_images_in_sheet(doc, sheet, page_no)

        return doc

    def _find_tables_in_sheet(
        self, doc: DoclingDocument, sheet: Worksheet, page_no: int
    ) -> DoclingDocument:
        """Find all tables in an Excel sheet and attach them to a DoclingDocument.

        Args:
            doc: The DoclingDocument to be updated.
            sheet: The Excel worksheet to be parsed.
            page_no: The page number for this sheet.

        Returns:
            The updated DoclingDocument.
        """

        if self.workbook is not None:
            tables = self._find_data_tables(sheet)

            for excel_table in tables:
                origin_col = excel_table.anchor[0]
                origin_row = excel_table.anchor[1]
                num_rows = excel_table.num_rows
                num_cols = excel_table.num_cols

                table_data = TableData(
                    num_rows=num_rows,
                    num_cols=num_cols,
                    table_cells=[],
                )

                for excel_cell in excel_table.data:
                    cell = TableCell(
                        text=excel_cell.text,
                        row_span=excel_cell.row_span,
                        col_span=excel_cell.col_span,
                        start_row_offset_idx=excel_cell.row,
                        end_row_offset_idx=excel_cell.row + excel_cell.row_span,
                        start_col_offset_idx=excel_cell.col,
                        end_col_offset_idx=excel_cell.col + excel_cell.col_span,
                        column_header=excel_cell.row == 0,
                        row_header=False,
                    )
                    table_data.table_cells.append(cell)

                doc.add_table(
                    data=table_data,
                    parent=self.parents[0],
                    prov=ProvenanceItem(
                        page_no=page_no,
                        charspan=(0, 0),
                        bbox=BoundingBox.from_tuple(
                            (
                                origin_col,
                                origin_row,
                                origin_col + num_cols,
                                origin_row + num_rows,
                            ),
                            origin=CoordOrigin.TOPLEFT,
                        ),
                    ),
                )

        return doc

    def _find_data_tables(self, sheet: Worksheet) -> list[ExcelTable]:
        """Find all compact rectangular data tables in an Excel worksheet.

        Args:
            sheet: The Excel worksheet to be parsed.

        Returns:
            A list of ExcelTable objects representing the data tables.
        """
        tables: list[ExcelTable] = []
        visited: set[tuple[int, int]] = set()

        # Get all non-empty cells
        non_empty_cells = []
        for row in range(1, 1000):  # Reasonable limit
            for col in range(1, 100):   # Reasonable limit
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    non_empty_cells.append((row-1, col-1))  # Convert to 0-based

        # Group adjacent cells into tables
        for row, col in non_empty_cells:
            if (row, col) in visited:
                continue

            # Find table bounds starting from this cell
            table_bounds, visited_cells = self._find_table_bounds(sheet, row, col)
            visited.update(visited_cells)
            tables.append(table_bounds)

        return tables

    def _find_table_bounds(
        self,
        sheet: Worksheet,
        start_row: int,
        start_col: int,
    ) -> tuple[ExcelTable, set[tuple[int, int]]]:
        """Determine the bounds of a compact rectangular table.

        Args:
            sheet: The Excel worksheet to be parsed.
            start_row: The row number of the starting cell (0-based).
            start_col: The column number of the starting cell (0-based).

        Returns:
            A tuple with an Excel table and a set of cell coordinates.
        """
        _log.debug("find_table_bounds")

        max_row = self._find_table_bottom(sheet, start_row, start_col)
        max_col = self._find_table_right(sheet, start_row, start_col)

        # Collect the data within the bounds
        data = []
        visited_cells: set[tuple[int, int]] = set()
        
        for row in range(start_row, max_row + 1):
            for col in range(start_col, max_col + 1):
                # Convert to 1-based for our cell access
                cell = sheet.cell(row + 1, col + 1)
                
                # Check for merged cells (simplified - assume no merging for now)
                row_span = 1
                col_span = 1

                if (row, col) not in visited_cells:
                    cell_value = cell.value if cell.value is not None else ""
                    data.append(
                        ExcelCell(
                            row=row - start_row,
                            col=col - start_col,
                            text=str(cell_value),
                            row_span=row_span,
                            col_span=col_span,
                        )
                    )

                    # Mark cells in span as visited
                    for span_row in range(row, row + row_span):
                        for span_col in range(col, col + col_span):
                            visited_cells.add((span_row, span_col))

        return (
            ExcelTable(
                anchor=(start_col, start_row),
                num_rows=max_row + 1 - start_row,
                num_cols=max_col + 1 - start_col,
                data=data,
            ),
            visited_cells,
        )

    def _find_table_bottom(
        self, sheet: Worksheet, start_row: int, start_col: int
    ) -> int:
        """Find the bottom boundary of a table."""
        max_row = start_row

        for row in range(start_row + 1, 1000):  # Reasonable limit
            cell = sheet.cell(row + 1, start_col + 1)  # Convert to 1-based
            if cell.value is None or not str(cell.value).strip():
                break
            max_row = row

        return max_row

    def _find_table_right(
        self, sheet: Worksheet, start_row: int, start_col: int
    ) -> int:
        """Find the right boundary of a table."""
        max_col = start_col

        for col in range(start_col + 1, 100):  # Reasonable limit
            cell = sheet.cell(start_row + 1, col + 1)  # Convert to 1-based
            if cell.value is None or not str(cell.value).strip():
                break
            max_col = col

        return max_col

    def _find_images_in_sheet(
        self, doc: DoclingDocument, sheet: Worksheet, page_no: int
    ) -> DoclingDocument:
        """Find images in the Excel sheet and attach them to the DoclingDocument.

        Args:
            doc: The DoclingDocument to be updated.
            sheet: The Excel worksheet to be parsed.
            page_no: The page number for this sheet.

        Returns:
            The updated DoclingDocument.
        """
        if self.workbook is not None:
            # Check if the sheet has images (simplified implementation)
            if hasattr(sheet, 'images') and sheet.images:
                for image in sheet.images:
                    try:
                        # Convert our Image to PIL Image for compatibility
                        if hasattr(image, 'data') and image.data:
                            pil_image = PILImage.open(BytesIO(image.data))
                            
                            # Get anchor information (simplified)
                            anchor = (0, 0, 5, 5)  # Default anchor
                            if hasattr(image, 'anchor') and image.anchor:
                                anchor = (
                                    getattr(image.anchor, 'col', 0),
                                    getattr(image.anchor, 'row', 0),
                                    getattr(image.anchor, 'col', 0) + 5,
                                    getattr(image.anchor, 'row', 0) + 5,
                                )
                            
                            doc.add_picture(
                                parent=self.parents[0],
                                image=ImageRef.from_pil(image=pil_image, dpi=72),
                                caption=None,
                                prov=ProvenanceItem(
                                    page_no=page_no,
                                    charspan=(0, 0),
                                    bbox=BoundingBox.from_tuple(
                                        anchor, origin=CoordOrigin.TOPLEFT
                                    ),
                                ),
                            )
                    except Exception as e:
                        _log.warning(f"Could not extract image from sheet: {e}")

        return doc

    @staticmethod
    def _find_page_size(
        doc: DoclingDocument, page_no: PositiveInt
    ) -> tuple[float, float]:
        left: float = -1.0
        top: float = -1.0
        right: float = -1.0
        bottom: float = -1.0
        for item, _ in doc.iterate_items(traverse_pictures=True, page_no=page_no):
            if not isinstance(item, DocItem):
                continue
            for provenance in item.prov:
                bbox = provenance.bbox
                left = min(left, bbox.l) if left != -1 else bbox.l
                right = max(right, bbox.r) if right != -1 else bbox.r
                top = min(top, bbox.t) if top != -1 else bbox.t
                bottom = max(bottom, bbox.b) if bottom != -1 else bbox.b

        return (max(right - left, 10.0), max(bottom - top, 10.0))