"""
Custom exception classes for Aspose.Cells operations.
"""


class AsposeException(Exception):
    """Base exception for all Aspose.Cells errors."""
    pass


class FileFormatError(AsposeException):
    """Raised when file format is not supported or invalid."""
    pass


class InvalidCoordinateError(AsposeException):
    """Raised when cell coordinate is invalid."""
    pass


class WorksheetNotFoundError(AsposeException):
    """Raised when worksheet cannot be found."""
    pass


class CellValueError(AsposeException):
    """Raised when cell value operation fails."""
    pass


class ExportError(AsposeException):
    """Raised when data export operation fails."""
    pass


class ExcelValidationError(AsposeException):
    """Raised when Excel data validation fails."""
    pass


class CellRangeError(AsposeException):
    """Raised when cell range operation fails."""
    pass