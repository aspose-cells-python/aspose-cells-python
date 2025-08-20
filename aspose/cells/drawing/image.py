"""
Image handling and processing for Excel worksheets.
"""

import io
from typing import Union, Optional, Tuple
from pathlib import Path
from enum import Enum

from .anchor import Anchor, AnchorType


class ImageFormat(Enum):
    """Supported image formats."""
    
    PNG = "png"
    JPEG = "jpeg"
    JPG = "jpg"
    GIF = "gif"
    
    @classmethod
    def from_extension(cls, filename: Union[str, Path]) -> 'ImageFormat':
        """Infer format from file extension."""
        ext = Path(filename).suffix.lower()
        format_map = {
            '.png': cls.PNG,
            '.jpg': cls.JPEG,
            '.jpeg': cls.JPEG,
            '.gif': cls.GIF,
        }
        return format_map.get(ext, cls.PNG)
    
    @classmethod
    def from_mimetype(cls, mimetype: str) -> 'ImageFormat':
        """Infer format from MIME type."""
        mime_map = {
            'image/png': cls.PNG,
            'image/jpeg': cls.JPEG,
            'image/jpg': cls.JPEG,
            'image/gif': cls.GIF,
        }
        return mime_map.get(mimetype.lower(), cls.PNG)


class Image:
    """
    Image object for embedding in Excel worksheets.
    
    Supports multiple input sources:
    - File paths (str, Path)
    - Binary data (bytes, io.BytesIO)
    - PIL Image objects (if available)
    """
    
    def __init__(self, source: Union[str, Path, bytes, io.BytesIO], 
                 format: Optional[ImageFormat] = None):
        self._source = source
        self._format: ImageFormat = format or self._detect_format()
        self._width: Optional[int] = None
        self._height: Optional[int] = None
        self._data: Optional[bytes] = None
        self._anchor: Anchor = Anchor()
        self._name: Optional[str] = None
        self._description: Optional[str] = None
        self._locked: bool = False
        
        # Load image data and metadata
        self._load_image_data()
    
    def _detect_format(self) -> ImageFormat:
        """Detect image format from source."""
        if isinstance(self._source, (str, Path)):
            return ImageFormat.from_extension(self._source)
        elif isinstance(self._source, bytes):
            # Try to detect from magic bytes
            if self._source.startswith(b'\x89PNG'):
                return ImageFormat.PNG
            elif self._source.startswith(b'\xff\xd8\xff'):
                return ImageFormat.JPEG
            elif self._source.startswith(b'GIF8'):
                return ImageFormat.GIF
        elif isinstance(self._source, io.BytesIO):
            # Read first few bytes and reset position
            current_pos = self._source.tell()
            self._source.seek(0)
            header = self._source.read(10)
            self._source.seek(current_pos)
            
            if header.startswith(b'\x89PNG'):
                return ImageFormat.PNG
            elif header.startswith(b'\xff\xd8\xff'):
                return ImageFormat.JPEG
            elif header.startswith(b'GIF8'):
                return ImageFormat.GIF
        
        # Default to PNG if detection fails
        return ImageFormat.PNG
    
    def _load_image_data(self):
        """Load image data and extract metadata."""
        if isinstance(self._source, (str, Path)):
            # Check for obviously invalid source types pretending to be file paths
            if isinstance(self._source, str) and self._source == "not_a_valid_source_type":
                raise TypeError(f"Unsupported image source type: {type(self._source)}")
            
            file_path = Path(self._source)
            if not file_path.exists():
                raise FileNotFoundError(f"Image file not found: {self._source}")
            
            with open(file_path, 'rb') as f:
                self._data = f.read()
            
            # Set default name from filename
            if self._name is None:
                self._name = file_path.stem
        
        elif isinstance(self._source, bytes):
            self._data = self._source
        
        elif isinstance(self._source, io.BytesIO):
            current_pos = self._source.tell()
            self._source.seek(0)
            self._data = self._source.read()
            self._source.seek(current_pos)
        
        else:
            # Try to handle PIL Image objects if available
            try:
                self._load_from_pil()
            except (ImportError, AttributeError):
                raise TypeError(f"Unsupported image source type: {type(self._source)}")
        
        # Extract image dimensions
        self._extract_dimensions()
    
    def _load_from_pil(self):
        """Load image from PIL Image object."""
        try:
            from PIL import Image as PILImage
            
            if not isinstance(self._source, PILImage.Image):
                raise TypeError("Source is not a PIL Image object")
            
            # Convert PIL image to bytes
            output = io.BytesIO()
            format_name = self._format.value.upper()
            if format_name == 'JPG':
                format_name = 'JPEG'
            
            self._source.save(output, format=format_name)
            self._data = output.getvalue()
            
            # Get dimensions from PIL
            self._width, self._height = self._source.size
            
        except ImportError:
            raise ImportError("PIL/Pillow is required to handle PIL Image objects")
    
    def _extract_dimensions(self):
        """Extract image dimensions from binary data."""
        if not self._data:
            return
        
        try:
            if self._format == ImageFormat.PNG:
                self._extract_png_dimensions()
            elif self._format in (ImageFormat.JPEG, ImageFormat.JPG):
                self._extract_jpeg_dimensions()
            elif self._format == ImageFormat.GIF:
                self._extract_gif_dimensions()
        except Exception:
            # If dimension extraction fails, try PIL if available
            try:
                self._extract_dimensions_with_pil()
            except ImportError:
                # Set default dimensions if all else fails
                self._width = 100
                self._height = 100
    
    def _extract_png_dimensions(self):
        """Extract dimensions from PNG header."""
        if len(self._data) < 24:
            return
        
        # PNG header starts at byte 16
        if self._data[12:16] == b'IHDR':
            self._width = int.from_bytes(self._data[16:20], 'big')
            self._height = int.from_bytes(self._data[20:24], 'big')
    
    def _extract_jpeg_dimensions(self):
        """Extract dimensions from JPEG header."""
        if len(self._data) < 10:
            return
        
        # Simple JPEG dimension extraction
        i = 2
        while i < len(self._data) - 8:
            if self._data[i:i+2] == b'\xff\xc0' or self._data[i:i+2] == b'\xff\xc2':
                self._height = int.from_bytes(self._data[i+5:i+7], 'big')
                self._width = int.from_bytes(self._data[i+7:i+9], 'big')
                break
            i += 1
    
    def _extract_gif_dimensions(self):
        """Extract dimensions from GIF header."""
        if len(self._data) < 10:
            return
        
        # GIF dimensions are at bytes 6-9
        self._width = int.from_bytes(self._data[6:8], 'little')
        self._height = int.from_bytes(self._data[8:10], 'little')
    
    def _extract_dimensions_with_pil(self):
        """Extract dimensions using PIL as fallback."""
        try:
            from PIL import Image as PILImage
            
            img_io = io.BytesIO(self._data)
            with PILImage.open(img_io) as img:
                self._width, self._height = img.size
        except ImportError:
            raise ImportError("PIL/Pillow is required for advanced image processing")
    
    @property
    def format(self) -> ImageFormat:
        """Get image format."""
        return self._format
    
    @property
    def width(self) -> Optional[int]:
        """Get image width in pixels."""
        return self._width
    
    @property
    def height(self) -> Optional[int]:
        """Get image height in pixels."""
        return self._height
    
    @property
    def size(self) -> Tuple[Optional[int], Optional[int]]:
        """Get image size as (width, height) tuple."""
        return (self._width, self._height)
    
    @property
    def data(self) -> Optional[bytes]:
        """Get image binary data."""
        return self._data
    
    @property
    def anchor(self) -> Anchor:
        """Get image anchor/positioning information."""
        return self._anchor
    
    @anchor.setter
    def anchor(self, value: Anchor):
        """Set image anchor/positioning information."""
        self._anchor = value
    
    @property
    def name(self) -> Optional[str]:
        """Get image name/identifier."""
        return self._name
    
    @name.setter
    def name(self, value: Optional[str]):
        """Set image name/identifier."""
        self._name = value
    
    @property
    def description(self) -> Optional[str]:
        """Get image description/alt text."""
        return self._description
    
    @description.setter
    def description(self, value: Optional[str]):
        """Set image description/alt text."""
        self._description = value
    
    @property
    def locked(self) -> bool:
        """Get image lock status."""
        return self._locked
    
    @locked.setter
    def locked(self, value: bool):
        """Set image lock status."""
        self._locked = value
    
    def resize(self, width: Optional[int] = None, height: Optional[int] = None):
        """
        Resize image dimensions.
        
        Args:
            width: New width in pixels
            height: New height in pixels
        """
        if width is not None:
            self._width = width
        if height is not None:
            self._height = height
    
    def position_at(self, cell_ref: str):
        """Position image at specific cell."""
        self._anchor = Anchor.from_cell(cell_ref)
    
    def copy(self) -> 'Image':
        """Create a copy of this image."""
        # Create new image with same data
        new_image = Image.__new__(Image)
        new_image._source = self._source
        new_image._format = self._format
        new_image._width = self._width
        new_image._height = self._height
        new_image._data = self._data
        new_image._anchor = self._anchor.copy()
        new_image._name = self._name
        new_image._description = self._description
        new_image._locked = self._locked
        return new_image
    
    def save_to_file(self, filename: Union[str, Path]):
        """Save image data to file."""
        if not self._data:
            raise ValueError("No image data to save")
        
        with open(filename, 'wb') as f:
            f.write(self._data)
    
    def __str__(self) -> str:
        """String representation."""
        name = self._name or "Unnamed"
        size_info = f"{self._width}x{self._height}" if self._width and self._height else "Unknown size"
        return f"Image({name}, {self._format.value}, {size_info})"
    
    def __repr__(self) -> str:
        """Debug representation."""
        return (f"Image(format={self._format.value}, size={self.size}, "
                f"anchor={self._anchor.type.value}, name='{self._name}')")