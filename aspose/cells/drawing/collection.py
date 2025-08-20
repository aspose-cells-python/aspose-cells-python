"""
Image collection management for worksheets.
"""

from typing import List, Union, Optional, Iterator
from pathlib import Path
import io

from .image import Image, ImageFormat
from .anchor import Anchor


class ImageCollection:
    """Collection manager for worksheet images with multiple access patterns."""
    
    def __init__(self, worksheet: 'Worksheet'):
        self._worksheet = worksheet
        self._images: List[Image] = []
    
    def add(self, source: Union[str, Path, bytes, io.BytesIO, Image], 
            cell_ref: str = "A1", 
            width: Optional[int] = None,
            height: Optional[int] = None,
            name: Optional[str] = None) -> Image:
        """
        Add image to worksheet at specified cell.
        
        Args:
            source: Image source (file path, bytes, Image object)
            cell_ref: Cell reference for positioning (default: "A1")
            width: Optional width in pixels
            height: Optional height in pixels  
            name: Optional image name
            
        Returns:
            Image: Added image object
        """
        if isinstance(source, Image):
            image = source.copy()
        else:
            image = Image(source)
        
        # Set positioning
        image.position_at(cell_ref)
        
        # Resize if specified
        if width is not None or height is not None:
            image.resize(width, height)
        
        # Set name
        if name is not None:
            image.name = name
        elif image.name is None:
            image.name = f"Image{len(self._images) + 1}"
        
        # Ensure unique name
        original_name = image.name
        counter = 1
        while any(img.name == image.name for img in self._images):
            image.name = f"{original_name}_{counter}"
            counter += 1
        
        self._images.append(image)
        return image
    
    def extract(self, target: Union[str, int]) -> bytes:
        """
        Extract image data as bytes.
        
        Args:
            target: Image identifier (name or index)
            
        Returns:
            bytes: Image binary data
        """
        image = self.get(target)
        if not image.data:
            raise ValueError("Image has no data to extract")
        return image.data
    
    def move(self, target: Union[str, int], cell_ref: str):
        """
        Move image to new cell position.
        
        Args:
            target: Image identifier (name or index)
            cell_ref: New cell reference
        """
        image = self.get(target)
        image.position_at(cell_ref)
    
    def resize(self, target: Union[str, int], width: Optional[int] = None, height: Optional[int] = None):
        """
        Resize image.
        
        Args:
            target: Image identifier (name or index)
            width: New width in pixels
            height: New height in pixels
        """
        image = self.get(target)
        image.resize(width, height)
    
    def remove(self, target: Union[str, int, Image]):
        """
        Remove image by name, index, or object.
        
        Args:
            target: Image identifier (name, index, or Image object)
        """
        if isinstance(target, Image):
            if target in self._images:
                self._images.remove(target)
            else:
                raise ValueError("Image not found in collection")
        
        elif isinstance(target, str):
            # Remove by name
            for i, image in enumerate(self._images):
                if image.name == target:
                    del self._images[i]
                    return
            raise ValueError(f"Image with name '{target}' not found")
        
        elif isinstance(target, int):
            # Remove by index
            if 0 <= target < len(self._images):
                del self._images[target]
            else:
                raise IndexError(f"Image index {target} out of range")
        
        else:
            raise TypeError(f"Invalid target type: {type(target)}")
    
    def get(self, identifier: Union[str, int]) -> Image:
        """
        Get image by name or index.
        
        Args:
            identifier: Image name or index
            
        Returns:
            Image: Found image object
        """
        if isinstance(identifier, str):
            for image in self._images:
                if image.name == identifier:
                    return image
            raise ValueError(f"Image with name '{identifier}' not found")
        
        elif isinstance(identifier, int):
            if 0 <= identifier < len(self._images):
                return self._images[identifier]
            else:
                raise IndexError(f"Image index {identifier} out of range")
        
        else:
            raise TypeError(f"Invalid identifier type: {type(identifier)}")
    
    def clear(self):
        """Remove all images from the collection."""
        self._images.clear()
    
    def get_by_position(self, cell_ref: str) -> List[Image]:
        """
        Get all images positioned at or near a specific cell.
        
        Args:
            cell_ref: Cell reference (e.g., "A1", "B2")
            
        Returns:
            List[Image]: Images at the specified position
        """
        from ..utils.coordinates import coordinate_to_tuple
        
        target_row, target_col = coordinate_to_tuple(cell_ref)
        target_row -= 1  # Convert to 0-based
        target_col -= 1
        
        matching_images = []
        for image in self._images:
            anchor = image.anchor
            from_row, from_col = anchor.from_position
            
            # Check if image starts at this position
            if from_row == target_row and from_col == target_col:
                matching_images.append(image)
            
            # For TWO_CELL anchors, check if position is within range
            elif anchor.to_position is not None:
                to_row, to_col = anchor.to_position
                if (from_row <= target_row <= to_row and 
                    from_col <= target_col <= to_col):
                    matching_images.append(image)
        
        return matching_images
    
    @property
    def names(self) -> List[str]:
        """Get list of all image names."""
        return [img.name for img in self._images if img.name]
    
    def __len__(self) -> int:
        """Number of images in collection."""
        return len(self._images)
    
    def __iter__(self) -> Iterator[Image]:
        """Iterate over images."""
        return iter(self._images)
    
    def __getitem__(self, key: Union[str, int]) -> Image:
        """Get image by name or index."""
        return self.get(key)
    
    def __contains__(self, item: Union[str, Image]) -> bool:
        """Check if image exists in collection."""
        if isinstance(item, str):
            return any(img.name == item for img in self._images)
        elif isinstance(item, Image):
            return item in self._images
        else:
            return False
    
    def __str__(self) -> str:
        """String representation."""
        return f"ImageCollection({len(self._images)} images)"
    
    def __repr__(self) -> str:
        """Debug representation."""
        names = [img.name for img in self._images[:3]]  # Show first 3
        if len(self._images) > 3:
            names.append("...")
        return f"ImageCollection(images={names})"