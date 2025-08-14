"""
Simplified styling system for Excel cells and ranges.
"""

from typing import Optional, Union


class Font:
    """Font styling properties."""
    
    def __init__(self):
        self.name: str = "Calibri"
        self.size: Union[int, float] = 11
        self.bold: bool = False
        self.italic: bool = False
        self.underline: bool = False
        self.color: str = "black"
    
    def copy(self) -> 'Font':
        """Create a copy of this font."""
        new_font = Font()
        new_font.name = self.name
        new_font.size = self.size
        new_font.bold = self.bold
        new_font.italic = self.italic
        new_font.underline = self.underline
        new_font.color = self.color
        return new_font


class Fill:
    """Fill/background styling properties."""
    
    def __init__(self):
        self.color: str = "white"
        self.pattern: str = "none"
        self.gradient: Optional[str] = None
    
    def copy(self) -> 'Fill':
        """Create a copy of this fill."""
        new_fill = Fill()
        new_fill.color = self.color
        new_fill.pattern = self.pattern
        new_fill.gradient = self.gradient
        return new_fill


class BorderSide:
    """Individual border side styling."""
    
    def __init__(self):
        self.style: str = "none"  # none, thin, thick, medium, dashed, dotted, double
        self.color: str = "black"
    
    def copy(self) -> 'BorderSide':
        """Create a copy of this border side."""
        new_side = BorderSide()
        new_side.style = self.style
        new_side.color = self.color
        return new_side


class Border:
    """Comprehensive border styling properties."""
    
    def __init__(self):
        self._left: Optional[BorderSide] = None
        self._right: Optional[BorderSide] = None
        self._top: Optional[BorderSide] = None
        self._bottom: Optional[BorderSide] = None
        self._diagonal: Optional[BorderSide] = None
        self._diagonal_up: bool = False
        self._diagonal_down: bool = False
    
    @property
    def left(self) -> BorderSide:
        """Get or create left border."""
        if self._left is None:
            self._left = BorderSide()
        return self._left
    
    @left.setter
    def left(self, value: BorderSide):
        """Set left border."""
        self._left = value
    
    @property
    def right(self) -> BorderSide:
        """Get or create right border."""
        if self._right is None:
            self._right = BorderSide()
        return self._right
    
    @right.setter
    def right(self, value: BorderSide):
        """Set right border."""
        self._right = value
    
    @property
    def top(self) -> BorderSide:
        """Get or create top border."""
        if self._top is None:
            self._top = BorderSide()
        return self._top
    
    @top.setter
    def top(self, value: BorderSide):
        """Set top border."""
        self._top = value
    
    @property
    def bottom(self) -> BorderSide:
        """Get or create bottom border."""
        if self._bottom is None:
            self._bottom = BorderSide()
        return self._bottom
    
    @bottom.setter
    def bottom(self, value: BorderSide):
        """Set bottom border."""
        self._bottom = value
    
    @property
    def diagonal(self) -> BorderSide:
        """Get or create diagonal border."""
        if self._diagonal is None:
            self._diagonal = BorderSide()
        return self._diagonal
    
    @diagonal.setter
    def diagonal(self, value: BorderSide):
        """Set diagonal border."""
        self._diagonal = value
    
    def set_all_borders(self, style: str = "thin", color: str = "black"):
        """Set all borders to the same style and color."""
        for side in ['left', 'right', 'top', 'bottom']:
            border_side = getattr(self, side)
            border_side.style = style
            border_side.color = color
    
    def set_outline(self, style: str = "thin", color: str = "black"):
        """Set outline borders (all four sides)."""
        self.set_all_borders(style, color)
    
    def remove_all_borders(self):
        """Remove all borders."""
        self._left = None
        self._right = None
        self._top = None
        self._bottom = None
        self._diagonal = None
    
    def copy(self) -> 'Border':
        """Create a copy of this border."""
        new_border = Border()
        if self._left:
            new_border._left = self._left.copy()
        if self._right:
            new_border._right = self._right.copy()
        if self._top:
            new_border._top = self._top.copy()
        if self._bottom:
            new_border._bottom = self._bottom.copy()
        if self._diagonal:
            new_border._diagonal = self._diagonal.copy()
        new_border._diagonal_up = self._diagonal_up
        new_border._diagonal_down = self._diagonal_down
        return new_border


class Alignment:
    """Text alignment properties."""
    
    def __init__(self):
        self.horizontal: str = "general"
        self.vertical: str = "bottom"
        self.wrap_text: bool = False
        self.shrink_to_fit: bool = False
        self.indent: int = 0
        self.text_rotation: int = 0
    
    def copy(self) -> 'Alignment':
        """Create a copy of this alignment."""
        new_alignment = Alignment()
        new_alignment.horizontal = self.horizontal
        new_alignment.vertical = self.vertical
        new_alignment.wrap_text = self.wrap_text
        new_alignment.shrink_to_fit = self.shrink_to_fit
        new_alignment.indent = self.indent
        new_alignment.text_rotation = self.text_rotation
        return new_alignment


class Style:
    """Complete cell style container."""
    
    def __init__(self):
        self._font: Optional[Font] = None
        self._fill: Optional[Fill] = None
        self._border: Optional[Border] = None
        self._alignment: Optional[Alignment] = None
        self._number_format: str = "General"
        self._protection: bool = True
    
    @property
    def font(self) -> Font:
        """Get or create font styling."""
        if self._font is None:
            self._font = Font()
        return self._font
    
    @font.setter
    def font(self, value: Font):
        """Set font styling."""
        self._font = value
    
    @property
    def fill(self) -> Fill:
        """Get or create fill styling."""
        if self._fill is None:
            self._fill = Fill()
        return self._fill
    
    @fill.setter
    def fill(self, value: Fill):
        """Set fill styling."""
        self._fill = value
    
    @property
    def border(self) -> Border:
        """Get or create border styling."""
        if self._border is None:
            self._border = Border()
        return self._border
    
    @border.setter
    def border(self, value: Border):
        """Set border styling."""
        self._border = value
    
    @property
    def alignment(self) -> Alignment:
        """Get or create alignment styling."""
        if self._alignment is None:
            self._alignment = Alignment()
        return self._alignment
    
    @alignment.setter
    def alignment(self, value: Alignment):
        """Set alignment styling."""
        self._alignment = value
    
    @property
    def number_format(self) -> str:
        """Get number format."""
        return self._number_format
    
    @number_format.setter
    def number_format(self, value: str):
        """Set number format."""
        self._number_format = value
    
    @property
    def protection(self) -> bool:
        """Get protection status."""
        return self._protection
    
    @protection.setter
    def protection(self, value: bool):
        """Set protection status."""
        self._protection = value
    
    def copy(self) -> 'Style':
        """Create a deep copy of this style."""
        new_style = Style()
        if self._font:
            new_style._font = self._font.copy()
        if self._fill:
            new_style._fill = self._fill.copy()
        if self._border:
            new_style._border = self._border.copy()
        if self._alignment:
            new_style._alignment = self._alignment.copy()
        new_style._number_format = self._number_format
        new_style._protection = self._protection
        return new_style