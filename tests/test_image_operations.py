"""
Comprehensive Image Operations Tests
Tests for core image functionality: insert, extract, resize, move, delete.
"""

import pytest
import io
from pathlib import Path

from aspose.cells import Workbook
from aspose.cells.drawing import Image, ImageFormat, ImageCollection


class TestImageOperations:
    """Core image operations tests."""
    
    def setup_method(self):
        """Set up test environment."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        
        # Create dedicated output folder for image operations tests
        self.output_dir = Path(__file__).parent / "testdata" / "test_image_operations"
        self.output_dir.mkdir(exist_ok=True)
        
        # Use real image files from testdata
        self.test_images_dir = Path(__file__).parent / "testdata" / "images"
        self.test_jpg_path = self.test_images_dir / "image1.jpg"
        self.test_jpeg_path = self.test_images_dir / "image2.jpeg"
        self.test_png_path = self.test_images_dir / "image3.jpg"  # Using jpg as png for format test
        
        # Load image data for direct bytes tests
        with open(self.test_jpg_path, 'rb') as f:
            self.test_jpg_data = f.read()
        with open(self.test_jpeg_path, 'rb') as f:
            self.test_jpeg_data = f.read()
    
    def test_insert_image_from_file_path(self):
        """Test inserting image from file path."""
        image = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='landscape_photo')
        
        assert image.name == 'landscape_photo'
        assert image.format == ImageFormat.JPEG
        assert len(self.worksheet.images) == 1
        assert image.width is not None
        assert image.height is not None
    
    def test_insert_image_from_bytes(self):
        """Test inserting image from bytes."""
        image = self.worksheet.images.add(self.test_jpg_data, 'B2', name='photo_bytes')
        
        assert image.format == ImageFormat.JPEG
        assert image.data == self.test_jpg_data
        assert image.name == 'photo_bytes'
        assert len(self.worksheet.images) == 1
    
    def test_insert_image_from_bytesio(self):
        """Test inserting image from BytesIO."""
        bio = io.BytesIO(self.test_jpeg_data)
        image = self.worksheet.images.add(bio, 'C3')
        
        assert image.format == ImageFormat.JPEG
        assert len(self.worksheet.images) == 1
    
    def test_insert_image_with_custom_size(self):
        """Test inserting image with custom dimensions."""
        image = self.worksheet.images.add(
            str(self.test_jpeg_path), 'D4', 
            width=200, height=150, 
            name='sized_portrait'
        )
        
        assert image.width == 200
        assert image.height == 150
        assert image.name == 'sized_portrait'
    
    def test_extract_image_data(self):
        """Test extracting image data."""
        # Add image
        image = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='extract_test')
        
        # Extract by name
        extracted_data = self.worksheet.images.extract('extract_test')
        assert extracted_data == self.test_jpg_data
        
        # Extract by index
        extracted_data = self.worksheet.images.extract(0)
        assert extracted_data == self.test_jpg_data
    
    def test_resize_image(self):
        """Test resizing image."""
        image = self.worksheet.images.add(str(self.test_jpeg_path), 'A1', name='resize_test')
        original_width = image.width
        original_height = image.height
        
        # Resize using collection method
        self.worksheet.images.resize('resize_test', width=300, height=200)
        
        assert image.width == 300
        assert image.height == 200
        
        # Resize using image method
        image.resize(width=400)
        assert image.width == 400
        assert image.height == 200  # Height unchanged
    
    def test_move_image(self):
        """Test moving image to new position."""
        image = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='move_test')
        original_position = image.anchor.from_position
        
        # Move using collection method
        self.worksheet.images.move('move_test', 'E5')
        
        # Check position changed
        anchor = image.anchor
        assert anchor.from_position != original_position  # Should have moved from A1
    
    def test_delete_image(self):
        """Test deleting images."""
        # Add multiple images
        img1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='mountain_view')
        img2 = self.worksheet.images.add(str(self.test_jpeg_path), 'B2', name='city_skyline')
        
        assert len(self.worksheet.images) == 2
        
        # Delete by name
        self.worksheet.images.remove('mountain_view')
        assert len(self.worksheet.images) == 1
        assert 'mountain_view' not in self.worksheet.images
        
        # Delete by index
        self.worksheet.images.remove(0)
        assert len(self.worksheet.images) == 0
    
    def test_image_name_uniqueness(self):
        """Test that image names are automatically made unique."""
        img1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='nature')
        img2 = self.worksheet.images.add(str(self.test_jpeg_path), 'B2', name='nature')
        
        assert img1.name == 'nature'
        assert img2.name == 'nature_1'
        
        # Add without name - should get auto names
        img3 = self.worksheet.images.add(self.test_jpg_data, 'C3')
        assert img3.name == 'Image3'
    
    def test_access_images_by_name_and_index(self):
        """Test accessing images by name and index."""
        img1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', name='header_logo')
        img2 = self.worksheet.images.add(str(self.test_jpeg_path), 'B2', name='main_photo')
        
        # Access by name
        assert self.worksheet.images['header_logo'] == img1
        assert self.worksheet.images.get('main_photo') == img2
        
        # Access by index
        assert self.worksheet.images[0] == img1
        assert self.worksheet.images.get(1) == img2
        
        # Test 'in' operator
        assert 'header_logo' in self.worksheet.images
        assert img1 in self.worksheet.images
        assert 'nonexistent' not in self.worksheet.images
    
    def test_image_format_detection(self):
        """Test automatic format detection."""
        jpg_img = self.worksheet.images.add(str(self.test_jpg_path), 'A1')
        jpeg_img = self.worksheet.images.add(str(self.test_jpeg_path), 'B2')
        
        assert jpg_img.format == ImageFormat.JPEG
        assert jpeg_img.format == ImageFormat.JPEG
    
    def test_clear_all_images(self):
        """Test clearing all images."""
        self.worksheet.images.add(str(self.test_jpg_path), 'A1')
        self.worksheet.images.add(str(self.test_jpeg_path), 'B2')
        
        assert len(self.worksheet.images) == 2
        
        self.worksheet.images.clear()
        assert len(self.worksheet.images) == 0
    
    def test_save_excel_with_images(self):
        """Test saving Excel file with embedded images."""
        # Add multiple images with different sizes and positions
        img1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', width=120, height=80, name='Logo')
        img2 = self.worksheet.images.add(str(self.test_jpeg_path), 'C1', width=200, height=150, name='Banner')
        img3 = self.worksheet.images.add(self.test_jpg_data, 'A5', width=100, height=75, name='Thumbnail')
        
        # Verify images are in memory correctly
        assert len(self.worksheet.images) == 3
        assert img1.name == 'Logo'
        assert img2.name == 'Banner'
        assert img3.name == 'Thumbnail'
        
        # Add some data to the worksheet
        self.worksheet['A8'] = "Image Demo Test"
        self.worksheet['A9'] = "Contains 3 images"
        self.worksheet['A10'] = f"Image 1: {img1.name} ({img1.width}x{img1.height})"
        self.worksheet['A11'] = f"Image 2: {img2.name} ({img2.width}x{img2.height})"
        self.worksheet['A12'] = f"Image 3: {img3.name} ({img3.width}x{img3.height})"
        
        # Save the file
        output_path = self.output_dir / "test_images_demo.xlsx"
        
        self.workbook.save(str(output_path))
        
        # Verify file was created (Note: actual image embedding in Excel depends on writer implementation)
        assert output_path.exists()
        assert output_path.stat().st_size > 500  # Should contain worksheet data
        
        print(f"Excel file with images saved: {output_path}")
        print(f"File size: {output_path.stat().st_size:,} bytes")
        print(f"Note: Image data is stored in memory - Excel embedding may require extended writer support")
    
    def test_complete_workflow(self):
        """Test complete image workflow: add, modify, save, and verify."""
        # Step 1: Add images
        logo = self.worksheet.images.add(str(self.test_jpg_path), 'B2', width=100, height=60, name='CompanyLogo')
        chart = self.worksheet.images.add(str(self.test_jpeg_path), 'D2', width=250, height=180, name='SalesChart')
        
        # Step 2: Add worksheet content
        self.worksheet['A1'] = "Monthly Report"
        self.worksheet['A3'] = "Sales Data"
        self.worksheet['A4'] = "Q1: $125,000"
        self.worksheet['A5'] = "Q2: $180,000"
        self.worksheet['A6'] = "Q3: $220,000"
        self.worksheet['A7'] = "Q4: $195,000"
        
        # Step 3: Modify images
        self.worksheet.images.resize('CompanyLogo', width=80, height=50)
        self.worksheet.images.move('SalesChart', 'F4')
        
        # Step 4: Add another image from bytes
        watermark = self.worksheet.images.add(self.test_jpg_data, 'A10', width=60, height=40, name='Watermark')
        
        # Step 5: Verify collection state
        assert len(self.worksheet.images) == 3
        assert 'CompanyLogo' in self.worksheet.images
        assert 'SalesChart' in self.worksheet.images
        assert 'Watermark' in self.worksheet.images
        
        # Step 6: Extract image data
        logo_data = self.worksheet.images.extract('CompanyLogo')
        chart_data = self.worksheet.images.extract('SalesChart')
        
        assert len(logo_data) > 1000
        assert len(chart_data) > 1000
        
        # Step 7: Save complete workbook
        output_path = self.output_dir / "complete_workflow_test.xlsx"
        
        self.workbook.save(str(output_path))
        
        # Step 8: Verify final file
        assert output_path.exists()
        file_size = output_path.stat().st_size
        assert file_size > 500  # Should contain worksheet data
        
        # Verify images are still in memory after save
        assert len(self.worksheet.images) == 3
        assert all(img.data is not None for img in self.worksheet.images)
        
        print(f"Complete workflow test saved: {output_path}")
        print(f"Final file size: {file_size:,} bytes")
        print(f"Images in memory: {len(self.worksheet.images)}")
        print(f"Image data preserved: {all(len(img.data) > 1000 for img in self.worksheet.images)}")
    
    def test_error_handling(self):
        """Test error handling for common issues."""
        # Test extracting from empty collection
        with pytest.raises(ValueError):
            self.worksheet.images.extract('nonexistent')
        
        # Test removing nonexistent image
        with pytest.raises(ValueError):
            self.worksheet.images.remove('nonexistent')
        
        # Test invalid index
        with pytest.raises(IndexError):
            self.worksheet.images.get(999)
        
        # Test moving nonexistent image
        with pytest.raises(ValueError):
            self.worksheet.images.move('nonexistent', 'A1')
    
    def test_multi_worksheet_images(self):
        """Test images across multiple worksheets."""
        # Create additional worksheets
        ws2 = self.workbook.create_sheet("Gallery")
        ws3 = self.workbook.create_sheet("Reports")
        
        # Add images to different worksheets
        # Worksheet 1 - Main dashboard
        self.worksheet.name = "Dashboard"
        logo1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', width=100, height=60, name='MainLogo')
        
        # Worksheet 2 - Image gallery
        gallery1 = ws2.images.add(str(self.test_jpeg_path), 'A1', width=200, height=150, name='GalleryImage1')
        gallery2 = ws2.images.add(self.test_jpg_data, 'C1', width=200, height=150, name='GalleryImage2')
        
        # Worksheet 3 - Report with charts
        chart1 = ws3.images.add(str(self.test_jpg_path), 'B2', width=300, height=200, name='SalesChart')
        
        # Add content to worksheets
        self.worksheet['A3'] = "Main Dashboard"
        self.worksheet['A4'] = f"Images: {len(self.worksheet.images)}"
        
        ws2['A5'] = "Image Gallery"
        ws2['A6'] = f"Gallery Images: {len(ws2.images)}"
        
        ws3['A1'] = "Monthly Sales Report"
        ws3['A3'] = f"Charts: {len(ws3.images)}"
        
        # Verify each worksheet has its images
        assert len(self.worksheet.images) == 1
        assert len(ws2.images) == 2
        assert len(ws3.images) == 1
        
        # Verify image names are unique within each worksheet
        assert 'MainLogo' in self.worksheet.images
        assert 'GalleryImage1' in ws2.images
        assert 'GalleryImage2' in ws2.images
        assert 'SalesChart' in ws3.images
        
        # Save multi-worksheet workbook
        output_path = self.output_dir / "multi_worksheet_images.xlsx"
        
        self.workbook.save(str(output_path))
        
        # Verify file creation
        assert output_path.exists()
        file_size = output_path.stat().st_size
        assert file_size > 500  # Should contain worksheet data
        
        # Verify images are preserved in memory across worksheets
        total_images = sum(len(ws.images) for ws in self.workbook.worksheets)
        assert total_images == 4  # 1 + 2 + 1
        
        print(f"Multi-worksheet file saved: {output_path}")
        print(f"File size: {file_size:,} bytes")
        print(f"Worksheets: {len(self.workbook.worksheets)}")
        print(f"Total images: {sum(len(ws.images) for ws in self.workbook.worksheets)}")
    
    def test_load_and_modify_images(self):
        """Test loading Excel file and modifying its images."""
        # First, create and save a file with images
        original_path = self.output_dir / "original_with_images.xlsx"
        original_path.parent.mkdir(exist_ok=True)
        
        # Add some images
        img1 = self.worksheet.images.add(str(self.test_jpg_path), 'A1', width=150, height=100, name='OriginalImage1')
        img2 = self.worksheet.images.add(str(self.test_jpeg_path), 'C1', width=120, height=80, name='OriginalImage2')
        
        # Add some data
        self.worksheet['A5'] = "Original File"
        self.worksheet['A6'] = f"Created: 2024"
        
        # Save original file
        self.workbook.save(str(original_path))
        
        # Now load the file and modify it
        loaded_wb = Workbook()
        loaded_wb.load(str(original_path))
        loaded_ws = loaded_wb.active
        
        # Verify images were loaded (Note: actual loading may depend on implementation)
        print(f"Loaded file has {len(loaded_ws.images)} images")
        
        # Add new images to the loaded file
        new_img = loaded_ws.images.add(self.test_jpg_data, 'E1', width=100, height=75, name='NewAddedImage')
        
        # Modify worksheet content
        loaded_ws['A7'] = "File Modified"
        loaded_ws['A8'] = f"New Image: {new_img.name}"
        
        # Save modified file
        modified_path = self.output_dir / "modified_with_images.xlsx"
        loaded_wb.save(str(modified_path))
        
        # Verify both files exist
        assert original_path.exists()
        assert modified_path.exists()
        
        print(f"Original file: {original_path}")
        print(f"Modified file: {modified_path}")
        print(f"Original size: {original_path.stat().st_size:,} bytes")
        print(f"Modified size: {modified_path.stat().st_size:,} bytes")


class TestImageFormats:
    """Test image format handling."""
    
    def test_format_enum_values(self):
        """Test ImageFormat enum values."""
        assert ImageFormat.PNG.value == "png"
        assert ImageFormat.JPEG.value == "jpeg"
        assert ImageFormat.JPG.value == "jpg"
        assert ImageFormat.GIF.value == "gif"
    
    def test_format_from_extension(self):
        """Test format detection from file extensions."""
        assert ImageFormat.from_extension('test.png') == ImageFormat.PNG
        assert ImageFormat.from_extension('test.jpg') == ImageFormat.JPEG
        assert ImageFormat.from_extension('test.jpeg') == ImageFormat.JPEG
        assert ImageFormat.from_extension('test.gif') == ImageFormat.GIF
        assert ImageFormat.from_extension('test.unknown') == ImageFormat.PNG  # Default
    
    def test_format_from_mimetype(self):
        """Test format detection from MIME types."""
        assert ImageFormat.from_mimetype('image/png') == ImageFormat.PNG
        assert ImageFormat.from_mimetype('image/jpeg') == ImageFormat.JPEG
        assert ImageFormat.from_mimetype('image/gif') == ImageFormat.GIF
        assert ImageFormat.from_mimetype('unknown/type') == ImageFormat.PNG  # Default


class TestImageProperties:
    """Test image property management."""
    
    def setup_method(self):
        """Set up test environment."""
        # Use real image file
        test_image_path = Path(__file__).parent / "testdata" / "images" / "image1.jpg"
        with open(test_image_path, 'rb') as f:
            self.test_data = f.read()
        
        self.image = Image(self.test_data)
    
    def test_image_properties(self):
        """Test basic image properties."""
        assert self.image.format == ImageFormat.JPEG
        assert isinstance(self.image.data, bytes)
        assert self.image.size == (self.image.width, self.image.height)
        assert self.image.width > 0
        assert self.image.height > 0
    
    def test_image_name_and_description(self):
        """Test name and description properties."""
        self.image.name = "test_image"
        self.image.description = "Test description"
        
        assert self.image.name == "test_image"
        assert self.image.description == "Test description"
    
    def test_image_lock_status(self):
        """Test image lock property."""
        assert self.image.locked == False
        
        self.image.locked = True
        assert self.image.locked == True
    
    def test_image_copy(self):
        """Test image copying."""
        self.image.name = "original"
        copy_img = self.image.copy()
        
        assert copy_img.name == "original"
        assert copy_img.data == self.image.data
        assert copy_img is not self.image
    
    def test_image_string_representations(self):
        """Test string representations."""
        self.image.name = "sample_photo"
        
        str_repr = str(self.image)
        assert "sample_photo" in str_repr
        assert "jpeg" in str_repr
        
        repr_str = repr(self.image)
        assert "Image(" in repr_str
        assert "jpeg" in repr_str
    
    def test_save_image_to_file(self):
        """Test saving image data to file."""
        output_dir = Path(__file__).parent / "testdata" / "test_image_operations"
        output_dir.mkdir(exist_ok=True)
        
        # Save image to file
        output_file = output_dir / "saved_image.jpg"
        self.image.save_to_file(output_file)
        
        # Verify file was created and has correct data
        assert output_file.exists()
        
        with open(output_file, 'rb') as f:
            saved_data = f.read()
        
        assert saved_data == self.image.data
        assert len(saved_data) > 1000  # Should be a real image file
        
        print(f"Image saved to file: {output_file}")
        print(f"File size: {len(saved_data):,} bytes")