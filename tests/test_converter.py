import pytest
from mdtoword import MarkdownToDocxConverter
import os
from docx import Document

@pytest.fixture
def converter():
    return MarkdownToDocxConverter()

@pytest.fixture
def test_output_file():
    file_path = "test_output.docx"
    yield file_path
    # Cleanup after test
    if os.path.exists(file_path):
        os.remove(file_path)

def test_basic_conversion(converter, test_output_file):
    """Test basic markdown to docx conversion"""
    markdown_text = "# Test Heading\nThis is a test paragraph."
    converter.convert_string(markdown_text, test_output_file)
    
    # Verify the file was created
    assert os.path.exists(test_output_file)
    
    # Open and verify content
    doc = Document(test_output_file)
    paragraphs = [p.text for p in doc.paragraphs if p.text]
    assert "Test Heading" in paragraphs[0]
    assert "This is a test paragraph." in paragraphs[1]

def test_formatting(converter, test_output_file):
    """Test various markdown formatting options"""
    markdown_text = """
# Heading 1
## Heading 2
**Bold text**
*Italic text*
- List item 1
- List item 2
"""
    converter.convert_string(markdown_text, test_output_file)
    assert os.path.exists(test_output_file)

def test_custom_styles(test_output_file):
    """Test converter with custom styles"""
    custom_styles = {
        'document': {
            'font_name': 'Arial',
            'font_size': 12,
            'margins': {
                'top': 2.0,
                'bottom': 2.0,
                'left': 2.0,
                'right': 2.0
            }
        }
    }
    converter = MarkdownToDocxConverter(styles=custom_styles)
    markdown_text = "# Custom Style Test"
    converter.convert_string(markdown_text, test_output_file)
    assert os.path.exists(test_output_file)

def test_file_conversion(converter, test_output_file, tmp_path):
    """Test conversion from markdown file"""
    # Create a temporary markdown file
    md_file = tmp_path / "test.md"
    md_file.write_text("# Test File\nThis is a test.")
    
    converter.convert_file(str(md_file), test_output_file)
    assert os.path.exists(test_output_file)
