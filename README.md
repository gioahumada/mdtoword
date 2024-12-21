# MDToWord

A Python package that converts Markdown files to Word documents with advanced formatting options.

## Installation

```bash
pip install mdtoword
```

## Usage

```python
from mdtoword import MarkdownToDocxConverter

# Create a converter instance
converter = MarkdownToDocxConverter()

# Convert markdown to docx
converter.convert_file("input.md", "output.docx")

# Or convert markdown string directly
markdown_text = "# Hello World\nThis is a **bold** text"
converter.convert_string(markdown_text, "output.docx")
```

## Features

- Converts Markdown to Word documents (.docx)
- Supports various Markdown formatting:
  - Headers
  - Bold and italic text
  - Lists (ordered and unordered)
  - Code blocks
  - Tables
  - Links
  - Images
- Customizable styles and formatting
- Page numbering support
- Table of contents generation

## Configuration

You can customize the document styling:

```python
converter = MarkdownToDocxConverter(
    styles={
        'document': {
            'font_name': 'Arial',
            'font_size': 12,
            'margins': {
                'top': 1.0,
                'bottom': 1.0,
                'left': 1.0,
                'right': 1.0
            }
        }
    }
)
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
