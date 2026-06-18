# Skill: docx — .docx Document Generator

Generate Microsoft Word (.docx) documents using `python-docx`.

## Requirements

Install the dependency:

```bash
pip install python-docx
```

## Usage

When the user asks to create a document:

1. Parse their request to determine content (headings, paragraphs, tables, lists, images, styling).
2. Write a Python script that uses `python-docx` (`from docx import Document`) to build the document.
3. Run the script to produce the `.docx` file.
4. Return the file path to the user.

## Script template

```python
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT

doc = Document()

# --- Content goes here ---
doc.add_heading("Title", level=0)
doc.add_paragraph("Body text here.")

doc.save("output.docx")
print("Document saved as output.docx")
```

## Styling helpers

| Style | Method |
|-------|--------|
| Bold / Italic | `run.bold = True`, `run.italic = True` |
| Font size | `run.font.size = Pt(12)` |
| Font color | `run.font.color.rgb = RGBColor(r, g, b)` |
| Alignment | `paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER` |
| Page margins | `section.left_margin = Cm(2.5)` |
| Tables | `doc.add_table(rows, cols)` then `cell.text = "..."` |

## File output

Save documents into the working directory with a descriptive filename.
