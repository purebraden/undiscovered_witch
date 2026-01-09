# KDP Document Generator

This script generates a Kindle Direct Publishing (KDP) ready DOCX file from your markdown chapter files.

## Features

- ✅ Title page with book title and subtitle
- ✅ Copyright page with placeholder fields
- ✅ Auto-generated Table of Contents (TOC)
- ✅ All chapters with proper page breaks
- ✅ Each chapter (prologue, chapters, epilogue) starts on a new page
- ✅ Proper formatting with Times New Roman font
- ✅ First-line paragraph indentation
- ✅ Support for markdown formatting (bold, italic, headers, lists)

## Installation

1. Install Python 3.7 or higher
2. Install required dependencies:

```bash
py -m pip install -r requirements.txt
```

(On Windows, use `py -m pip` instead of `pip`)

## Usage

1. Make sure all your markdown files are in the `book1/` directory:
   - `prologue.md`
   - `chapter_01.md` through `chapter_14.md`
   - `epilogue.md`

2. Run the script:

```bash
py generate_kdp_docx.py
```

(On Windows, use `py` instead of `python`)

3. The script will generate `book1/Undiscovered_Witch_KDP_Interior_Final.docx`

## After Generation

1. **Update the TOC in Word:**
   - Open the generated DOCX file in Microsoft Word
   - Right-click on the Table of Contents
   - Select "Update Field"
   - Choose "Update entire table"
   - Or press `Ctrl+A` then `F9` to update all fields

2. **Update placeholder information:**
   - Author name (currently shows `[AUTHOR NAME]`)
   - Copyright year (currently `2024`)
   - ISBN number
   - Cover designer name
   - Publisher name

3. **Review formatting:**
   - Check that all chapters start on new pages
   - Verify paragraph indentation
   - Ensure proper spacing between chapters
   - Check that headings are formatted correctly

## Customization

You can edit `generate_kdp_docx.py` to customize:

- Book title and subtitle (lines ~140-141)
- Author name (line ~142)
- Copyright year (line ~143)
- Font sizes and styles
- Page layout settings

## Notes

- The script preserves markdown formatting (bold `**text**`, italic `*text*`)
- Headers (#, ##, ###) are converted to Word heading styles
- Bullet points and numbered lists are preserved
- The TOC will automatically pick up Heading 1, 2, and 3 styles

