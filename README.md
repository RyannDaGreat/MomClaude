# Citation Extractor

Extracts citations from Word documents, detects duplicate references, and generates a modified document with renumbered citations.

## Usage

```bash
python extract_citations.py
```

Edit the `docx_path` variable in `main()` to point to your document.

## What it does

1. **Extracts citations** from paragraphs and tables
2. **Detects duplicate references** using 90% text similarity
3. **Generates a markdown report** (`citations_by_paragraph.md`) with:
   - Citations by table
   - Citations by paragraph
   - Duplicate reference mappings
   - Modification plan
4. **Creates a modified Word document** (`*_modified.docx`) showing:
   - Old citation numbers in red with strikethrough
   - New citation numbers in green bold
   - Duplicate references struck through
   - Kept references renumbered to fill gaps

## Requirements

- Python 3.10+
- python-docx
- lxml
