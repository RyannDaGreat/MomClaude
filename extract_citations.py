#!/usr/bin/env python3
"""
Extract paragraphs and their citations from a Word document.

Output format:
A markdown file with a numbered list where each paragraph is identified by
its first three words as a header, followed by numbered citations.

Includes both numeric citations (superscript) and Table references (Roman numerals).
"""

import zipfile
from xml.etree import ElementTree as ET
import re
from pathlib import Path


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

# Roman numeral pattern for tables
ROMAN_PATTERN = re.compile(r'\b[Tt]ables?\s+([IVXivx]+)\b')


def extract_paragraph_with_citations(p):
    """
    Extract text and citation references from a paragraph element.

    Returns:
        tuple: (text_string, list_of_citations_in_order)

    Citations are returned in the order they appear, and can be either:
    - Numeric citations (from superscript): "Citation 17"
    - Table references: "Table I"
    """
    # First pass: collect all runs with their properties
    runs = []  # List of (text, is_superscript, char_position)
    char_pos = 0

    for r in p.findall('.//w:r', NAMESPACES):
        rPr = r.find('w:rPr', NAMESPACES)
        is_superscript = False
        if rPr is not None:
            vertAlign = rPr.find('w:vertAlign', NAMESPACES)
            if vertAlign is not None:
                val = vertAlign.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if val == 'superscript':
                    is_superscript = True

        t = r.find('w:t', NAMESPACES)
        if t is not None and t.text:
            runs.append((t.text, is_superscript, char_pos))
            char_pos += len(t.text)

    # Build full text (non-superscript only for display)
    text_parts = [text for text, is_sup, _ in runs if not is_sup]
    full_text = ''.join(text_parts)

    # Build complete text including superscripts for position tracking
    all_text = ''.join(text for text, _, _ in runs)

    # Collect citations in order of appearance
    citations = []  # (position, citation_string)

    # Extract superscript citations
    current_pos = 0
    for text, is_sup, char_pos in runs:
        if is_sup:
            nums = re.findall(r'\d+', text)
            for num in nums:
                citations.append((char_pos, f"Citation {num}"))
        current_pos += len(text)

    # Extract Table references from full text (non-superscript)
    full_non_sup = ''.join(text for text, is_sup, _ in runs if not is_sup)

    # We need position in the original stream, so rebuild with positions
    non_sup_pos = 0
    char_to_orig_pos = {}
    for text, is_sup, orig_pos in runs:
        if not is_sup:
            for i, ch in enumerate(text):
                char_to_orig_pos[non_sup_pos + i] = orig_pos + i
            non_sup_pos += len(text)

    for match in ROMAN_PATTERN.finditer(full_non_sup):
        roman = match.group(1).upper()
        # Get approximate position in original stream
        match_pos = match.start()
        orig_pos = char_to_orig_pos.get(match_pos, match_pos)
        citations.append((orig_pos, f"Table {roman}"))

    # Sort by position
    citations.sort(key=lambda x: x[0])

    # Remove duplicates while preserving order
    seen = set()
    unique_citations = []
    for pos, cite in citations:
        if cite not in seen:
            seen.add(cite)
            unique_citations.append(cite)

    return full_text, unique_citations


def process_docx(docx_path):
    """
    Process a .docx file and extract paragraphs with their citations.

    Returns:
        list of tuples: [(first_three_words, full_text, [citations]), ...]
    """
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

    results = []
    for p in root.findall('.//w:p', NAMESPACES):
        text, citations = extract_paragraph_with_citations(p)
        text = text.strip()

        # Skip empty or very short paragraphs
        if not text or len(text) < 30:
            continue

        # Skip if no citations
        if not citations:
            continue

        # Get first three words
        words = text.split()
        first_three = ' '.join(words[:3])

        results.append((first_three, text, citations))

    return results


def generate_markdown(results, output_path):
    """Generate markdown output file."""
    lines = ["# Citation Extraction by Paragraph\n"]

    for i, (first_three, full_text, citations) in enumerate(results, 1):
        # Header: first three words
        lines.append(f"## {i}. {first_three}...")
        lines.append("")

        # Numbered list of citations (in order of appearance)
        for j, cite in enumerate(citations, 1):
            lines.append(f"- {j}. {cite}")

        lines.append("")

    output_path.write_text('\n'.join(lines))
    return len(results)


def main():
    docx_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/JAIP 6391.docx')
    output_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/citations_by_paragraph.md')

    print(f"Processing: {docx_path}")
    results = process_docx(docx_path)

    count = generate_markdown(results, output_path)
    print(f"Extracted {count} paragraphs with citations")
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
