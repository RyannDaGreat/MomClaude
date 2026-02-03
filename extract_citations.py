#!/usr/bin/env python3
"""
Extract paragraphs and their citations from a Word document.

Output format:
A markdown file with a numbered list where each paragraph is identified by
its first three words as a header, followed by numbered citations.
"""

import zipfile
from xml.etree import ElementTree as ET
import re
from pathlib import Path


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}


def extract_paragraph_with_citations(p):
    """
    Extract text and citation numbers from a paragraph element.

    Returns:
        tuple: (text_string, list_of_citation_numbers)
    """
    text_parts = []
    citations = []

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
            if is_superscript:
                # Extract numeric citations from superscript text
                nums = re.findall(r'\d+', t.text)
                citations.extend(nums)
            else:
                text_parts.append(t.text)

    # Remove duplicates while preserving order
    seen = set()
    unique_citations = []
    for c in citations:
        if c not in seen:
            seen.add(c)
            unique_citations.append(c)

    return ''.join(text_parts), unique_citations


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

        # Numbered list of citations
        for j, cite in enumerate(citations, 1):
            lines.append(f"- {j}. Citation {cite}")

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
