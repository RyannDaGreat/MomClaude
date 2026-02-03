#!/usr/bin/env python3
"""
Extract paragraphs and their citations from a Word document.

Output format:
A markdown file with two sections:
1. Citation Extraction by Table - citations in each table (row-first order)
2. Citation Extraction by Paragraph - paragraphs with their citations

Includes both numeric citations (superscript) and Table references (Roman numerals).
"""

import zipfile
from xml.etree import ElementTree as ET
import re
from pathlib import Path


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

ROMAN_PATTERN = re.compile(r'\b[Tt]ables?\s+([IVXivx]+)\b')

ROMAN_TO_INT = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10}
INT_TO_ROMAN = {v: k for k, v in ROMAN_TO_INT.items()}


def extract_citations_from_element(elem):
    """Extract superscript citation numbers from an element."""
    citations = []
    for r in elem.findall('.//w:r', NAMESPACES):
        rPr = r.find('w:rPr', NAMESPACES)
        is_sup = False
        if rPr is not None:
            vertAlign = rPr.find('w:vertAlign', NAMESPACES)
            if vertAlign is not None:
                val = vertAlign.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if val == 'superscript':
                    is_sup = True
        t = r.find('w:t', NAMESPACES)
        if t is not None and t.text and is_sup:
            nums = re.findall(r'\d+', t.text)
            citations.extend(nums)
    return citations


def extract_table_citations(tbl):
    """
    Extract citations from a table, reading row-first.

    Returns list of citation numbers in order of appearance (row by row).
    """
    citations = []
    rows = tbl.findall('.//w:tr', NAMESPACES)

    for row in rows:
        cells = row.findall('.//w:tc', NAMESPACES)
        for cell in cells:
            cell_citations = extract_citations_from_element(cell)
            citations.extend(cell_citations)

    # Remove duplicates while preserving order
    seen = set()
    unique = []
    for c in citations:
        if c not in seen:
            seen.add(c)
            unique.append(c)

    return unique


def process_tables(root):
    """
    Process document to extract tables with their Roman numeral labels.

    Returns dict: {roman_numeral: [citations]}
    """
    body = root.find('.//w:body', NAMESPACES)

    # Find table labels and group physical tables
    table_groups = {}  # roman_numeral -> list of table elements
    current_label = None

    for elem in body:
        tag = elem.tag.replace('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}', '')

        if tag == 'p':
            texts = []
            for t in elem.findall('.//w:t', NAMESPACES):
                if t.text:
                    texts.append(t.text)
            text = ''.join(texts).strip()

            # Check for table label (e.g., "TABLE I." or "TABLE II.")
            match = re.search(r'\bTABLE\s+([IVX]+)\b', text.upper())
            if match:
                roman = match.group(1)
                current_label = roman
                if roman not in table_groups:
                    table_groups[roman] = []

        elif tag == 'tbl' and current_label:
            table_groups[current_label].append(elem)

    # Extract citations from each table group
    result = {}
    for roman, tables in table_groups.items():
        all_citations = []
        for tbl in tables:
            citations = extract_table_citations(tbl)
            all_citations.extend(citations)

        # Remove duplicates while preserving order
        seen = set()
        unique = []
        for c in all_citations:
            if c not in seen:
                seen.add(c)
                unique.append(c)

        result[roman] = unique

    return result


def extract_paragraph_with_citations(p):
    """Extract text and citation references from a paragraph element."""
    runs = []
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

    text_parts = [text for text, is_sup, _ in runs if not is_sup]
    full_text = ''.join(text_parts)

    citations = []

    for text, is_sup, char_pos in runs:
        if is_sup:
            nums = re.findall(r'\d+', text)
            for num in nums:
                citations.append((char_pos, f"Citation {num}"))

    full_non_sup = ''.join(text for text, is_sup, _ in runs if not is_sup)

    non_sup_pos = 0
    char_to_orig_pos = {}
    for text, is_sup, orig_pos in runs:
        if not is_sup:
            for i, ch in enumerate(text):
                char_to_orig_pos[non_sup_pos + i] = orig_pos + i
            non_sup_pos += len(text)

    for match in ROMAN_PATTERN.finditer(full_non_sup):
        roman = match.group(1).upper()
        match_pos = match.start()
        orig_pos = char_to_orig_pos.get(match_pos, match_pos)
        citations.append((orig_pos, f"Table {roman}"))

    citations.sort(key=lambda x: x[0])

    seen = set()
    unique_citations = []
    for pos, cite in citations:
        if cite not in seen:
            seen.add(cite)
            unique_citations.append(cite)

    return full_text, unique_citations


def process_paragraphs(root):
    """Process paragraphs and extract citations."""
    results = []
    for p in root.findall('.//w:p', NAMESPACES):
        text, citations = extract_paragraph_with_citations(p)
        text = text.strip()

        if not text or len(text) < 30:
            continue

        if not citations:
            continue

        words = text.split()
        first_three = ' '.join(words[:3])

        results.append((first_three, text, citations))

    return results


def generate_markdown(table_citations, paragraph_results, output_path):
    """Generate markdown output file with both sections."""
    lines = []

    # Section 1: Citation Extraction by Table
    lines.append("# Section 1: Citation Extraction by Table\n")

    # Sort tables by Roman numeral value
    sorted_tables = sorted(table_citations.items(), key=lambda x: ROMAN_TO_INT.get(x[0], 99))

    for roman, citations in sorted_tables:
        lines.append(f"## Table {roman}")
        lines.append("")

        for j, cite in enumerate(citations, 1):
            lines.append(f"- {j}. Citation {cite}")

        lines.append("")

    # Section 2: Citation Extraction by Paragraph
    lines.append("# Section 2: Citation Extraction by Paragraph\n")

    for i, (first_three, full_text, citations) in enumerate(paragraph_results, 1):
        lines.append(f"## {i}. {first_three}...")
        lines.append("")

        for j, cite in enumerate(citations, 1):
            lines.append(f"- {j}. {cite}")

        lines.append("")

    output_path.write_text('\n'.join(lines))
    return len(sorted_tables), len(paragraph_results)


def main():
    docx_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/JAIP 6391.docx')
    output_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/citations_by_paragraph.md')

    print(f"Processing: {docx_path}")

    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

    table_citations = process_tables(root)
    paragraph_results = process_paragraphs(root)

    table_count, para_count = generate_markdown(table_citations, paragraph_results, output_path)
    print(f"Extracted {table_count} tables with citations")
    print(f"Extracted {para_count} paragraphs with citations")
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
