#!/usr/bin/env python3
"""
Extract paragraphs and their citations from a Word document.

Output format:
A markdown file with two sections:
1. Citation Extraction by Table - citations in each table (row-first order)
2. Citation Extraction by Paragraph - paragraphs with their citations

=== CITATION EXTRACTION RULES ===

Section 1: Table Citation Extraction Rules

1. Table Identification: Tables are identified by their Roman numeral labels
   (e.g., "TABLE I", "TABLE II") in paragraph text preceding the table elements.
   Uses is_roman_numeral() function to validate (currently permissive, can be
   swapped for strict validation later).

2. Table Grouping: Multiple physical <w:tbl> elements following a single TABLE
   label are treated as one logical table (e.g., TABLE I may span 2 physical tables).

3. Row-First Reading Order: Citations within tables are extracted row-by-row,
   left-to-right within each row (not column-first).

4. Table References in Paragraphs: When "Table I" or "Table II" etc. appears in
   paragraph text, it's recorded as a reference in its position relative to other
   citations.

Section 2: Paragraph Citation Extraction Rules

5. Left vs Right Position Rule:
   - Superscript numbers appearing AFTER text (right side) = citations -> include
   - Superscript numbers appearing BEFORE text (left side) = chemical/isotope -> skip

6. Range Detection Rule:
   - Superscripts containing hyphen between numbers (e.g., 197-199) = range
   - Output as "Citations 197-199" not individual citations

7. Citation Format Whitelist: Only accept superscripts matching \\d+ or \\d+-\\d+
   patterns (implicitly excludes author affiliations like a, b, c).

8. No Deduplication: If a citation appears multiple times, include it multiple
   times in order of appearance.
"""

import zipfile
from xml.etree import ElementTree as ET
import re
from pathlib import Path


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

ROMAN_PATTERN = re.compile(r'\b[Tt]ables?\s+([IVXLCDMivxlcdm]+)\b')

# Citation patterns: single number or range
CITATION_SINGLE = re.compile(r'^(\d+)$')
CITATION_RANGE = re.compile(r'^(\d+)-(\d+)$')

ROMAN_TO_INT = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10}


def is_roman_numeral(s):
    """
    Check if string is a valid Roman numeral.

    Currently permissive - just checks if only contains valid Roman numeral chars.
    Can be swapped for strict validation later.
    """
    return bool(s) and all(c in 'IVXLCDMivxlcdm' for c in s)


def parse_citation(text):
    """
    Parse superscript text into a citation string.

    Returns list with one citation string, or empty list if invalid.
    Input like "19-22,42" -> ["Citations 19-22, 42"]
    Input like "42" -> ["Citation 42"]
    """
    text = text.strip()
    parts = [p.strip() for p in text.split(',') if p.strip()]
    valid = [p for p in parts if CITATION_SINGLE.match(p) or CITATION_RANGE.match(p)]

    if not valid:
        return []

    combined = ', '.join(valid)
    plural = len(valid) > 1 or '-' in combined
    return [f"Citation{'s' if plural else ''} {combined}"]


def extract_citations_from_runs(runs):
    """
    Extract citations from a list of (text, is_superscript) tuples.

    Applies left vs right position rule: only include superscripts that appear
    AFTER regular text (right side), not before (left side = chemical notation).

    Returns list of citation strings in order of appearance.
    """
    citations = []
    has_seen_regular_text = False
    i = 0

    while i < len(runs):
        text, is_sup = runs[i]

        if not is_sup:
            if text.strip():
                has_seen_regular_text = True
            i += 1
        else:
            # Superscript - combine consecutive superscript runs
            combined_sup = text
            j = i + 1
            while j < len(runs) and runs[j][1]:  # while next run is also superscript
                combined_sup += runs[j][0]
                j += 1

            # Look at what comes immediately after all these superscripts
            next_text = ''
            if j < len(runs):
                next_text = runs[j][0]

            # Left vs Right position rule for chemical/isotope detection:
            # If superscript is on the LEFT side of text (immediately followed by 1-2
            # letter element symbol like "Xe", "He", "F"), it's isotope notation -> skip
            # Pattern: superscript followed by 1-2 letters then non-letter or end
            if next_text:
                # Check for element-like pattern: 1-2 letters followed by non-letter
                element_match = re.match(r'^([A-Z][a-z]?)(?:[^a-zA-Z]|$)', next_text)
                if element_match and len(element_match.group(1)) <= 2:
                    # This looks like isotope notation (e.g., 129Xe, 3He, 19F)
                    i = j
                    continue

            # Only include if we've seen regular text before this (right-side rule)
            if has_seen_regular_text:
                parsed = parse_citation(combined_sup)
                citations.extend(parsed)

            i = j

    return citations


def extract_table_citations(tbl):
    """
    Extract citations from a table, reading row-first.

    Returns list of citation strings in order of appearance (row by row).
    """
    citations = []
    rows = tbl.findall('.//w:tr', NAMESPACES)

    for row in rows:
        cells = row.findall('.//w:tc', NAMESPACES)
        for cell in cells:
            # Build runs for this cell
            runs = []
            for r in cell.findall('.//w:r', NAMESPACES):
                rPr = r.find('w:rPr', NAMESPACES)
                is_sup = False
                if rPr is not None:
                    vertAlign = rPr.find('w:vertAlign', NAMESPACES)
                    if vertAlign is not None:
                        val = vertAlign.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        if val == 'superscript':
                            is_sup = True
                t = r.find('w:t', NAMESPACES)
                if t is not None and t.text:
                    runs.append((t.text, is_sup))

            cell_citations = extract_citations_from_runs(runs)
            citations.extend(cell_citations)

    return citations


def process_tables(root):
    """
    Process document to extract tables with their Roman numeral labels.

    Returns dict: {roman_numeral: [citations]}
    """
    body = root.find('.//w:body', NAMESPACES)

    table_groups = {}
    current_label = None

    for elem in body:
        tag = elem.tag.replace('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}', '')

        if tag == 'p':
            texts = []
            for t in elem.findall('.//w:t', NAMESPACES):
                if t.text:
                    texts.append(t.text)
            text = ''.join(texts).strip()

            match = re.search(r'\bTABLE\s+([IVXLCDMivxlcdm]+)\b', text.upper())
            if match:
                roman = match.group(1).upper()
                if is_roman_numeral(roman):
                    current_label = roman
                    if roman not in table_groups:
                        table_groups[roman] = []

        elif tag == 'tbl' and current_label:
            table_groups[current_label].append(elem)

    result = {}
    for roman, tables in table_groups.items():
        all_citations = []
        for tbl in tables:
            citations = extract_table_citations(tbl)
            all_citations.extend(citations)
        result[roman] = all_citations

    return result


def extract_paragraph_with_citations(p):
    """Extract text and citation references from a paragraph element."""
    runs = []

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
            runs.append((t.text, is_superscript))

    # Build full text (non-superscript only)
    text_parts = [text for text, is_sup in runs if not is_sup]
    full_text = ''.join(text_parts)

    # Extract citations using left/right position rule
    citations = extract_citations_from_runs(runs)

    # Also extract Table references from the text
    for match in ROMAN_PATTERN.finditer(full_text):
        roman = match.group(1).upper()
        if is_roman_numeral(roman):
            citations.append(f"Table {roman}")

    return full_text, citations


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

    sorted_tables = sorted(table_citations.items(), key=lambda x: ROMAN_TO_INT.get(x[0], 99))

    for roman, citations in sorted_tables:
        lines.append(f"## Table {roman}")
        lines.append("")

        for j, cite in enumerate(citations, 1):
            lines.append(f"- {j}. {cite}")

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
