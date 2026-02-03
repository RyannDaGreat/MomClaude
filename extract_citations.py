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

9. Reference List Filtering: Bibliography/reference entries (lines starting with
   a number followed by a period, e.g., "197. Kaireit TF...") are excluded from
   paragraph extraction.

=== CODE CONVENTIONS ===

- Pure functions should be marked with "Pure function." at the start of their
  docstring. Pure functions take simple built-in types (str, int, list, tuple,
  dict) and return deterministic output with no side effects.

- Non-pure functions (those with side effects like file I/O) should be marked
  with "Not a pure function." at the start of their docstring.

- All pure functions must have at least 3 informative doctest examples in
  ">>> ..." format so the function's behavior is clear without reading the code.

- This docstring must be kept in sync with the code. When adding new rules or
  conventions, update this docstring to reflect those changes.
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


def is_roman_numeral(s: str) -> bool:
    """
    Pure function. Check if string contains only valid Roman numeral characters.

    Currently permissive - can be swapped for strict validation later.

    >>> is_roman_numeral('I')
    True
    >>> is_roman_numeral('IV')
    True
    >>> is_roman_numeral('XVII')
    True
    >>> is_roman_numeral('')
    False
    >>> is_roman_numeral('ABC')
    False
    >>> is_roman_numeral('I2')
    False
    """
    return bool(s) and all(c in 'IVXLCDMivxlcdm' for c in s)


def parse_citation(text: str) -> list:
    """
    Pure function. Parse superscript text into a citation string.

    Returns list[str] with one citation string, or empty list if invalid.
    Uses "Citation" for single numbers, "Citations" for ranges or multiple.

    >>> parse_citation('42')
    ['Citation 42']
    >>> parse_citation('19-22')
    ['Citations 19-22']
    >>> parse_citation('19-22,42')
    ['Citations 19-22, 42']
    >>> parse_citation('1,2,3')
    ['Citations 1, 2, 3']
    >>> parse_citation('abc')
    []
    >>> parse_citation('')
    []
    """
    text = text.strip()
    parts = [p.strip() for p in text.split(',') if p.strip()]
    valid = [p for p in parts if CITATION_SINGLE.match(p) or CITATION_RANGE.match(p)]

    if not valid:
        return []

    combined = ', '.join(valid)
    plural = len(valid) > 1 or '-' in combined
    return [f"Citation{'s' if plural else ''} {combined}"]


def extract_citations_from_runs(runs: list) -> list:
    """
    Pure function. Extract citations from a list of (str, bool) tuples.

    Each tuple is (text, is_superscript). Applies left vs right position rule:
    only include superscripts that appear AFTER regular text (right side),
    not before (left side = chemical notation). Consecutive superscript runs
    are combined before parsing.

    Returns list[str] of citation strings.

    >>> extract_citations_from_runs([('Hello.', False), ('42', True)])
    ['Citation 42']
    >>> extract_citations_from_runs([('Text.', False), ('1', True), (',', True), ('2', True)])
    ['Citations 1, 2']
    >>> extract_citations_from_runs([('129', True), ('Xe', False)])
    []
    >>> extract_citations_from_runs([('Study by Smith', False), ('10', True), (' Jones', False), ('11', True)])
    ['Citation 10', 'Citation 11']
    >>> extract_citations_from_runs([('a', True), ('Text', False)])
    []
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

            # Left vs Right position rule:
            # If superscript is on the LEFT side of text (followed by short word), skip it
            if is_left_side_superscript(next_text):
                i = j
                continue

            # Only include if we've seen regular text before this (right-side rule)
            if has_seen_regular_text:
                parsed = parse_citation(combined_sup)
                citations.extend(parsed)

            i = j

    return citations


def is_left_side_superscript(next_text: str) -> bool:
    """
    Pure function. Check if superscript is on the LEFT side of text (not a citation).

    Isotope notation (129Xe, 3He, 19F) has superscript numbers directly attached to
    short element symbols (1-2 letters). Citations are followed by space, punctuation,
    end of text, or longer words (author names like "Pavord").

    Rule: If next_text starts with 1-2 letters followed by non-letter (space, punct,
    or end), treat as left-side (isotope). Otherwise it's a citation.

    >>> is_left_side_superscript('Xe MRI')
    True
    >>> is_left_side_superscript('He)')
    True
    >>> is_left_side_superscript('F MRI')
    True
    >>> is_left_side_superscript('x')
    True
    >>> is_left_side_superscript('')
    False
    >>> is_left_side_superscript(' some text')
    False
    >>> is_left_side_superscript(', more text')
    False
    >>> is_left_side_superscript('Pavord et al')
    False
    >>> is_left_side_superscript('Jones et al')
    False
    >>> is_left_side_superscript('Götschke et al')
    False
    """
    if not next_text:
        return False

    # Check if starts with a letter
    if not next_text[0].isalpha():
        return False

    # Find the first word (consecutive letters)
    first_word = ''
    for c in next_text:
        if c.isalpha():
            first_word += c
        else:
            break

    # If first word is 1-2 letters, treat as isotope notation (left side)
    return len(first_word) <= 2


def is_reference_entry(text: str) -> bool:
    """
    Pure function. Check if text is a bibliography/reference list entry.

    Reference entries typically start with a number followed by a period.

    >>> is_reference_entry('197. Kaireit TF, Kern A...')
    True
    >>> is_reference_entry('1. First reference entry')
    True
    >>> is_reference_entry('Eosinophils are granulocytic cells...')
    False
    >>> is_reference_entry('The study by Smith et al.')
    False
    """
    return bool(re.match(r'^\d+\.\s', text.strip()))


def xml_to_runs(xml_str: str) -> list:
    """
    Pure function. Convert XML string to list of (str, bool) tuples.

    Parses Word XML and extracts text runs with superscript information.
    Returns list of (text, is_superscript) tuples.

    >>> xml = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>Hello.</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>42</w:t></w:r>
    ... </w:p>'''
    >>> xml_to_runs(xml)
    [('Hello.', False), ('42', True)]
    >>> xml_empty = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>No superscripts here.</w:t></w:r>
    ... </w:p>'''
    >>> xml_to_runs(xml_empty)
    [('No superscripts here.', False)]
    >>> xml_multi = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>A</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
    ...   <w:r><w:t>B</w:t></w:r>
    ... </w:p>'''
    >>> xml_to_runs(xml_multi)
    [('A', False), ('1', True), ('B', False)]
    """
    elem = ET.fromstring(xml_str)
    runs = []

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
        if t is not None and t.text:
            runs.append((t.text, is_sup))

    return runs


def extract_table_citations(xml_str: str) -> list:
    """
    Pure function. Extract citations from table XML string.

    Reads row-first, returns list[str] of citation strings in order of appearance.

    >>> xml = '''<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:tr><w:tc><w:r><w:t>Text</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>42</w:t></w:r></w:tc></w:tr>
    ... </w:tbl>'''
    >>> extract_table_citations(xml)
    ['Citation 42']
    >>> xml_empty = '''<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:tr><w:tc><w:r><w:t>No citations here</w:t></w:r></w:tc></w:tr>
    ... </w:tbl>'''
    >>> extract_table_citations(xml_empty)
    []
    >>> xml_multi = '''<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:tr><w:tc><w:r><w:t>A</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r></w:tc>
    ...   <w:tc><w:r><w:t>B</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:tc></w:tr>
    ... </w:tbl>'''
    >>> extract_table_citations(xml_multi)
    ['Citation 1', 'Citation 2']
    """
    tbl = ET.fromstring(xml_str)
    citations = []
    rows = tbl.findall('.//w:tr', NAMESPACES)

    for row in rows:
        cells = row.findall('.//w:tc', NAMESPACES)
        for cell in cells:
            cell_xml = ET.tostring(cell, encoding='unicode')
            runs = xml_to_runs(cell_xml)
            cell_citations = extract_citations_from_runs(runs)
            citations.extend(cell_citations)

    return citations


def extract_paragraph_with_citations(xml_str: str) -> tuple:
    """
    Pure function. Extract text and citations from paragraph XML string.

    Returns tuple of (str, list[str]): (full_text, citation_strings).

    >>> xml = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>Hello world.</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>42</w:t></w:r>
    ... </w:p>'''
    >>> extract_paragraph_with_citations(xml)
    ('Hello world.', ['Citation 42'])
    >>> xml_empty = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>No citations here.</w:t></w:r>
    ... </w:p>'''
    >>> extract_paragraph_with_citations(xml_empty)
    ('No citations here.', [])
    >>> xml_table = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:r><w:t>See Table I for details.</w:t></w:r>
    ... </w:p>'''
    >>> extract_paragraph_with_citations(xml_table)
    ('See Table I for details.', ['Table I'])
    """
    runs = xml_to_runs(xml_str)

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


def process_tables(xml_str: str) -> dict:
    """
    Pure function. Extract tables with their Roman numeral labels from document XML.

    Returns dict[str, list[str]]: {roman_numeral: [citation_strings]}

    >>> xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body>
    ...     <w:p><w:r><w:t>TABLE I. Title</w:t></w:r></w:p>
    ...     <w:tbl><w:tr><w:tc><w:r><w:t>Data</w:t></w:r>
    ...     <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r></w:tc></w:tr></w:tbl>
    ...   </w:body>
    ... </w:document>'''
    >>> process_tables(xml)
    {'I': ['Citation 1']}
    >>> xml_empty = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>No tables here</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> process_tables(xml_empty)
    {}
    >>> xml_two = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body>
    ...     <w:p><w:r><w:t>TABLE I.</w:t></w:r></w:p>
    ...     <w:tbl><w:tr><w:tc><w:r><w:t>A</w:t></w:r></w:tc></w:tr></w:tbl>
    ...     <w:p><w:r><w:t>TABLE II.</w:t></w:r></w:p>
    ...     <w:tbl><w:tr><w:tc><w:r><w:t>B</w:t></w:r></w:tc></w:tr></w:tbl>
    ...   </w:body>
    ... </w:document>'''
    >>> sorted(process_tables(xml_two).keys())
    ['I', 'II']
    """
    root = ET.fromstring(xml_str)
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
            tbl_xml = ET.tostring(elem, encoding='unicode')
            table_groups[current_label].append(tbl_xml)

    result = {}
    for roman, table_xmls in table_groups.items():
        all_citations = []
        for tbl_xml in table_xmls:
            citations = extract_table_citations(tbl_xml)
            all_citations.extend(citations)
        result[roman] = all_citations

    return result


def process_paragraphs(xml_str: str) -> list:
    """
    Pure function. Extract paragraphs with citations from document XML.

    Skips short paragraphs (<30 chars), paragraphs without citations, and
    reference list entries. Returns list of tuples:
    (first_three_words: str, full_text: str, citations: list[str])

    >>> xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body>
    ...     <w:p><w:r><w:t>This is a paragraph with enough text to pass the length filter.</w:t></w:r>
    ...     <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r></w:p>
    ...   </w:body>
    ... </w:document>'''
    >>> result = process_paragraphs(xml)
    >>> len(result)
    1
    >>> result[0][0]
    'This is a'
    >>> xml_short = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>Too short.</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> process_paragraphs(xml_short)
    []
    >>> xml_ref = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>1. This is a reference entry that should be skipped entirely.</w:t></w:r>
    ...   <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>99</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> process_paragraphs(xml_ref)
    []
    """
    root = ET.fromstring(xml_str)
    results = []

    for p in root.findall('.//w:p', NAMESPACES):
        p_xml = ET.tostring(p, encoding='unicode')
        text, citations = extract_paragraph_with_citations(p_xml)
        text = text.strip()

        if not text or len(text) < 30:
            continue

        if not citations:
            continue

        # Skip reference list entries (e.g., "197. Kaireit TF...")
        if is_reference_entry(text):
            continue

        words = text.split()
        first_three = ' '.join(words[:3])

        results.append((first_three, text, citations))

    return results


def generate_markdown(table_citations: dict, paragraph_results: list, output_path: str) -> tuple:
    """
    Not a pure function. Writes to file.

    Generate markdown output file with both sections.
    Returns tuple of (int, int): (table_count, paragraph_count).
    """
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

    Path(output_path).write_text('\n'.join(lines))
    return len(sorted_tables), len(paragraph_results)


def main():
    docx_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/JAIP 6391.docx')
    output_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/citations_by_paragraph.md')

    print(f"Processing: {docx_path}")

    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            xml_str = f.read().decode('utf-8')

    table_citations = process_tables(xml_str)
    paragraph_results = process_paragraphs(xml_str)

    table_count, para_count = generate_markdown(table_citations, paragraph_results, str(output_path))
    print(f"Extracted {table_count} tables with citations")
    print(f"Extracted {para_count} paragraphs with citations")
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
