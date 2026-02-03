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
from functools import lru_cache
from difflib import SequenceMatcher


@lru_cache(maxsize=4)
def load_docx_xml(docx_path: str) -> str:
    """
    Not a pure function. Load and cache document.xml from a .docx file.

    Returns the raw XML string for word/document.xml.

    >>> # Can't doctest file I/O, but usage is:
    >>> # xml_str = load_docx_xml('/path/to/doc.docx')
    """
    with zipfile.ZipFile(docx_path) as z:
        return z.read('word/document.xml').decode('utf-8')


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

# === REGEX PATTERNS ===
ROMAN_PATTERN = re.compile(r'\b[Tt]ables?\s+([IVXLCDMivxlcdm]+)\b')  # Table references in text
CITATION_SINGLE = re.compile(r'^(\d+)$')                             # Single citation number
CITATION_RANGE = re.compile(r'^(\d+)-(\d+)$')                        # Citation range (e.g., 19-22)
AUTHOR_PATTERN = re.compile(r'([A-Za-zÀ-ÿ\-\']+)\s+[A-Z]{1,3}(?:[,;.]|$)')  # Author names in refs
REF_ENTRY_PATTERN = re.compile(r'^\d+\.\s+')                         # Reference entry (e.g., "1. ")

# === CONSTANTS ===
ROMAN_TO_INT = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10}
DUPLICATE_THRESHOLD = 0.90  # Similarity threshold for duplicate detection


def detect_duplicate_references(references: dict, threshold: float = DUPLICATE_THRESHOLD) -> dict:
    """
    Pure function. Detect duplicate references by text similarity.

    Takes dict of {citation_number: reference_text} and returns dict mapping
    duplicate citation numbers to the original (lower) citation number they
    should be merged into.

    Uses SequenceMatcher ratio - pairs with similarity >= threshold are duplicates.

    >>> refs = {1: "Smith J. Paper title. Journal 2020;1:1-10.",
    ...         2: "Jones K. Different paper. Other J 2021;2:20-30.",
    ...         3: "Smith J. Paper title. Journal 2020;1:1-10."}
    >>> detect_duplicate_references(refs)
    {3: 1}
    >>> refs2 = {1: "Smith J. Paper. J 2020.", 2: "Smith J. Paper. J 2020.", 3: "Smith J. Paper. J 2020."}
    >>> detect_duplicate_references(refs2)
    {2: 1, 3: 1}
    >>> refs3 = {1: "Completely different", 2: "Nothing alike"}
    >>> detect_duplicate_references(refs3)
    {}
    >>> detect_duplicate_references({})
    {}
    """
    duplicates = {}
    nums = sorted(references.keys())

    for i, n1 in enumerate(nums):
        # Skip if n1 is already marked as a duplicate
        if n1 in duplicates:
            continue

        for n2 in nums[i+1:]:
            # Skip if n2 is already marked as a duplicate
            if n2 in duplicates:
                continue

            ratio = SequenceMatcher(None, references[n1], references[n2]).ratio()
            if ratio >= threshold:
                # n2 is duplicate of n1 (n1 is lower, so it's the original)
                duplicates[n2] = n1

    return duplicates


def extract_numbers_from_citation(cite_str: str) -> list:
    """
    Pure function. Extract individual citation numbers from a citation string.

    Handles single citations, ranges, and comma-separated lists.
    Returns list of ints in order of appearance.

    >>> extract_numbers_from_citation('Citation 42')
    [42]
    >>> extract_numbers_from_citation('Citations 1, 2')
    [1, 2]
    >>> extract_numbers_from_citation('Citations 19-22')
    [19, 20, 21, 22]
    >>> extract_numbers_from_citation('Citations 19-22, 42')
    [19, 20, 21, 22, 42]
    >>> extract_numbers_from_citation('Table I')
    []
    >>> extract_numbers_from_citation('Citations 1, 2, 3')
    [1, 2, 3]
    """
    if cite_str.startswith('Table'):
        return []

    numbers = []
    # Remove "Citation" or "Citations" prefix
    text = re.sub(r'^Citations?\s+', '', cite_str)

    # Split by comma
    parts = [p.strip() for p in text.split(',')]

    for part in parts:
        range_match = re.match(r'^(\d+)-(\d+)$', part)
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            numbers.extend(range(start, end + 1))
        elif re.match(r'^\d+$', part):
            numbers.append(int(part))

    return numbers


def build_canonical_order(table_citations: dict, paragraph_results: list) -> list:
    """
    Pure function. Build canonical citation order by walking through paragraphs.

    When a Table reference is encountered, its citations are expanded inline.
    Returns list of citation numbers in order of first appearance.

    >>> tables = {'I': ['Citation 9', 'Citation 10']}
    >>> paras = [('First para', 'text', ['Citation 1', 'Table I', 'Citation 2'])]
    >>> build_canonical_order(tables, paras)
    [1, 9, 10, 2]
    >>> tables2 = {}
    >>> paras2 = [('P1', 't', ['Citation 5', 'Citation 3']), ('P2', 't', ['Citation 3', 'Citation 7'])]
    >>> build_canonical_order(tables2, paras2)
    [5, 3, 7]
    >>> build_canonical_order({}, [])
    []
    """
    seen = set()
    order = []

    for first_three, full_text, citations in paragraph_results:
        for cite in citations:
            if cite.startswith('Table'):
                # Extract roman numeral and expand table citations
                match = re.match(r'Table\s+([IVXLCDMivxlcdm]+)', cite)
                if match:
                    roman = match.group(1).upper()
                    if roman in table_citations:
                        for table_cite in table_citations[roman]:
                            for num in extract_numbers_from_citation(table_cite):
                                if num not in seen:
                                    seen.add(num)
                                    order.append(num)
            else:
                for num in extract_numbers_from_citation(cite):
                    if num not in seen:
                        seen.add(num)
                        order.append(num)

    return order


def build_conversion_table(canonical_order: list) -> dict:
    """
    Pure function. Build conversion table from current to canonical numbers.

    Returns dict mapping old_number -> new_number, only for numbers that change.

    >>> build_conversion_table([1, 2, 3])
    {}
    >>> build_conversion_table([5, 3, 7])
    {5: 1, 3: 2, 7: 3}
    >>> build_conversion_table([1, 9, 10, 2])
    {9: 2, 10: 3, 2: 4}
    >>> build_conversion_table([])
    {}
    """
    conversion = {}
    for new_num, old_num in enumerate(canonical_order, 1):
        if old_num != new_num:
            conversion[old_num] = new_num
    return conversion


def extract_author_names(doc_xml: str) -> set:
    """
    Pure function. Extract author last names from reference entries in document XML.

    Parses reference entries (lines starting with "N. ") and extracts author names
    using the pattern "LastName Initials," format common in academic citations.

    Returns set of author last names (preserving original case).

    >>> xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>1. Smith AB, Jones CD. Title here.</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> sorted(extract_author_names(xml))
    ['Jones', 'Smith']
    >>> xml2 = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>Regular paragraph text.</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> extract_author_names(xml2)
    set()
    >>> xml3 = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ...   <w:body><w:p><w:r><w:t>1. Li X, Du Y, Ho Z. Study.</w:t></w:r></w:p></w:body>
    ... </w:document>'''
    >>> sorted(extract_author_names(xml3))
    ['Du', 'Ho', 'Li']
    """
    root = ET.fromstring(doc_xml)
    author_names = set()

    for p in root.findall('.//w:p', NAMESPACES):
        text = ''.join(t.text for t in p.findall('.//w:t', NAMESPACES) if t.text)
        if not REF_ENTRY_PATTERN.match(text):
            continue

        # Remove leading number
        text = REF_ENTRY_PATTERN.sub('', text)

        # Get author section (before title - first ". " followed by capital letter)
        parts = re.split(r'\.\s+(?=[A-Z])', text, maxsplit=1)
        if not parts:
            continue

        author_section = parts[0]
        names = AUTHOR_PATTERN.findall(author_section)

        for name in names:
            if name.lower() not in ('et', 'al'):
                author_names.add(name)

    return author_names


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


def extract_citations_from_runs(runs: list, author_names: set = None) -> list:
    """
    Pure function. Extract citations from a list of (str, bool) tuples.

    Each tuple is (text, is_superscript). Applies left vs right position rule:
    only include superscripts that appear AFTER regular text (right side),
    not before (left side = chemical notation). Consecutive superscript runs
    are combined before parsing.

    Optional author_names set whitelists short names that should not be treated
    as isotope symbols (e.g., "Li", "Ho" are real author names, not elements).

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
    >>> extract_citations_from_runs([('Text', False), ('42', True), ('Li et al', False)], {'Li'})
    ['Citation 42']
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
            if is_left_side_superscript(next_text, author_names):
                i = j
                continue

            # Only include if we've seen regular text before this (right-side rule)
            if has_seen_regular_text:
                parsed = parse_citation(combined_sup)
                citations.extend(parsed)

            i = j

    return citations


def is_left_side_superscript(next_text: str, author_names: set = None) -> bool:
    """
    Pure function. Check if superscript is on the LEFT side of text (not a citation).

    Isotope notation (129Xe, 3He, 19F) has superscript numbers directly attached to
    short element symbols (1-2 letters). Citations are followed by space, punctuation,
    end of text, or longer words (author names like "Pavord").

    Rule: If next_text starts with 1-2 letters followed by non-letter (space, punct,
    or end), treat as left-side (isotope) UNLESS that word is a known author name.

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
    >>> is_left_side_superscript('Li et al', {'Li', 'Du'})
    False
    >>> is_left_side_superscript('Xe et al', {'Xe'})
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

    # If it's a known author name, it's not isotope notation
    if author_names and first_word in author_names:
        return False

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


def extract_table_citations(xml_str: str, author_names: set = None) -> list:
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
            cell_citations = extract_citations_from_runs(runs, author_names)
            citations.extend(cell_citations)

    return citations


def extract_paragraph_with_citations(xml_str: str, author_names: set = None) -> tuple:
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
    citations = extract_citations_from_runs(runs, author_names)

    # Also extract Table references from the text
    for match in ROMAN_PATTERN.finditer(full_text):
        roman = match.group(1).upper()
        if is_roman_numeral(roman):
            citations.append(f"Table {roman}")

    return full_text, citations


def process_tables(xml_str: str, author_names: set = None) -> dict:
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
            citations = extract_table_citations(tbl_xml, author_names)
            all_citations.extend(citations)
        result[roman] = all_citations

    return result


def process_paragraphs(xml_str: str, author_names: set = None) -> list:
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
        text, citations = extract_paragraph_with_citations(p_xml, author_names)
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

    Generate markdown output file with three sections.
    Returns tuple of (int, int, int): (table_count, paragraph_count, conversion_count).
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

    # Section 2: Citation Extraction by Paragraph (with Table expansion)
    lines.append("# Section 2: Citation Extraction by Paragraph\n")

    cite_counter = 0
    for i, (first_three, full_text, citations) in enumerate(paragraph_results, 1):
        lines.append(f"## {i}. {first_three}...")
        lines.append("")

        j = 0
        for cite in citations:
            j += 1
            if cite.startswith('Table'):
                # Expand table reference with sub-bullets
                lines.append(f"- {j}. {cite}")
                match = re.match(r'Table\s+([IVXLCDMivxlcdm]+)', cite)
                if match:
                    roman = match.group(1).upper()
                    if roman in table_citations:
                        for k, table_cite in enumerate(table_citations[roman], 1):
                            cite_counter += 1
                            lines.append(f"  - {j}.{k}. {table_cite}")
            else:
                cite_counter += 1
                lines.append(f"- {j}. {cite}")

        lines.append("")

    # Section 3: Citation Conversion Table
    canonical_order = build_canonical_order(table_citations, paragraph_results)
    conversion = build_conversion_table(canonical_order)

    lines.append("# Section 3: Citation Conversion Table\n")

    if conversion:
        lines.append("| From | To |")
        lines.append("|------|-----|")

        for old_num in sorted(conversion.keys()):
            new_num = conversion[old_num]
            lines.append(f"| {old_num} | {new_num} |")
    else:
        lines.append("No conversions needed - all citations are already in canonical order.")

    lines.append("")

    Path(output_path).write_text('\n'.join(lines))
    return len(sorted_tables), len(paragraph_results), len(conversion)


def main():
    docx_path = '/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/JAIP 6391.docx'
    output_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/citations_by_paragraph.md')

    print(f"Processing: {docx_path}")

    xml_str = load_docx_xml(docx_path)
    author_names = extract_author_names(xml_str)
    print(f"Found {len(author_names)} author names in references")

    table_citations = process_tables(xml_str, author_names)
    paragraph_results = process_paragraphs(xml_str, author_names)

    table_count, para_count, conv_count = generate_markdown(table_citations, paragraph_results, str(output_path))
    print(f"Extracted {table_count} tables with citations")
    print(f"Extracted {para_count} paragraphs with citations")
    print(f"Citations needing renumbering: {conv_count}")
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
