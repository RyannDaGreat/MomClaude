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


def extract_references(docx_path: str) -> dict:
    """
    Not a pure function. Extract all references from a .docx file.

    Returns dict of {citation_number: reference_text}.
    References are identified as paragraphs starting with "N. " pattern.
    """
    xml_str = load_docx_xml(docx_path)
    root = ET.fromstring(xml_str)

    refs = {}
    for p in root.findall('.//w:p', NAMESPACES):
        text = ''.join(t.text for t in p.findall('.//w:t', NAMESPACES) if t.text)
        match = re.match(r'^(\d+)\.\s+', text)
        if match:
            num = int(match.group(1))
            content = re.sub(r'^\d+\.\s+', '', text).strip()
            refs[num] = content

    return refs


def write_sorted_references(references: dict, output_path: str) -> None:
    """
    Not a pure function. Write references sorted by text to a file.

    Useful for manual review - duplicates cluster together when sorted.
    """
    sorted_refs = sorted(references.items(), key=lambda x: x[1])

    with open(output_path, 'w') as f:
        for num, text in sorted_refs:
            f.write(f"{num}. {text}\n\n")


def _extract_ref_key_parts(text: str) -> tuple:
    """
    Pure function. Extract (first_author, year, volume) for duplicate detection.

    >>> _extract_ref_key_parts("Smith J, Jones K. Title. J Name 2020;15:100-10.")
    ('Smith J', '2020', '15')
    >>> _extract_ref_key_parts("No year or volume here")
    ('No year or volume here', None, None)
    >>> _extract_ref_key_parts("Author A, et al. Paper. J 2018;79:1092-105.")
    ('Author A', '2018', '79')
    """
    first_author = text.split(',')[0].strip() if ',' in text else text.split('.')[0].strip()
    year_match = re.search(r'(19|20)\d{2}', text)
    vol_match = re.search(r';(\d+):', text)
    return (
        first_author,
        year_match.group(0) if year_match else None,
        vol_match.group(1) if vol_match else None,
    )


def _resolve_duplicate_chains(duplicates: dict) -> dict:
    """
    Pure function. Resolve chains in duplicate mapping to point to ultimate original.

    >>> _resolve_duplicate_chains({3: 2, 2: 1})
    {3: 1, 2: 1}
    >>> _resolve_duplicate_chains({5: 3, 3: 1, 4: 1})
    {5: 1, 3: 1, 4: 1}
    >>> _resolve_duplicate_chains({})
    {}
    """
    resolved = {}
    for dup, orig in duplicates.items():
        # Follow chain to find ultimate original
        while orig in duplicates:
            orig = duplicates[orig]
        resolved[dup] = orig
    return resolved


def detect_duplicate_references(references: dict, threshold: float = DUPLICATE_THRESHOLD) -> dict:
    """
    Pure function. Detect duplicate references using text similarity.

    Uses SequenceMatcher to find references with >= threshold similarity (default 90%).
    Always uses the lower citation number as the original (preserves first index).
    Resolves chains so all duplicates point to the ultimate original.

    Returns dict mapping duplicate citation numbers to the original citation
    number they should be merged into.

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

    # Pass 1: High text similarity
    for i, n1 in enumerate(nums):
        if n1 in duplicates:
            continue

        for n2 in nums[i+1:]:
            if n2 in duplicates:
                continue

            ratio = SequenceMatcher(None, references[n1], references[n2]).ratio()
            if ratio >= threshold:
                duplicates[n2] = n1

    # Resolve chains
    return _resolve_duplicate_chains(duplicates)


def detect_duplicate_references_with_ai(references: dict, known_duplicates: dict = None) -> dict:
    """
    Not a pure function. Use Claude AI to detect remaining duplicates.

    Filters out known_duplicates first, writes sorted refs to temp file,
    calls Claude subprocess to review, parses output for duplicate pairs.

    Returns dict mapping duplicate citation numbers to originals.
    """
    import subprocess
    import tempfile
    import os

    if known_duplicates is None:
        known_duplicates = {}

    # Filter out known duplicates
    remaining = {k: v for k, v in references.items() if k not in known_duplicates}

    if len(remaining) < 2:
        return {}

    # Write sorted refs to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
        temp_path = f.name
        sorted_refs = sorted(remaining.items(), key=lambda x: x[1])
        for num, text in sorted_refs:
            f.write(f"{num}. {text}\n\n")

    prompt = f'''You are reviewing a bibliography for duplicate references.

Read the file at {temp_path} which contains references sorted alphabetically by text.
Because they are sorted, duplicate or near-duplicate references will appear adjacent to each other.

Your task: Find any duplicate references (same paper cited twice with different numbers).

Look for:
- Identical references
- Same paper with minor formatting differences
- Abbreviated versions (e.g., "Author A, et al." vs full author list)
- Same author + journal + year + volume but different page numbers (likely a typo)

For each duplicate pair found, output EXACTLY in this format (one per line):
DUPLICATE: [higher_number] = [lower_number]

Example:
DUPLICATE: 196 = 129

If no duplicates found, output:
NO_DUPLICATES_FOUND

Only output DUPLICATE lines or NO_DUPLICATES_FOUND. No other text.'''

    try:
        result = subprocess.run(
            ['claude', '--dangerously-skip-permissions', '-p', prompt],
            capture_output=True,
            text=True,
            timeout=120
        )
        output = result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        os.unlink(temp_path)
        raise RuntimeError(f"Claude subprocess failed: {e}")
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)

    # Parse output for DUPLICATE lines
    duplicates = {}
    for line in output.split('\n'):
        line = line.strip()
        match = re.match(r'DUPLICATE:\s*(\d+)\s*=\s*(\d+)', line)
        if match:
            higher = int(match.group(1))
            lower = int(match.group(2))
            if higher > lower:
                duplicates[higher] = lower
            else:
                duplicates[lower] = higher

    return duplicates


def generate_numerical_conversion_table(duplicates: dict) -> str:
    """
    Pure function. Generate markdown table showing numerical conversion.

    Maps old citation numbers to new numbers after duplicate removal.
    Duplicates map to their original; non-duplicates get renumbered to fill gaps.

    >>> generate_numerical_conversion_table({3: 1, 5: 2})[:50]
    '| Old # | New # |\\n|-------|-------|\\n| 1 | 1 |\\n| 2 '
    >>> generate_numerical_conversion_table({})
    '| Old # | New # |\\n|-------|-------|\\n'
    >>> '3 | → 1' in generate_numerical_conversion_table({3: 1})
    True
    """
    if not duplicates:
        return "| Old # | New # |\n|-------|-------|\n"

    # Find all original numbers
    all_originals = set(duplicates.values())
    all_duplicates = set(duplicates.keys())
    max_num = max(max(all_duplicates), max(all_originals))

    # Build mapping: kept numbers get renumbered sequentially
    kept = sorted([n for n in range(1, max_num + 1) if n not in all_duplicates])
    old_to_new = {old: new for new, old in enumerate(kept, 1)}

    lines = ["| Old # | New # |", "|-------|-------|"]
    for old in range(1, max_num + 1):
        if old in all_duplicates:
            # Duplicate - show arrow to original
            orig = duplicates[old]
            new = old_to_new.get(orig, '?')
            lines.append(f"| {old} | → {new} (was {orig}) |")
        else:
            new = old_to_new.get(old, '?')
            lines.append(f"| {old} | {new} |")

    return '\n'.join(lines)


def generate_duplicate_comparison_table(references: dict, duplicates: dict) -> str:
    """
    Pure function. Generate markdown table showing kept vs deleted references.

    Sorted alphabetically by reference text to show duplicates adjacent.
    Uses markdown styling: **bold** for kept, ~~strikethrough~~ for deleted.

    >>> refs = {1: "Alpha paper", 2: "Beta paper", 3: "Alpha paper"}
    >>> table = generate_duplicate_comparison_table(refs, {3: 1})
    >>> '**1**' in table  # Kept
    True
    >>> '~~3~~' in table  # Deleted
    True
    >>> 'Alpha paper' in table
    True
    """
    sorted_refs = sorted(references.items(), key=lambda x: x[1])

    lines = ["| Status | # | Reference Text |", "|--------|---|----------------|"]

    for num, text in sorted_refs:
        truncated = f"{text[:80]}..." if len(text) > 80 else text
        if num in duplicates:
            # This is a duplicate (will be deleted)
            orig = duplicates[num]
            lines.append(f"| ~~DELETED~~ | ~~{num}~~ | ~~{truncated}~~ (→ {orig}) |")
        else:
            # Check if this is an original that has duplicates pointing to it
            has_dups = num in duplicates.values()
            if has_dups:
                lines.append(f"| **KEPT** | **{num}** | **{truncated}** |")
            else:
                lines.append(f"| - | {num} | {truncated} |")

    return '\n'.join(lines)


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


def generate_markdown(
    table_citations: dict,
    paragraph_results: list,
    output_path: str,
    references: dict = None,
    duplicates: dict = None,
) -> tuple:
    """
    Not a pure function. Writes to file.

    Generate markdown output file with sections for tables, paragraphs, conversion,
    and optionally duplicate reference tables.
    Returns tuple of (int, int, int, int): (table_count, paragraph_count, conversion_count, duplicate_count).
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

    # Section 4 & 5: Duplicate Reference Tables (if provided)
    duplicate_count = 0
    if references is not None and duplicates is not None:
        duplicate_count = len(duplicates)

        lines.append("# Section 4: Duplicate Reference Conversion\n")
        lines.append(f"**Total references:** {len(references)}")
        lines.append(f"**Duplicates found:** {duplicate_count}")
        lines.append(f"**Unique references after dedup:** {len(references) - duplicate_count}")
        lines.append("")
        lines.append(generate_numerical_conversion_table(duplicates))
        lines.append("")

        lines.append("# Section 5: Duplicate Comparison (Sorted Alphabetically)\n")
        lines.append("- **KEPT** = Original reference retained")
        lines.append("- ~~DELETED~~ = Duplicate removed (shows which original it maps to)")
        lines.append("- `-` = Unique reference (no duplicates)")
        lines.append("")
        lines.append(generate_duplicate_comparison_table(references, duplicates))
        lines.append("")

    Path(output_path).write_text('\n'.join(lines))
    return len(sorted_tables), len(paragraph_results), len(conversion), duplicate_count


def main():
    docx_path = '/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/JAIP 6391.docx'
    output_path = Path('/Users/ryan/CleanCode/Sandbox/RP_Dumps/MomClaude2/citations_by_paragraph.md')

    print(f"Processing: {docx_path}")

    xml_str = load_docx_xml(docx_path)
    author_names = extract_author_names(xml_str)
    print(f"Found {len(author_names)} author names in references")

    table_citations = process_tables(xml_str, author_names)
    paragraph_results = process_paragraphs(xml_str, author_names)

    # Extract references and detect duplicates
    references = extract_references(docx_path)
    duplicates = detect_duplicate_references(references)
    print(f"Found {len(references)} references, {len(duplicates)} duplicates")

    table_count, para_count, conv_count, dup_count = generate_markdown(
        table_citations, paragraph_results, str(output_path),
        references=references, duplicates=duplicates
    )
    print(f"Extracted {table_count} tables with citations")
    print(f"Extracted {para_count} paragraphs with citations")
    print(f"Citations needing renumbering: {conv_count}")
    print(f"Duplicate references: {dup_count}")
    print(f"Output written to: {output_path}")


if __name__ == '__main__':
    main()
