"""
Email template merge engine.

Supports:
  - {{placeholder}} syntax mapped to spreadsheet column headers
  - Case-insensitive matching
  - Configurable missing-value replacement
  - Round-robin rotation of subject / body templates
  - Optional chaser templates
  - Round-robin sender assignment
"""
import re
from typing import List, Dict, Optional

PLACEHOLDER_RE = re.compile(r'\{\{([\w\s]+)\}\}')


def _normalize(s: str) -> str:
    """Normalise a string for fuzzy column matching: lowercase, strip spaces/underscores/hyphens."""
    return re.sub(r'[\s_\-]+', '', s.lower())


def extract_placeholders(template: str) -> List[str]:
    """Return a deduplicated list of placeholder names found in `template`."""
    return list(dict.fromkeys(PLACEHOLDER_RE.findall(template)))


def _build_header_map(headers: List[str]) -> Dict[str, str]:
    """
    Return a dict mapping normalised placeholder key → original header name.
    Supports exact case-insensitive match AND underscore/space/hyphen normalisation.
    e.g. 'first_name' → 'First Name', 'firstname' → 'First Name'
    """
    m: Dict[str, str] = {}
    for h in headers:
        m[h.lower()] = h           # exact lower-case match
        m[_normalize(h)] = h       # normalised match
    return m


def validate_templates(templates: List[str], headers: List[str]) -> List[str]:
    """
    Check that every {{placeholder}} in every template corresponds to a column header.
    Matching is case-insensitive and ignores spaces/underscores/hyphens.
    Returns a list of human-readable error strings (empty list means valid).
    """
    header_map = _build_header_map(headers)
    errors: List[str] = []
    seen_errors: set = set()

    for template in templates:
        for ph in extract_placeholders(template):
            key = _normalize(ph)
            if key not in header_map and ph.lower() not in header_map:
                if key not in seen_errors:
                    seen_errors.add(key)
                    errors.append(
                        f"Placeholder '{{{{{ph}}}}}' does not match any column. "
                        f"Available columns: {', '.join(headers)}"
                    )
    return errors


def _fill(template: str, row: Dict[str, str],
          header_map: Dict[str, str], missing: str) -> str:
    """Replace all {{placeholders}} in `template` using `row` values."""
    def replacer(match: re.Match) -> str:
        raw = match.group(1)
        # Try exact lower, then normalised
        col = header_map.get(raw.lower()) or header_map.get(_normalize(raw))
        if col is None:
            return missing
        value = row.get(col, '').strip()
        return value if value else missing

    return PLACEHOLDER_RE.sub(replacer, template)


def _expand_inline_variants(template: str, copy_index: int = 0) -> str:
    """
    Replace each {opt1|opt2|...} block sequentially, starting at `copy_index`.
    The first block picks option[copy_index % n], the second picks
    option[(copy_index+1) % n], etc.

    `copy_index` is the per-sender copy counter, so:
      - Copy A emails (index 0): block 0 → opt[0], block 1 → opt[1]
      - Copy B emails (index 1): block 0 → opt[1], block 1 → opt[2]

    This means a single {A|B} block naturally cycles with the A/B copy rotation,
    and multiple blocks within one email each pick the next option in sequence.
    Handles {{placeholders}} safely by hiding them before expansion.
    """
    # Step 1: Hide {{placeholders}} so their braces don't interfere with the regex
    tokens: Dict[str, str] = {}
    ph_counter = [0]

    def hide(m: re.Match) -> str:
        tok = f'\x00PH{ph_counter[0]}\x00'
        tokens[tok] = m.group(0)
        ph_counter[0] += 1
        return tok

    hidden = PLACEHOLDER_RE.sub(hide, template)

    # Step 2: Expand each {opt1|opt2} block — counter starts at copy_index.
    # Regex requires at least one | so bare {word} blocks are never touched.
    block_counter = [copy_index]

    def pick(m: re.Match) -> str:
        options = m.group(1).split('|')
        chosen = options[block_counter[0] % len(options)]
        block_counter[0] += 1
        return chosen

    expanded = re.sub(r'\{([^{}]*\|[^{}]*)\}', pick, hidden)

    # Step 3: Restore {{placeholders}}
    for tok, original in tokens.items():
        expanded = expanded.replace(tok, original)

    return expanded


def perform_merge(
    rows: List[Dict[str, str]],
    headers: List[str],
    subject_templates: List[str],
    body_templates: List[str],
    chaser_body: str,
    sender_emails: List[str],
    missing_value: str = '[MISSING]',
    email_column: str = '',
) -> List[Dict]:
    """
    Perform a row-by-row mail merge.

    Returns a list of enriched row dicts with these extra keys:
      __recipient_email__  – value from the identified email column
      __chaser_body__      – merged chaser body (if supplied)

    Subject / body / variant are deliberately NOT filled here.  The caller
    sorts rows by schedule (date → sender → time) and then calls
    reassign_templates(), which fills __subject_line__ / __email_body__ /
    __template_variant__ from each row's final position so the output reads
    1-2-3-1-2-3.  __sender_account__ likewise comes from the schedule, not
    from this function.
    """
    if not subject_templates:
        raise ValueError("At least one subject line template is required.")
    if not body_templates:
        raise ValueError("At least one email body template is required.")
    if not sender_emails:
        raise ValueError("At least one sender email address is required.")

    # Build normalised column lookup (supports 'first_name' → 'First Name' etc.)
    header_map: Dict[str, str] = _build_header_map(headers)

    # Auto-detect email column if not specified
    if not email_column:
        for h in headers:
            if h.lower() in ('email', 'email address', 'emailaddress',
                             'e-mail', 'recipient email', 'prospect email'):
                email_column = h
                break
        if not email_column and headers:
            # Fall back to first column containing "email"
            for h in headers:
                if 'email' in h.lower():
                    email_column = h
                    break

    merged_rows: List[Dict] = []

    for i, row in enumerate(rows):
        inline_idx = i // len(body_templates)

        enriched = dict(row)
        enriched['__recipient_email__'] = row.get(email_column, '') if email_column else ''
        if chaser_body:
            enriched['__chaser_body__'] = _fill(
                _expand_inline_variants(chaser_body, inline_idx), row, header_map, missing_value)

        merged_rows.append(enriched)

    return merged_rows


def reassign_templates(
    sorted_rows: List[Dict],
    subject_templates: List[str],
    body_templates: List[str],
    headers: List[str],
    missing_value: str = '[MISSING]',
) -> None:
    """
    Re-fill subject / body / variant for every row based on its FINAL position
    in the sorted output so the sheet displays a clean 1-2-3-1-2-3 pattern.

    Must be called AFTER the rows have been sorted (date → sender → time).
    Mutates rows in-place.
    """
    header_map = _build_header_map(headers)
    n_s = len(subject_templates)
    n_b = len(body_templates)

    for j, row in enumerate(sorted_rows):
        s_idx    = j % n_s
        b_idx    = j % n_b
        inline_i = j // max(n_s, n_b)

        row["__subject_line__"] = _fill(
            _expand_inline_variants(subject_templates[s_idx], inline_i),
            row, header_map, missing_value,
        )
        row["__email_body__"] = _fill(
            _expand_inline_variants(body_templates[b_idx], inline_i),
            row, header_map, missing_value,
        )
        row["__template_variant__"] = f"S{s_idx + 1}/B{b_idx + 1}"
