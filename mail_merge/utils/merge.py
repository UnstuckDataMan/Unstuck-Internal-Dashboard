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


def perform_merge(
    rows: List[Dict[str, str]],
    headers: List[str],
    subject_templates: List[str],
    body_templates: List[str],
    chaser_subject: str,
    chaser_body: str,
    sender_emails: List[str],
    missing_value: str = '[MISSING]',
    email_column: str = '',
) -> List[Dict]:
    """
    Perform a row-by-row mail merge.

    Returns a list of enriched row dicts with these extra keys:
      __sender_account__   – the rotated sender email
      __recipient_email__  – value from the identified email column
      __subject_line__     – merged subject
      __email_body__       – merged body
      __chaser_subject__   – merged chaser subject (if supplied)
      __chaser_body__      – merged chaser body (if supplied)
      __template_variant__ – e.g. "S1/B2"
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
        s_idx = i % len(subject_templates)
        b_idx = i % len(body_templates)
        sender = sender_emails[i % len(sender_emails)]

        subject = _fill(subject_templates[s_idx], row, header_map, missing_value)
        body = _fill(body_templates[b_idx], row, header_map, missing_value)

        enriched = dict(row)
        enriched['__sender_account__'] = sender
        enriched['__recipient_email__'] = row.get(email_column, '') if email_column else ''
        enriched['__subject_line__'] = subject
        enriched['__email_body__'] = body
        enriched['__template_variant__'] = f"S{s_idx + 1}/B{b_idx + 1}"

        if chaser_subject:
            enriched['__chaser_subject__'] = _fill(
                chaser_subject, row, header_map, missing_value)
        if chaser_body:
            enriched['__chaser_body__'] = _fill(
                chaser_body, row, header_map, missing_value)

        merged_rows.append(enriched)

    return merged_rows
