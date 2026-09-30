"""Sanitiser for the account manager's report write-up.

The write-up is authored in a rich text editor and stored as HTML, then
rendered on a page clients open from a link. That makes it the one place in
Pulse where markup crosses from an author to a reader, so it is cleaned on the
way IN — once, at the point of writing — rather than trusted on the way out.

Allow-list, not block-list. A block-list has to anticipate every way a script
can be smuggled in; an allow-list only has to name the handful of tags a
write-up legitimately needs. Anything not named here is dropped, and its text
kept: a pasted `<div>` of prose becomes prose rather than disappearing.

No external dependency: bleach would be the obvious choice, but this needs
about forty lines and the project deliberately carries very few packages.
"""
from __future__ import annotations

import re
from html import escape
from html.parser import HTMLParser

# Formatting a write-up actually uses. Deliberately no images, no tables and
# no <span style>: an account manager needs structure and emphasis, and every
# extra tag is another thing to have to reason about on a client's page.
ALLOWED_TAGS: frozenset[str] = frozenset({
    "p", "br",
    "strong", "b", "em", "i", "u", "s",
    "ul", "ol", "li",
    "h2", "h3",
    "blockquote",
    "a",
})

# Tags whose entire contents are discarded, not just the tag. Keeping the text
# of a <script> would put the source of the script into the report.
_DROP_CONTENT: frozenset[str] = frozenset({"script", "style", "title"})

VOID_TAGS: frozenset[str] = frozenset({"br"})

# A <p> may not contain these. contenteditable emits <p><ul>...</ul></p>
# routinely; browsers unwrap it on parse, but storing invalid markup that
# renders on a client's page is not something to rely on a parser to rescue.
_BLOCK_LEVEL: frozenset[str] = frozenset({
    "p", "ul", "ol", "h2", "h3", "blockquote",
})

# Only on <a>, and only after the scheme check below.
ALLOWED_ATTRS: dict[str, frozenset[str]] = {"a": frozenset({"href"})}

# mailto is allowed because a report may well say "reply to your account
# manager". javascript:, data: and vbscript: are the ones this exists to stop.
_SAFE_SCHEMES: tuple[str, ...] = ("http://", "https://", "mailto:")

MAX_BODY_CHARS = 20_000

_EMPTY_P = re.compile(r"<p>\s*</p>")


class _Cleaner(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self._open: list[str] = []
        self._muted = 0          # depth inside a drop-content tag

    # ── Tags ────────────────────────────────────────────────────────────────
    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in _DROP_CONTENT:
            self._muted += 1
            return
        if self._muted or tag not in ALLOWED_TAGS:
            return
        # Close an open <p> before a block that cannot live inside one. The
        # stray </p> that follows is dropped by handle_endtag, which only
        # closes tags it actually opened.
        if tag in _BLOCK_LEVEL and "p" in self._open:
            while self._open:
                current = self._open.pop()
                self.out.append(f"</{current}>")
                if current == "p":
                    break
        kept = self._attrs(tag, attrs)
        if tag in VOID_TAGS:
            self.out.append(f"<{tag}>")
            return
        self._open.append(tag)
        self.out.append(f"<{tag}{kept}>")

    def handle_startendtag(self, tag: str, attrs) -> None:
        if not self._muted and tag in ALLOWED_TAGS:
            self.out.append(f"<{tag}>" if tag in VOID_TAGS else f"<{tag}></{tag}>")

    def handle_endtag(self, tag: str) -> None:
        if tag in _DROP_CONTENT:
            self._muted = max(0, self._muted - 1)
            return
        if self._muted or tag not in ALLOWED_TAGS or tag in VOID_TAGS:
            return
        # Close only a tag we actually opened, and close anything left open
        # inside it. Editors emit mismatched markup often enough that trusting
        # the input would let a stray </p> unbalance the whole document.
        if tag not in self._open:
            return
        while self._open:
            current = self._open.pop()
            self.out.append(f"</{current}>")
            if current == tag:
                break

    def handle_data(self, data: str) -> None:
        if not self._muted:
            self.out.append(escape(data, quote=False))

    # ── Attributes ──────────────────────────────────────────────────────────
    def _attrs(self, tag: str, attrs) -> str:
        allowed = ALLOWED_ATTRS.get(tag)
        if not allowed:
            return ""
        out = []
        for name, value in attrs:
            if name not in allowed or not value:
                continue
            if name == "href":
                if not _safe_href(value):
                    continue
                # Client-facing and off-site: open in a new tab, and deny the
                # destination any handle back on the report's window.
                out.append(f' href="{escape(value, quote=True)}"'
                           f' target="_blank" rel="noopener noreferrer nofollow"')
        return "".join(out)

    def close_all(self) -> str:
        while self._open:
            self.out.append(f"</{self._open.pop()}>")
        return "".join(self.out)


def _safe_href(value: str) -> bool:
    # Strip whitespace and control characters first: "java\tscript:alert(1)" is
    # a URL browsers will happily run and a naive startswith will happily pass.
    cleaned = "".join(c for c in value if c.isprintable() and not c.isspace()).lower()
    if cleaned.startswith("/") or cleaned.startswith("#"):
        return True
    return cleaned.startswith(_SAFE_SCHEMES)


def sanitize(html: str) -> str:
    """Clean authored HTML for storage. Returns markup safe to render as-is."""
    if not html:
        return ""
    cleaner = _Cleaner()
    cleaner.feed(html[:MAX_BODY_CHARS])
    cleaner.close()
    out = cleaner.close_all().strip()
    # Unwrapping a block out of a <p> leaves the empty shell behind. A
    # deliberate blank line is <p><br></p> and is kept.
    return _EMPTY_P.sub("", out)


def to_text(html: str) -> str:
    """Plain text of a write-up, for previews and for "is this empty?".

    An editor that has been focused and emptied still emits `<p><br></p>`,
    which is not blank as a string but is blank as a report.
    """
    parser = _TextOnly()
    parser.feed(html or "")
    parser.close()
    return " ".join("".join(parser.parts).split())


# Flattening has to put whitespace at these boundaries or the last word of a
# paragraph joins the first word of the next: "the founder sequence."
# + "Recommendations" reads as one word in a summary.
_BLOCK_TAGS: frozenset[str] = frozenset({
    "p", "br", "li", "ul", "ol", "h2", "h3", "blockquote", "div", "tr",
})


class _TextOnly(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._muted = 0

    def handle_starttag(self, tag, attrs):
        if tag in _DROP_CONTENT:
            self._muted += 1
        elif tag in _BLOCK_TAGS:
            self.parts.append(" ")

    def handle_startendtag(self, tag, attrs):
        if tag in _BLOCK_TAGS:
            self.parts.append(" ")

    def handle_endtag(self, tag):
        if tag in _DROP_CONTENT:
            self._muted = max(0, self._muted - 1)
        elif tag in _BLOCK_TAGS:
            self.parts.append(" ")

    def handle_data(self, data):
        if not self._muted:
            self.parts.append(data)
