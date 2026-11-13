"""
lines.py - split card text into "lines" (sentences) and delete one.

The Read view shows a card's text one sentence per row with a ✕ to delete it.
``line_spans`` finds each sentence as a (start, end) span in the original
text, so ``remove_line`` can cut exactly that sentence and leave every other
character (paragraph breaks, lists, spacing) as it was.

A new line starts at every line break and after ., !, ? or … followed by a
space and a word that doesn't start in lower case. Very short pieces such as
list numbers ("1.", "a)") stay with the sentence after them.
"""
from __future__ import annotations

import re

HEADING_SEP = " › "  # breadcrumb separator written by parser.build_sections
_PART_RE = re.compile(r"\s*\(part \d+\)$")
_END_RE = re.compile(r"[.!?…]+[\"'”’)\]]*(?=\s+(?![a-z])|$)")
_MARKER_RE = re.compile(r"^\(?[\w]{1,3}[.)]$")  # "1.", "a)", "(iv)", "Dr." stays joined


def heading_path(heading: str | None, fallback: str) -> tuple[str, ...]:
    """Split a breadcrumb heading ("Ch 2 › Cells (part 2)") into its levels."""
    text = _PART_RE.sub("", (heading or "").strip())
    parts = tuple(p.strip() for p in text.split(HEADING_SEP) if p.strip())
    return parts or (fallback,)


def line_spans(text: str) -> list[tuple[int, int]]:
    """(start, end) of every sentence in ``text``, in order, whitespace trimmed."""
    spans: list[tuple[int, int]] = []

    def add(a: int, b: int) -> None:
        while a < b and text[a].isspace():
            a += 1
        while b > a and text[b - 1].isspace():
            b -= 1
        if a < b:
            spans.append((a, b))

    for m in re.finditer(r"[^\n]+", text):
        start, line = m.start(), m.group()
        pos = 0
        for end in _END_RE.finditer(line):
            piece = line[pos:end.end()].strip()
            if len(piece) < 4 or _MARKER_RE.match(piece):
                continue  # too short to stand alone; keep it with what follows
            add(start + pos, start + end.end())
            pos = end.end()
        add(start + pos, start + len(line))
    return spans


def split_lines(text: str) -> list[str]:
    return [text[a:b] for a, b in line_spans(text)]


def remove_line(text: str, index: int) -> str:
    """``text`` without its ``index``-th sentence (see ``line_spans``)."""
    a, b = line_spans(text)[index]
    line_start = text.rfind("\n", 0, a) + 1
    line_end = text.find("\n", b)
    if line_end == -1:
        line_end = len(text)
    before, after = text[line_start:a], text[b:line_end]
    if not before.strip() and not after.strip():  # the sentence was the whole line
        new = text[:line_start] + text[line_end + 1:]
    elif not after.strip():  # last sentence on its line: drop the space before it too
        new = text[:a].rstrip(" \t") + text[line_end:]
    else:
        new = text[:a] + text[b:].lstrip(" \t")
    return re.sub(r"\n{3,}", "\n\n", new).strip()
