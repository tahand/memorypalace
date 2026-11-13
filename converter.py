"""
converter.py - Turn source documents into editable Word (.docx) copies.

This is step 1 of the workflow: the original PDF / EPUB / DOC / TXT / MD is
converted to a Word file that stays as close to the original as possible
(text, formatting and images), so it can be edited and cut down in Word
before it is added to the palace as cards.

* PDF  - ``pdf2docx`` rebuilds pages with their layout, fonts, tables and
         images. Paragraphs printed noticeably larger than the body text are
         given a heading outline level, so they show up in Word's navigation
         pane and become chapter headings when the file is imported later.
         Without pdf2docx, the text and images are extracted the same way an
         import does it.
* EPUB - every chapter's HTML is rebuilt in Word: headings, bold / italic /
         underline, lists, quotes, tables, line breaks and images, with a page
         break between chapters.
* DOC  - converted by LibreOffice or Microsoft Word (see ``parser``).
* DOCX - copied unchanged.
* TXT / MD - headings, paragraphs and (for Markdown) images.

Everything here runs on worker threads: no Tk, no SQLite.
"""
from __future__ import annotations

import io
import logging
import posixpath
import re
import shutil
import tempfile
import warnings as _warnings
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote, urlparse

from parser import (
    SUPPORTED_EXTENSIONS,
    Block,
    MediaWriter,
    MissingDependencyError,
    ParseError,
    _bs_features,
    _parse_md,
    _parse_pdf,
    _parse_txt,
    _require_bs4,
    convert_doc_to_docx,
)

CONVERTIBLE_EXTENSIONS = SUPPORTED_EXTENSIONS
WORD_EXTENSIONS = (".docx",)

# Images python-docx can embed directly; anything else is converted to PNG.
_DOCX_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"}


@dataclass
class ConversionResult:
    source: Path
    output: Path
    method: str
    warnings: list[str] = field(default_factory=list)


def unique_path(folder: Path, stem: str, suffix: str = ".docx") -> Path:
    """``folder/stem.docx``, or ``stem (2).docx`` ... if that name is taken."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" .") or "document"
    candidate = folder / f"{stem}{suffix}"
    n = 2
    while candidate.exists():
        candidate = folder / f"{stem} ({n}){suffix}"
        n += 1
    return candidate


def convert_to_docx(source: str | Path, out_dir: str | Path) -> ConversionResult:
    """Write a Word copy of ``source`` into ``out_dir`` and return where it went.

    Raises ``ParseError`` (or ``MissingDependencyError``) with a readable message.
    """
    source = Path(source)
    out_dir = Path(out_dir)
    if not source.is_file():
        raise ParseError(f"File not found: {source}")
    ext = source.suffix.lower()
    if ext not in CONVERTIBLE_EXTENSIONS:
        raise ParseError(f"Unsupported file type '{ext}'. Supported: {', '.join(CONVERTIBLE_EXTENSIONS)}")
    if source.stat().st_size == 0:
        raise ParseError(f"{source.name} is empty.")
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ParseError(f"Cannot create the folder {out_dir}: {exc}") from exc
    out = unique_path(out_dir, source.stem)

    try:
        if ext == ".docx":
            shutil.copy2(source, out)
            return ConversionResult(source, out, "copied (already a Word file)")
        if ext == ".doc":
            with tempfile.TemporaryDirectory() as tmp:
                converted = convert_doc_to_docx(source, Path(tmp))
                shutil.move(str(converted), out)
            return ConversionResult(source, out, "converted from .doc")
        if ext == ".pdf":
            return _convert_pdf(source, out)
        if ext == ".epub":
            return _convert_epub(source, out)
        return _convert_with_parser(source, out, _parse_md if ext in (".md", ".markdown") else _parse_txt,
                                    "rebuilt from text")
    except ParseError:
        _remove(out)
        raise
    except MemoryError as exc:
        _remove(out)
        raise ParseError(f"{source.name} is too large to convert.") from exc
    except Exception as exc:
        _remove(out)
        raise ParseError(f"Could not convert {source.name}: {type(exc).__name__}: {exc}") from exc


def _remove(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Shared Word helpers
# ---------------------------------------------------------------------------
def _require_docx():
    try:
        import docx
    except ImportError as exc:
        raise MissingDependencyError("python-docx", "create Word documents") from exc
    return docx


_XML_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff￾￿]")


def _safe(text: str) -> str:
    return _XML_ILLEGAL.sub("", text)


def _image_stream(data: bytes, ext: str) -> tuple[io.BytesIO, int | None, int | None] | None:
    """Bytes python-docx can embed, plus pixel size. None if unusable (e.g. SVG)."""
    ext = ext.lower().lstrip(".")
    try:
        from PIL import Image
    except ImportError:
        return (io.BytesIO(data), None, None) if ext in _DOCX_IMAGE_EXTS else None
    try:
        with Image.open(io.BytesIO(data)) as im:
            w, h = im.size
            if (im.format or "").lower() in ("png", "jpeg", "gif", "bmp", "tiff"):
                return io.BytesIO(data), w, h
            buf = io.BytesIO()
            im.convert("RGBA" if im.mode in ("RGBA", "LA", "P") else "RGB").save(buf, "PNG")
            buf.seek(0)
            return buf, w, h
    except Exception:
        return None


def _content_size_in(document) -> tuple[float, float]:
    page = document.sections[0]
    emu = 914400
    width = (page.page_width - page.left_margin - page.right_margin) / emu
    height = (page.page_height - page.top_margin - page.bottom_margin) / emu
    return width, height


def _add_picture(run, document, data: bytes, ext: str) -> bool:
    from docx.shared import Inches
    prepared = _image_stream(data, ext)
    if not prepared:
        return False
    stream, w_px, h_px = prepared
    max_w, max_h = _content_size_in(document)
    max_h *= 0.9
    width = min(max_w, (w_px or 480) / 96)  # 96 dpi, never upscale, fit the page
    if w_px and h_px and width * h_px / w_px > max_h:
        width = max_h * w_px / h_px
    try:
        run.add_picture(stream, width=Inches(width))
    except Exception:
        return False
    return True


def _set_meta(document, title: str | None, author: str | None) -> None:
    if title:
        document.core_properties.title = _safe(title)[:250]
    if author:
        document.core_properties.author = _safe(author)[:250]


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def _convert_pdf(source: Path, out: Path) -> ConversionResult:
    try:
        from pdf2docx import Converter
    except ImportError:
        result = _convert_with_parser(source, out, _parse_pdf, "text and images only")
        result.warnings.insert(0, "Layout was not kept because pdf2docx is not installed "
                                  "(pip install pdf2docx).")
        return result

    logging.getLogger().setLevel(logging.WARNING)  # pdf2docx logs every page at INFO
    try:
        cv = Converter(str(source))
    except Exception as exc:
        raise ParseError(f"Could not open the PDF (corrupted or password-protected?): {exc}") from exc
    toc: list[tuple[int, str]] = []
    meta: dict = {}
    try:
        doc = getattr(cv, "fitz_doc", None)
        if doc is not None:
            if doc.needs_pass:
                raise ParseError(f"{source.name} is password-protected.")
            toc = [(int(lvl), str(title)) for lvl, title, *_ in doc.get_toc(simple=True)]
            meta = doc.metadata or {}
        cv.convert(str(out))
    finally:
        cv.close()
    if not out.is_file():
        raise ParseError(f"pdf2docx produced no output for {source.name}.")

    warnings = []
    try:
        marked = mark_headings(out, toc, title=(meta.get("title") or "").strip() or source.stem,
                               author=(meta.get("author") or "").strip() or None)
    except Exception as exc:  # headings are a bonus; the conversion itself worked
        marked = 0
        warnings.append(f"Could not detect headings: {exc}")
    if not marked:
        warnings.append("No headings were detected; cards will be split by length. "
                        "Apply Word's Heading styles to chapter titles to split them by chapter.")
    return ConversionResult(source, out, "layout kept (pdf2docx)", warnings)


_TOC_LINE_RE = re.compile(r"(\t|\.{3,}|\s{3,})\s*\d+\s*$")  # "Chapter One ....... 12"

# Children of <w:pPr> that must come after <w:outlineLvl> (schema order).
_AFTER_OUTLINE = ("divId", "cnfStyle", "rPr", "sectPr", "pPrChange")


def _norm(text: str) -> str:
    return re.sub(r"\W+", "", text).lower()


def _run_size(run) -> float | None:
    return run.font.size.pt if run.font.size is not None else None


def _set_outline_level(paragraph, level: int) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    ppr = paragraph._p.get_or_add_pPr()
    for old in ppr.findall(qn("w:outlineLvl")):
        ppr.remove(old)
    lvl = OxmlElement("w:outlineLvl")
    lvl.set(qn("w:val"), str(min(max(level, 1), 9) - 1))
    following = [ppr.find(qn(f"w:{tag}")) for tag in _AFTER_OUTLINE]
    following = [el for el in following if el is not None]
    if following:
        min(following, key=lambda el: ppr.index(el)).addprevious(lvl)
    else:
        ppr.append(lvl)


def _body_size(paragraphs) -> float | None:
    """The font size that carries the most characters."""
    weight: Counter[float] = Counter()
    for p in paragraphs:
        for run in p.runs:
            if run.text.strip() and _run_size(run):
                weight[round(_run_size(run), 1)] += len(run.text.strip())
    return weight.most_common(1)[0][0] if weight else None


def _heading_kind(text: str, runs, body: float | None, max_len: int) -> tuple[str, float] | None:
    """('big', size) for large type, ('bold', 0) for bold body-size type, or None."""
    text = text.strip()
    if (len(text) > max_len or sum(c.isalpha() for c in text) < 2
            or _TOC_LINE_RE.search(text) or text.endswith((".", ",", ";"))):
        return None
    runs = [r for r in runs if r.text.strip()]
    if not runs:
        return None
    sizes = [s for s in (_run_size(r) for r in runs) if s]
    size = min(sizes) if sizes else None
    if body and size and size >= body * 1.25:
        return "big", round(size, 1)
    if all(r.bold for r in runs) and (not body or not size or size >= body * 0.95):
        return "bold", 0.0
    return None


def _first_line(paragraph):
    """(runs before the first line break, the run holding the break) or None.

    pdf2docx often puts a heading and the text under it in one paragraph,
    separated by a line break in a run of its own.
    """
    from docx.oxml.ns import qn
    from docx.text.run import Run
    head = []
    for child in paragraph._p:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            return None  # hyperlinks, fields...: leave the paragraph alone
        br = child.find(qn("w:br"))
        if br is not None:
            if br.get(qn("w:type")) not in (None, "textWrapping"):
                return None
            if any(el.tag not in (qn("w:rPr"), qn("w:br")) for el in child) or not head:
                return None
            return head, child
        head.append(Run(child, paragraph))
    return None


def _split_heading_lines(document, toc_keys: set[str], body: float | None, max_len: int) -> None:
    """Move a heading-like first line into a paragraph of its own (same formatting)."""
    import copy

    from docx.oxml.ns import qn
    for p in list(document.paragraphs):
        found = _first_line(p)
        if not found:
            continue
        head, br_run = found
        text = "".join(r.text for r in head)
        if _norm(text) not in toc_keys and not _heading_kind(text, head, body, max_len):
            continue
        new_p = copy.deepcopy(p._p)
        for child in list(new_p):
            if child.tag != qn("w:pPr"):
                new_p.remove(child)
        ppr = new_p.find(qn("w:pPr"))
        if ppr is not None:  # a section break belongs to the original paragraph only
            for sect in ppr.findall(qn("w:sectPr")):
                ppr.remove(sect)
        for run in head:
            new_p.append(run._r)  # moves the run
        p._p.remove(br_run)
        p._p.addprevious(new_p)


def mark_headings(path: Path, toc: list[tuple[int, str]] | None = None, max_len: int = 120,
                  title: str | None = None, author: str | None = None) -> int:
    """Give heading paragraphs an outline level without changing how they look.

    With the PDF's bookmarks (``toc``), the paragraphs matching them are
    marked. Otherwise headings are guessed: short paragraphs printed at least
    25% larger than the body text (biggest = level 1), and short paragraphs
    entirely in bold at body size (the next level down). Table-of-contents
    lines ("Title .... 12") and sentences are left alone. A heading-like first
    line followed by a line break is split into its own paragraph first. Only
    body paragraphs are considered: an import reads table cells as plain text.

    ``title`` and ``author`` are stored in the document properties if given.
    Returns how many paragraphs were marked.
    """
    docx = _require_docx()
    document = docx.Document(str(path))
    _set_meta(document, title, author)
    body = _body_size(document.paragraphs)
    toc_keys = {_norm(title) for _, title in toc or []} - {""}
    _split_heading_lines(document, toc_keys, body, max_len)
    paragraphs = [p for p in document.paragraphs if p.text.strip()]
    marked = _mark_from_toc(paragraphs, toc) if toc else 0
    if not marked:
        marked = _mark_from_fonts(paragraphs, body, max_len)
    document.save(str(path))
    return marked


def _mark_from_toc(paragraphs, toc: list[tuple[int, str]]) -> int:
    norms = [_norm(p.text) for p in paragraphs]
    pos = marked = 0
    for level, title in toc:
        key = _norm(title)
        if not key:
            continue
        for i in range(pos, len(paragraphs)):
            if norms[i] == key:
                _set_outline_level(paragraphs[i], level)
                pos = i + 1
                marked += 1
                break
    return marked


def _mark_from_fonts(paragraphs, body: float | None, max_len: int) -> int:
    found = [(p, kind) for p in paragraphs if (kind := _heading_kind(p.text, p.runs, body, max_len))]
    big_sizes = sorted({size for _, (k, size) in found if k == "big"}, reverse=True)[:2]
    levels = {s: i + 1 for i, s in enumerate(big_sizes)}
    for p, (k, size) in found:
        _set_outline_level(p, levels.get(size, 2) if k == "big" else len(levels) + 1)
    return len(found)


# ---------------------------------------------------------------------------
# Blocks (TXT, Markdown, and PDF without pdf2docx)
# ---------------------------------------------------------------------------
def _convert_with_parser(source: Path, out: Path, handler, method: str) -> ConversionResult:
    docx = _require_docx()
    with tempfile.TemporaryDirectory() as tmp:
        media_dir = Path(tmp) / "media"
        title, author, blocks, warnings = handler(source, MediaWriter(media_dir, "media"))
        if not any(b.kind in ("heading", "text", "image") for b in blocks):
            raise ParseError(f"No readable text or images were found in {source.name}.")
        document = docx.Document()
        _set_meta(document, title or source.stem, author)
        write_blocks(document, blocks, Path(tmp))
        document.save(str(out))
    return ConversionResult(source, out, method, list(warnings))


def write_blocks(document, blocks: list[Block], media_root: Path) -> None:
    for block in blocks:
        if block.kind == "heading":
            document.add_heading(_safe(block.text), min(max(block.level, 1), 9))
        elif block.kind == "text":
            lines = _safe(block.text).split("\n")
            p = document.add_paragraph()
            run = p.add_run(lines[0])
            for line in lines[1:]:
                run.add_break()
                run = p.add_run(line)
        elif block.kind == "image":
            path = media_root / block.path
            try:
                data = path.read_bytes()
            except OSError:
                continue
            _add_picture(document.add_paragraph().add_run(), document, data, path.suffix)


# ---------------------------------------------------------------------------
# EPUB
# ---------------------------------------------------------------------------
_HEADINGS = {f"h{i}": i for i in range(1, 7)}
_BLOCK_TAGS = {"p", "div", "section", "article", "header", "footer", "aside", "main", "figure",
               "figcaption", "caption", "address", "dt", "dd", "center", "body"}
_SKIP_TAGS = {"script", "style", "head", "title", "noscript", "template", "nav", "svg", "math"}
_BOLD = {"b", "strong"}
_ITALIC = {"i", "em", "cite", "dfn", "var"}
_UNDERLINE = {"u", "ins"}
_STRIKE = {"s", "strike", "del"}
_MONO = {"code", "kbd", "samp", "tt"}


def _epub_meta(book, name: str) -> str | None:
    try:
        values = book.get_metadata("DC", name)
        return str(values[0][0]).strip() if values else None
    except Exception:
        return None


def _convert_epub(source: Path, out: Path) -> ConversionResult:
    try:
        import ebooklib
        from ebooklib import epub
    except ImportError as exc:
        raise MissingDependencyError("EbookLib", "convert EPUB files") from exc
    BeautifulSoup = _require_bs4()
    docx = _require_docx()

    try:
        book = epub.read_epub(str(source), {"ignore_ncx": True})
    except Exception as exc:
        raise ParseError(f"Could not open EPUB (corrupted or DRM-protected?): {exc}") from exc

    document = docx.Document()
    _set_meta(document, _epub_meta(book, "title") or source.stem, _epub_meta(book, "creator"))
    writer = _HtmlWriter(document)
    warnings: list[str] = []
    features = _bs_features()
    chapters = 0

    for entry in book.spine:
        idref = entry[0] if isinstance(entry, (tuple, list)) else entry
        item = book.get_item_with_id(idref)
        if item is None or item.get_type() != ebooklib.ITEM_DOCUMENT or isinstance(item, epub.EpubNav):
            continue
        try:
            with _warnings.catch_warnings():  # XHTML parsed as HTML is what we want here
                _warnings.simplefilter("ignore")
                soup = BeautifulSoup(item.get_content(), features)
        except Exception as exc:
            warnings.append(f"Skipped unreadable chapter {item.get_name()}: {exc}")
            continue
        base = posixpath.dirname(item.get_name())

        def load_image(src: str, base: str = base) -> tuple[bytes, str] | None:
            href = unquote(src.split("#")[0].split("?")[0])
            if not href or urlparse(href).scheme:
                return None
            href = posixpath.normpath(posixpath.join(base, href))
            img = book.get_item_with_href(href)
            return (img.get_content(), posixpath.splitext(href)[1]) if img else None

        if chapters and writer.has_content():
            writer.page_break()
        writer.load_image = load_image
        writer.write(soup.body or soup)
        chapters += 1

    if not writer.has_content():
        raise ParseError(f"No readable text or images were found in {source.name}.")
    if writer.skipped_images:
        warnings.append(f"{writer.skipped_images} image(s) could not be embedded (e.g. SVG).")
    document.save(str(out))
    return ConversionResult(source, out, "rebuilt chapter by chapter", warnings)


class _HtmlWriter:
    """Walk HTML and write it into a python-docx Document, keeping basic formatting."""

    def __init__(self, document):
        self.doc = document
        self.container = document  # where paragraphs go: the body or a table cell
        self.para = None           # current paragraph, created lazily on first text
        self.style: str | None = None
        self.skipped_images = 0
        self.load_image = lambda src: None
        self._content = False
        self._at_start = True      # nothing but whitespace in the current paragraph yet
        self._list_stack: list[str] = []

    def has_content(self) -> bool:
        return self._content

    def page_break(self) -> None:
        from docx.enum.text import WD_BREAK
        self.end_paragraph()
        self.container.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # -- paragraphs ---------------------------------------------------------
    def end_paragraph(self) -> None:
        self.para = None
        self._at_start = True

    def _style(self, name: str | None) -> str | None:
        if not name:
            return None
        try:
            self.doc.styles[name]
            return name
        except KeyError:
            return None

    def _current(self):
        if self.para is None:
            self.para = self.container.add_paragraph(style=self._style(self.style))
            self._at_start = True
        return self.para

    # -- runs ----------------------------------------------------------------
    def text(self, text: str, fmt: frozenset[str]) -> None:
        if "pre" not in fmt:
            text = re.sub(r"\s+", " ", text)
            if self._at_start:
                text = text.lstrip()
        if not text:
            return
        from docx.shared import Pt
        run = self._current().add_run(_safe(text))
        self._at_start = False
        if "b" in fmt:
            run.bold = True
        if "i" in fmt:
            run.italic = True
        if "u" in fmt:
            run.underline = True
        if "s" in fmt:
            run.font.strike = True
        if "sup" in fmt:
            run.font.superscript = True
        elif "sub" in fmt:
            run.font.subscript = True
        if "mono" in fmt or "pre" in fmt:
            run.font.name = "Consolas"
            run.font.size = Pt(10)
        self._content = True

    def line_break(self) -> None:
        self._current().add_run().add_break()
        self._at_start = True

    def image(self, src: str | None, inline: bool) -> None:
        loaded = self.load_image(src) if src else None
        if not loaded:
            if src:
                self.skipped_images += 1
            return
        if not inline:
            self.end_paragraph()
        para = self._current()
        if _add_picture(para.add_run(), self.doc, *loaded):
            self._content = True
            self._at_start = False
        else:
            self.skipped_images += 1
        if not inline:
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.end_paragraph()

    def _block(self, child, fmt: frozenset[str], style: str | None = None) -> None:
        """Write ``child`` as its own paragraph(s), in ``style`` if given."""
        outer = self.style
        self.end_paragraph()
        if style:
            self.style = style
        self.write(child, fmt)
        self.end_paragraph()
        self.style = outer

    # -- tree walk -----------------------------------------------------------
    def write(self, node, fmt: frozenset[str] = frozenset()) -> None:
        from bs4 import NavigableString, Tag
        from bs4.element import Comment, Doctype, ProcessingInstruction

        for child in node.children:
            if isinstance(child, (Comment, Doctype, ProcessingInstruction)):
                continue
            if isinstance(child, NavigableString):
                self.text(str(child), fmt)
                continue
            if not isinstance(child, Tag):
                continue
            name = (child.name or "").lower()
            if name in _SKIP_TAGS:
                if name == "svg":  # SVG-wrapped cover images: use their <image href>
                    img = child.find("image")
                    if img is not None:
                        self.image(img.get("xlink:href") or img.get("href"), inline=False)
                continue
            if name == "br":
                self.line_break()
            elif name == "img":
                self.image(child.get("src"), inline=self.para is not None and not self._at_start)
            elif name == "hr":
                self.end_paragraph()
            elif name in _HEADINGS:
                self._block(child, fmt, f"Heading {_HEADINGS[name]}")
            elif name in ("ul", "ol"):
                self._list_stack.append(name)
                self._block(child, fmt)
                self._list_stack.pop()
            elif name == "li":
                kind = self._list_stack[-1] if self._list_stack else "ul"
                depth = min(max(len(self._list_stack), 1), 3)
                base = "List Number" if kind == "ol" else "List Bullet"
                self._block(child, fmt, base if depth == 1 else f"{base} {depth}")
            elif name == "blockquote":
                self._block(child, fmt, "Quote")
            elif name == "pre":
                self._block(child, fmt | {"pre"})
            elif name == "table":
                self._table(child, fmt)
            elif name in _BLOCK_TAGS or name == "dl":
                self._block(child, fmt)
            else:  # inline formatting (unknown tags are treated as inline)
                extra = set()
                if name in _BOLD:
                    extra.add("b")
                if name in _ITALIC:
                    extra.add("i")
                if name in _UNDERLINE:
                    extra.add("u")
                if name in _STRIKE:
                    extra.add("s")
                if name in _MONO:
                    extra.add("mono")
                if name in ("sup", "sub"):
                    extra.add(name)
                self.write(child, fmt | extra)

    def _table(self, table, fmt: frozenset[str]) -> None:
        rows = []
        for tr in table.find_all("tr"):
            if tr.find_parent("table") is not table:
                continue  # a row of a nested table
            cells = [c for c in tr.find_all(["td", "th"]) if c.find_parent("tr") is tr]
            if cells:
                rows.append(cells)
        if not rows:
            return
        self.end_paragraph()
        word_table = self.container.add_table(rows=len(rows), cols=max(len(r) for r in rows))
        if self._style("Table Grid"):
            word_table.style = "Table Grid"
        outer, outer_style = self.container, self.style
        for r, cells in enumerate(rows):
            for c, cell in enumerate(cells):
                target = word_table.cell(r, c)
                self.container, self.style = target, None
                self.para, self._at_start = target.paragraphs[0], True
                self.write(cell, fmt | ({"b"} if cell.name == "th" else frozenset()))
        self.container, self.style = outer, outer_style
        self.end_paragraph()
        self._content = True


