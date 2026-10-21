"""
parser.py - Turn source documents into atomic, reviewable sections.

Pipeline
--------
1. A format-specific extractor walks the file in reading order and emits a
   flat stream of ``Block`` objects: headings (with a level), text
   paragraphs, images (already written to the media folder) and hard breaks.
2. ``build_sections`` turns that stream into ``ParsedSection`` objects:
   * a heading of level <= ``SPLIT_LEVEL`` starts a new section;
   * a section longer than ``max_chars`` is split at paragraph boundaries
     (and very long paragraphs at sentence boundaries) into "part 2", ...;
   * every image is attached to the section in which it appeared.

Because every format goes through the same chunker, sections from PDFs,
EPUBs, Word files and Markdown look and behave the same.

Optional dependencies are imported lazily. A missing package raises
``MissingDependencyError`` with the matching ``pip install`` command.
"""
from __future__ import annotations

import hashlib
import io
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from urllib.parse import unquote, urlparse

DEFAULT_MAX_CHARS = 1800
SPLIT_LEVEL = 3
SUPPORTED_EXTENSIONS = (".pdf", ".epub", ".docx", ".doc", ".txt", ".md", ".markdown")

# Formats the UI and exporters can use directly. Anything else is converted
# to PNG when Pillow is available.
_WEB_IMAGE_EXTS = {"png", "jpg", "gif", "bmp", "webp"}


class ParseError(Exception):
    """A file could not be turned into sections (corrupt, encrypted, empty...)."""


class MissingDependencyError(ParseError):
    def __init__(self, package: str, purpose: str):
        super().__init__(
            f"The '{package}' package is needed to {purpose}.\nInstall it with:  pip install {package}"
        )
        self.package = package


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class Block:
    kind: str            # 'heading' | 'text' | 'image' | 'break'
    text: str = ""
    level: int = 0       # heading level, 1 = top
    path: str = ""       # relative media path for images


@dataclass
class ParsedSection:
    heading: str
    content_text: str
    image_paths: list[str] = field(default_factory=list)


@dataclass
class ParsedDocument:
    title: str
    author: str | None
    format: str
    sections: list[ParsedSection]
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Media handling
# ---------------------------------------------------------------------------
class MediaWriter:
    """Write extracted images into one document's media folder.

    * Identical images (same bytes) are stored only once.
    * Tiny images (bullets, spacers, tracking pixels) are skipped.
    * With ``skip_repeats`` set, an image seen before is dropped rather than
      re-attached. PDFs use this so a logo printed on every page is kept once
      instead of appearing on every card.
    """

    def __init__(self, media_dir: Path, rel_prefix: str, min_bytes: int = 200, min_dim: int = 32):
        self.media_dir = Path(media_dir)
        self.rel_prefix = rel_prefix.rstrip("/")
        self.min_bytes = min_bytes
        self.min_dim = min_dim
        self.skip_repeats = False
        self._seen: dict[str, str] = {}
        self._count = 0

    def save(self, data: bytes, ext: str, width: int | None = None, height: int | None = None) -> str | None:
        if not data or len(data) < self.min_bytes:
            return None
        digest = hashlib.sha1(data).hexdigest()
        if digest in self._seen:
            return None if self.skip_repeats else self._seen[digest]

        ext = (ext or "").lower().lstrip(".") or "bin"
        if ext in ("jpeg", "jpe"):
            ext = "jpg"
        if width is None or height is None:
            width, height = _image_size(data)
        if width and height and min(width, height) < self.min_dim:
            self._seen[digest] = ""  # remember it was rejected
            return None
        if ext not in _WEB_IMAGE_EXTS:
            converted = _to_png(data)
            if converted is not None:
                data, ext = converted, "png"

        self.media_dir.mkdir(parents=True, exist_ok=True)
        self._count += 1
        name = f"img_{self._count:04d}.{ext}"
        (self.media_dir / name).write_bytes(data)
        rel = f"{self.rel_prefix}/{name}"
        self._seen[digest] = rel
        return rel

    def save_file(self, path: Path) -> str | None:
        try:
            return self.save(path.read_bytes(), path.suffix)
        except OSError:
            return None


def _image_size(data: bytes) -> tuple[int | None, int | None]:
    try:
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            return im.size
    except Exception:
        return None, None


def _to_png(data: bytes) -> bytes | None:
    """Convert exotic formats (TIFF, JPEG2000, WMF/EMF on Windows...) to PNG."""
    try:
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            im.load()
            if im.mode not in ("RGB", "RGBA", "L", "LA", "P"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            out = io.BytesIO()
            im.save(out, "PNG")
            return out.getvalue()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def clean_text(text: str, keep_newlines: bool = False) -> str:
    text = _CONTROL_CHARS.sub("", text.replace("\xa0", " ").replace("\u00ad", ""))
    if keep_newlines:
        lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.splitlines()]
        return "\n".join(lines).strip()
    return re.sub(r"\s+", " ", text).strip()


def _split_long_paragraph(text: str, max_chars: int) -> list[str]:
    """Split one long paragraph at sentence (then word) boundaries."""
    if len(text) <= max_chars:
        return [text]
    sentences = re.split(r"(?<=[.!?…])\s+", text)
    pieces: list[str] = []
    current = ""
    for sentence in sentences:
        while len(sentence) > max_chars:  # one enormous "sentence": hard wrap
            cut = sentence.rfind(" ", 0, max_chars)
            cut = cut if cut > max_chars // 2 else max_chars
            if current:
                pieces.append(current)
                current = ""
            pieces.append(sentence[:cut].strip())
            sentence = sentence[cut:].strip()
        if current and len(current) + 1 + len(sentence) > max_chars:
            pieces.append(current)
            current = sentence
        else:
            current = f"{current} {sentence}".strip()
    if current:
        pieces.append(current)
    return pieces


def build_sections(
    blocks: list[Block],
    fallback_title: str,
    max_chars: int = DEFAULT_MAX_CHARS,
    split_level: int = SPLIT_LEVEL,
) -> list[ParsedSection]:
    """Group a block stream into sections (see module docstring)."""
    sections: list[ParsedSection] = []
    heading_stack: list[tuple[int, str]] = []
    paras: list[str] = []
    images: list[str] = []
    size = 0
    part = 1

    def current_heading() -> str:
        crumb = " › ".join(text for _, text in heading_stack) or fallback_title
        return crumb if part == 1 else f"{crumb} (part {part})"

    def flush() -> None:
        nonlocal paras, images, size, part
        text = "\n\n".join(paras).strip()
        if text or images:
            sections.append(ParsedSection(current_heading(), text, images))
            part += 1
        paras, images, size = [], [], 0

    for block in blocks:
        if block.kind == "heading":
            text = clean_text(block.text)
            if not text:
                continue
            if block.level <= split_level:
                flush()
                part = 1
                while heading_stack and heading_stack[-1][0] >= block.level:
                    heading_stack.pop()
                heading_stack.append((block.level, text[:200]))
            else:
                # Minor headings stay inside the current section as a paragraph.
                paras.append(text)
                size += len(text)
        elif block.kind == "text":
            text = block.text.strip()
            if not text:
                continue
            for piece in _split_long_paragraph(text, max_chars):
                if size and size + len(piece) > max_chars:
                    flush()
                paras.append(piece)
                size += len(piece)
        elif block.kind == "image":
            if block.path and block.path not in images:
                images.append(block.path)
        elif block.kind == "break":
            flush()
            part = 1
    flush()
    return sections


def _title_from_filename(path: Path) -> str:
    return re.sub(r"[_\-]+", " ", path.stem).strip().title() or path.name


def _read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    if b"\x00" in raw[:4096]:
        raise ParseError(f"{path.name} does not look like a text file (binary content).")
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")  # never fails


# ---------------------------------------------------------------------------
# Plain text
# ---------------------------------------------------------------------------
_NUMBER_WORDS = r"one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|[0-9]+|[ivxlcdm]+"
_CHAPTER_RE = re.compile(
    rf"^(chapter|part|book|section|lesson|unit)\s+({_NUMBER_WORDS})\b[^\n]{{0,80}}$"
    r"|^(prologue|epilogue|introduction|preface|foreword|afterword|appendix|conclusion)\b[^\n.]{0,60}$",
    re.IGNORECASE,
)
_SEPARATOR_RE = re.compile(r"^\s*([*\-_=~#•]\s*){3,}$")


def _parse_txt(path: Path, media: MediaWriter):
    return None, None, text_to_blocks(_read_text_file(path)), []


def split_text(text: str, fallback_title: str, max_chars: int = DEFAULT_MAX_CHARS) -> list[ParsedSection]:
    """Split free text (typed, pasted or OCR'd) into sections like a .txt import."""
    return build_sections(text_to_blocks(text), fallback_title, max_chars)


def text_to_blocks(text: str) -> list[Block]:
    """Plain text -> blocks. Paragraphs are separated by blank lines; "Chapter N"
    lines and short ALL-CAPS lines become headings; ``***`` lines are breaks."""
    blocks: list[Block] = []
    for para in re.split(r"\n\s*\n", text):
        stripped = para.strip()
        if not stripped:
            continue
        if _SEPARATOR_RE.match(stripped):
            blocks.append(Block("break"))
            continue
        single_line = "\n" not in stripped
        if single_line and _CHAPTER_RE.match(stripped):
            blocks.append(Block("heading", stripped, 1))
        elif (
            single_line and len(stripped) <= 60 and stripped.isupper()
            and sum(c.isalpha() for c in stripped) >= 3
        ):
            blocks.append(Block("heading", stripped.title(), 2))
        else:
            # Hard-wrapped lines inside a paragraph are joined.
            blocks.append(Block("text", clean_text(stripped)))
    return blocks


# ---------------------------------------------------------------------------
# Markdown
# ---------------------------------------------------------------------------
_ATX_RE = re.compile(r"^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$")
_SETEXT_RE = re.compile(r"^\s{0,3}(=+|-+)\s*$")
_FENCE_RE = re.compile(r"^\s{0,3}(```|~~~)")
_MD_IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(\s*<?([^)\s>]+)>?(?:\s+[\"'][^\"']*[\"'])?\s*\)")
_HTML_IMG_RE = re.compile(r"<img\b[^>]*\bsrc=[\"']([^\"']+)[\"'][^>]*>", re.IGNORECASE)


def _parse_front_matter(text: str) -> tuple[dict[str, str], str]:
    if not text.startswith("---"):
        return {}, text
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.DOTALL)
    if not m:
        return {}, text
    meta = {}
    for line in m.group(1).splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            meta[key.strip().lower()] = value.strip().strip("'\"")
    return meta, text[m.end():]


def _parse_md(path: Path, media: MediaWriter):
    meta, text = _parse_front_matter(_read_text_file(path))
    base_dir = path.parent
    blocks: list[Block] = []
    warnings: list[str] = []
    para: list[str] = []

    def emit_paragraph(raw: bool = False) -> None:
        if not para:
            return
        content = "\n".join(para)
        para.clear()
        image_srcs: list[str] = []
        if not raw:
            image_srcs = [m.group(2) for m in _MD_IMAGE_RE.finditer(content)]
            image_srcs += [m.group(1) for m in _HTML_IMG_RE.finditer(content)]
            content = _HTML_IMG_RE.sub("", _MD_IMAGE_RE.sub("", content))
        content = content.strip("\n") if raw else clean_text(content, keep_newlines=True)
        if content.strip():
            blocks.append(Block("text", content))
        for src in image_srcs:
            if urlparse(src).scheme in ("http", "https", "data"):
                warnings.append(f"Skipped remote image {src[:80]}")
                continue
            img_path = (base_dir / unquote(src)).resolve()
            rel = media.save_file(img_path) if img_path.is_file() else None
            if rel:
                blocks.append(Block("image", path=rel))
            elif not img_path.is_file():
                warnings.append(f"Image not found: {src}")

    lines = text.splitlines()
    in_fence, fence = False, ""
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        fence_match = _FENCE_RE.match(line)
        if in_fence:
            para.append(line)
            if fence_match and stripped.startswith(fence):
                in_fence = False
                emit_paragraph(raw=True)
            i += 1
            continue
        if fence_match:
            emit_paragraph()
            in_fence, fence = True, fence_match.group(1)
            para.append(line)
            i += 1
            continue
        atx = _ATX_RE.match(line)
        if atx:
            emit_paragraph()
            blocks.append(Block("heading", atx.group(2), len(atx.group(1))))
            i += 1
            continue
        if (
            stripped and not para and i + 1 < len(lines)
            and _SETEXT_RE.match(lines[i + 1]) and not _SETEXT_RE.match(line)
        ):
            blocks.append(Block("heading", stripped, 1 if "=" in lines[i + 1] else 2))
            i += 2
            continue
        if not stripped:
            emit_paragraph()
        elif _SEPARATOR_RE.match(stripped) and not para:
            blocks.append(Block("break"))
        else:
            para.append(line)
        i += 1
    emit_paragraph(raw=in_fence)

    title = meta.get("title")
    if not title:  # fall back to the first H1
        title = next((b.text for b in blocks if b.kind == "heading" and b.level == 1), None)
    return title, meta.get("author"), blocks, warnings


# ---------------------------------------------------------------------------
# HTML (shared by EPUB)
# ---------------------------------------------------------------------------
_HEADING_TAGS = {f"h{i}": i for i in range(1, 7)}
_TEXT_BLOCK_TAGS = {"p", "li", "dt", "dd", "figcaption", "caption", "address", "pre"}
_SKIP_TAGS = {"script", "style", "head", "title", "noscript", "template", "nav"}
_INLINE_TAGS = {"a", "span", "em", "strong", "b", "i", "u", "small", "sup", "sub", "code", "abbr", "cite", "q", "mark", "s"}


def _require_bs4():
    try:
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise MissingDependencyError("beautifulsoup4", "read EPUB/HTML content") from exc
    return BeautifulSoup


def _bs_features() -> str:
    try:
        import lxml  # noqa: F401
        return "lxml"
    except ImportError:
        return "html.parser"


def html_to_blocks(node, blocks: list[Block], resolve_image: Callable[[str], str | None]) -> None:
    """Recursively walk an HTML tree in document order, emitting blocks."""
    from bs4 import Comment, NavigableString, Tag

    loose: list[str] = []

    def flush_loose() -> None:
        text = clean_text("".join(loose))
        loose.clear()
        if text:
            blocks.append(Block("text", text))

    def emit_images(tag: Tag) -> None:
        for img in tag.find_all(["img", "image"]):
            src = img.get("src") or img.get("xlink:href") or img.get("href")
            rel = resolve_image(src) if src else None
            if rel:
                blocks.append(Block("image", path=rel))

    for child in node.children:
        if isinstance(child, Comment):
            continue
        if isinstance(child, NavigableString):
            loose.append(str(child))
            continue
        if not isinstance(child, Tag):
            continue
        name = (child.name or "").lower()
        if name in _SKIP_TAGS:
            continue
        if name in _INLINE_TAGS and not child.find(["img", "image", "p", "div"]):
            loose.append(child.get_text())
            continue
        flush_loose()
        if name in _HEADING_TAGS:
            text = clean_text(child.get_text())
            if text:
                blocks.append(Block("heading", text, _HEADING_TAGS[name]))
            emit_images(child)
        elif name in ("img", "image"):
            src = child.get("src") or child.get("xlink:href") or child.get("href")
            rel = resolve_image(src) if src else None
            if rel:
                blocks.append(Block("image", path=rel))
        elif name in _TEXT_BLOCK_TAGS:
            if name == "pre":
                text = child.get_text().strip("\n")
            else:
                text = clean_text(child.get_text())
                if name == "li" and text:
                    text = f"• {text}"
            if text:
                blocks.append(Block("text", text))
            emit_images(child)
        elif name == "table":
            rows = []
            for tr in child.find_all("tr"):
                cells = [clean_text(td.get_text()) for td in tr.find_all(["td", "th"])]
                if any(cells):
                    rows.append(" | ".join(cells))
            if rows:
                blocks.append(Block("text", "\n".join(rows)))
            emit_images(child)
        else:  # div, section, blockquote, ul, figure, body, ...
            html_to_blocks(child, blocks, resolve_image)
    flush_loose()


# ---------------------------------------------------------------------------
# EPUB
# ---------------------------------------------------------------------------
def _epub_meta(book, name: str) -> str | None:
    try:
        values = book.get_metadata("DC", name)
        return str(values[0][0]).strip() if values else None
    except Exception:
        return None


def _parse_epub(path: Path, media: MediaWriter):
    try:
        import ebooklib
        from ebooklib import epub
    except ImportError as exc:
        raise MissingDependencyError("EbookLib", "import EPUB files") from exc
    BeautifulSoup = _require_bs4()

    try:
        book = epub.read_epub(str(path), {"ignore_ncx": True})
    except Exception as exc:
        raise ParseError(f"Could not open EPUB (corrupted or DRM-protected?): {exc}") from exc

    blocks: list[Block] = []
    warnings: list[str] = []
    cache: dict[str, str | None] = {}
    features = _bs_features()

    for entry in book.spine:
        idref = entry[0] if isinstance(entry, (tuple, list)) else entry
        item = book.get_item_with_id(idref)
        if item is None or item.get_type() != ebooklib.ITEM_DOCUMENT or isinstance(item, epub.EpubNav):
            continue
        try:
            soup = BeautifulSoup(item.get_content(), features)
        except Exception as exc:
            warnings.append(f"Skipped unreadable chapter {item.get_name()}: {exc}")
            continue
        for br in soup.find_all("br"):
            br.replace_with("\n")
        base = posixpath.dirname(item.get_name())

        def resolve(src: str, base: str = base) -> str | None:
            href = unquote(src.split("#")[0].split("?")[0])
            if not href or urlparse(href).scheme:
                return None
            href = posixpath.normpath(posixpath.join(base, href))
            if href not in cache:
                img_item = book.get_item_with_href(href)
                cache[href] = (
                    media.save(img_item.get_content(), posixpath.splitext(href)[1]) if img_item else None
                )
            return cache[href]

        blocks.append(Block("break"))  # each spine file is its own unit
        html_to_blocks(soup.body or soup, blocks, resolve)

    return _epub_meta(book, "title"), _epub_meta(book, "creator"), blocks, warnings


# ---------------------------------------------------------------------------
# DOCX / DOC
# ---------------------------------------------------------------------------
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_V = "urn:schemas-microsoft-com:vml"
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx_images(element, part, media: MediaWriter) -> list[Block]:
    """Images inside a paragraph or table: modern DrawingML and legacy VML."""
    out = []
    for el in element.iter(f"{{{_NS_A}}}blip", f"{{{_NS_V}}}imagedata"):
        rid = el.get(f"{{{_NS_R}}}embed") or el.get(f"{{{_NS_R}}}id")
        if not rid:
            continue
        try:
            target = part.related_parts[rid]
        except KeyError:  # externally linked image
            continue
        rel = media.save(target.blob, posixpath.splitext(str(target.partname))[1])
        if rel:
            out.append(Block("image", path=rel))
    return out


def _docx_heading_level(paragraph) -> int:
    try:
        style = paragraph.style
    except Exception:
        style = None
    depth = 0
    while style is not None and depth < 6:
        name = (style.name or "").lower()
        m = re.match(r"heading\s*(\d)", name)
        if m:
            return int(m.group(1))
        if name == "title":
            return 1
        style = style.base_style
        depth += 1
    levels = paragraph._p.xpath("./w:pPr/w:outlineLvl/@w:val")
    if levels:
        try:
            level = int(levels[0])
            if level < 9:
                return level + 1
        except ValueError:
            pass
    return 0


def _parse_docx(path: Path, media: MediaWriter):
    try:
        import docx
        from docx.table import Table
    except ImportError as exc:
        raise MissingDependencyError("python-docx", "import Word documents") from exc

    try:
        document = docx.Document(str(path))
    except Exception as exc:
        raise ParseError(f"Could not open the Word file (corrupted, or not a real .docx?): {exc}") from exc

    part = document.part
    blocks: list[Block] = []
    for item in document.iter_inner_content():
        if isinstance(item, Table):
            rows = []
            for tr in item._tbl.iter(f"{{{_NS_W}}}tr"):
                cells = []
                for tc in tr.iter(f"{{{_NS_W}}}tc"):
                    cells.append(clean_text("".join(t.text or "" for t in tc.iter(f"{{{_NS_W}}}t"))))
                if any(cells):
                    rows.append(" | ".join(cells))
            if rows:
                blocks.append(Block("text", "\n".join(rows)))
            blocks.extend(_docx_images(item._tbl, part, media))
            continue

        text = clean_text(item.text, keep_newlines=True)
        level = _docx_heading_level(item) if text else 0
        if level:
            blocks.append(Block("heading", text, level))
        elif text:
            try:
                style_name = (item.style.name or "") if item.style is not None else ""
            except Exception:
                style_name = ""
            if "list" in style_name.lower():
                text = f"• {text}"
            blocks.append(Block("text", text))
        blocks.extend(_docx_images(item._p, part, media))

    props = document.core_properties
    return props.title or None, props.author or None, blocks, []


def _find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.isfile(candidate):
            return candidate
    return None


def convert_doc_to_docx(path: Path, out_dir: Path) -> Path:
    """Convert a legacy binary .doc file using LibreOffice or Microsoft Word."""
    errors: list[str] = []
    out = out_dir / (path.stem + ".docx")

    soffice = _find_soffice()
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "docx", "--outdir", str(out_dir), str(path)],
                check=True, timeout=180, capture_output=True,
            )
            if out.is_file():
                return out
            errors.append("LibreOffice produced no output")
        except (subprocess.SubprocessError, OSError) as exc:
            errors.append(f"LibreOffice: {exc}")
    else:
        errors.append("LibreOffice not found")

    if os.name == "nt":
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()  # we are usually on a worker thread
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
                doc.SaveAs2(str(out.resolve()), FileFormat=16)  # wdFormatDocumentDefault
                doc.Close(False)
            finally:
                word.Quit()
            if out.is_file():
                return out
            errors.append("Microsoft Word produced no output")
        except ImportError:
            errors.append("pywin32 not installed (needed to drive Microsoft Word)")
        except Exception as exc:
            errors.append(f"Microsoft Word: {exc}")

    raise ParseError(
        "Legacy .doc files must be converted to .docx first. Install LibreOffice, or on Windows "
        "with Microsoft Word run 'pip install pywin32'. You can also re-save the file as .docx.\n"
        "Details: " + "; ".join(errors)
    )


def _parse_doc(path: Path, media: MediaWriter):
    with tempfile.TemporaryDirectory() as tmp:
        converted = convert_doc_to_docx(path, Path(tmp))
        title, author, blocks, warnings = _parse_docx(converted, media)
    return title, author, blocks, warnings + ["Converted from legacy .doc format"]


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
_PAGE_NUMBER_RE = re.compile(r"(page\s*)?\d{1,4}(\s*(of|/)\s*\d{1,4})?", re.IGNORECASE)


def _join_pdf_lines(lines: list[str]) -> str:
    """Join visual lines into a paragraph, removing end-of-line hyphenation."""
    out = ""
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if not out:
            out = line
        elif out.endswith("-") and line[:1].islower():
            out = out[:-1] + line
        else:
            out += " " + line
    return clean_text(out)


def _pdf_font_profile(doc, pymupdf, sample_pages: int = 80) -> Callable[[float], int]:
    """Infer heading levels from font sizes.

    The most common size (weighted by characters) is body text. Sizes about
    15% larger map to heading levels 1-3, largest first.
    """
    sizes: Counter[int] = Counter()
    flags = getattr(pymupdf, "TEXTFLAGS_TEXT", 0)
    for pno in range(min(doc.page_count, sample_pages)):
        try:
            data = doc.load_page(pno).get_text("dict", flags=flags)
        except Exception:
            continue
        for block in data.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if text:
                        sizes[round(span.get("size", 0))] += len(text)
    if not sizes:
        return lambda size: 0
    body = sizes.most_common(1)[0][0]
    threshold = body * 1.15 + 0.5
    candidates = sorted((s for s in sizes if s >= threshold), reverse=True)

    def level_for(size: float) -> int:
        r = round(size)
        if r < threshold:
            return 0
        return min(3, 1 + sum(1 for c in candidates if c > r))

    return level_for


def _pdf_text_block(block: dict, level_for: Callable[[float], int] | None) -> list[Block]:
    lines: list[tuple[int, str]] = []
    for line in block.get("lines", []):
        spans = [s for s in line.get("spans", []) if s.get("text")]
        text = "".join(s["text"] for s in spans).strip()
        if not text:
            continue
        size = max((s.get("size", 0) for s in spans if s["text"].strip()), default=0)
        is_heading_shape = len(text) <= 150 and sum(c.isalpha() for c in text) >= 2
        level = level_for(size) if (level_for and is_heading_shape) else 0
        lines.append((level, text))

    if len(lines) == 1 and _PAGE_NUMBER_RE.fullmatch(lines[0][1]):
        return []  # bare page number

    out: list[Block] = []
    group: list[str] = []
    group_level = 0
    for level, text in lines + [(-1, "")]:  # sentinel flushes the last group
        if group and level != group_level:
            if group_level > 0:
                out.append(Block("heading", " ".join(group), group_level))
            else:
                joined = _join_pdf_lines(group)
                if joined:
                    out.append(Block("text", joined))
            group = []
        group_level = level
        if text:
            group.append(text)
    return out


def _norm(text: str) -> str:
    return re.sub(r"\W+", "", text).lower()


def _inject_toc_headings(page_blocks: list[Block], entries: list[tuple[int, str]]) -> list[Block]:
    """Place bookmark headings before the block where their title appears."""
    if not entries:
        return page_blocks
    inserts: dict[int, list[Block]] = {}
    drop: set[int] = set()
    for level, title in entries:
        key = _norm(title)[:60]
        idx = 0
        if key:
            for i, b in enumerate(page_blocks):
                if b.kind in ("text", "heading") and i not in drop and _norm(b.text).startswith(key):
                    idx = i
                    if _norm(b.text) == _norm(title):
                        drop.add(i)  # the block *is* the heading: don't repeat it
                    break
        inserts.setdefault(idx, []).append(Block("heading", title.strip(), level))
    out: list[Block] = []
    for i, b in enumerate(page_blocks):
        out.extend(inserts.pop(i, []))
        if i not in drop:
            out.append(b)
    for remaining in inserts.values():  # page had no blocks at all
        out.extend(remaining)
    return out


def _parse_pdf(path: Path, media: MediaWriter):
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf  # older PyMuPDF releases
        except ImportError:
            return _parse_pdf_pypdf(path, media)

    try:
        doc = pymupdf.open(str(path))
    except Exception as exc:
        raise ParseError(f"Could not open PDF (corrupted or not a PDF?): {exc}") from exc

    media.skip_repeats = True
    warnings: list[str] = []
    try:
        if doc.needs_pass and not doc.authenticate(""):
            raise ParseError("This PDF is password-protected.")
        meta = doc.metadata or {}

        toc: dict[int, list[tuple[int, str]]] = {}
        try:
            raw_toc = doc.get_toc(simple=True)
        except Exception:
            raw_toc = []
        if len(raw_toc) >= 2:  # bookmarks are more reliable than font sizes
            for level, title, page in raw_toc:
                if page >= 1 and title and title.strip():
                    toc.setdefault(page - 1, []).append((level, title))
        level_for = None if toc else _pdf_font_profile(doc, pymupdf)

        blocks: list[Block] = []
        text_found = False
        for pno in range(doc.page_count):
            try:
                page_dict = doc.load_page(pno).get_text("dict", sort=True)
            except Exception as exc:
                warnings.append(f"Skipped unreadable page {pno + 1}: {exc}")
                continue
            page_blocks: list[Block] = []
            for b in page_dict.get("blocks", []):
                if b.get("type") == 1:  # image block
                    rel = media.save(b.get("image") or b"", b.get("ext", "png"), b.get("width"), b.get("height"))
                    if rel:
                        page_blocks.append(Block("image", path=rel))
                else:
                    page_blocks.extend(_pdf_text_block(b, level_for))
            text_found = text_found or any(b.kind == "text" for b in page_blocks)
            blocks.extend(_inject_toc_headings(page_blocks, toc.get(pno, [])))

        if not text_found:
            warnings.append("No selectable text found. This looks like a scanned PDF (run OCR first for text).")
        return meta.get("title") or None, meta.get("author") or None, blocks, warnings
    finally:
        doc.close()


def _parse_pdf_pypdf(path: Path, media: MediaWriter):
    """Fallback when PyMuPDF is unavailable: text plus page images, no layout."""
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise MissingDependencyError("pymupdf", "import PDF files") from exc
    try:
        reader = PdfReader(str(path))
        if reader.is_encrypted and not reader.decrypt(""):
            raise ParseError("This PDF is password-protected.")
    except ParseError:
        raise
    except Exception as exc:
        raise ParseError(f"Could not open PDF (corrupted or not a PDF?): {exc}") from exc

    media.skip_repeats = True
    warnings = ["PyMuPDF not installed: basic extraction only (no heading detection)."]
    blocks: list[Block] = []
    for pno, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            warnings.append(f"Skipped unreadable page {pno + 1}: {exc}")
            continue
        for para in re.split(r"\n\s*\n", text):
            joined = _join_pdf_lines(para.splitlines())
            if joined and not _PAGE_NUMBER_RE.fullmatch(joined):
                blocks.append(Block("text", joined))
        try:
            for img in page.images:
                rel = media.save(img.data, Path(img.name).suffix)
                if rel:
                    blocks.append(Block("image", path=rel))
        except Exception as exc:
            warnings.append(f"Could not extract images on page {pno + 1}: {exc}")
    meta = reader.metadata
    return (meta.title if meta else None), (meta.author if meta else None), blocks, warnings


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
_HANDLERS = {
    ".pdf": ("pdf", _parse_pdf),
    ".epub": ("epub", _parse_epub),
    ".docx": ("docx", _parse_docx),
    ".doc": ("doc", _parse_doc),
    ".txt": ("txt", _parse_txt),
    ".md": ("md", _parse_md),
    ".markdown": ("md", _parse_md),
}


def parse_file(
    path: str | Path,
    media_dir: str | Path,
    rel_prefix: str,
    max_chars: int = DEFAULT_MAX_CHARS,
) -> ParsedDocument:
    """Parse ``path`` into sections, writing images under ``media_dir``.

    ``rel_prefix`` is the media folder's path relative to the data directory,
    which is what the image paths in the result are built from.

    Raises ``ParseError`` (or its subclass ``MissingDependencyError``) with a
    readable message for every failure. If parsing fails, the caller should
    delete ``media_dir``.
    """
    path = Path(path)
    if not path.is_file():
        raise ParseError(f"File not found: {path}")
    ext = path.suffix.lower()
    if ext not in _HANDLERS:
        raise ParseError(f"Unsupported file type '{ext}'. Supported: {', '.join(SUPPORTED_EXTENSIONS)}")
    if path.stat().st_size == 0:
        raise ParseError(f"{path.name} is empty.")

    fmt, handler = _HANDLERS[ext]
    media = MediaWriter(Path(media_dir), rel_prefix)
    try:
        title, author, blocks, warnings = handler(path, media)
    except ParseError:
        raise
    except MemoryError as exc:
        raise ParseError(f"{path.name} is too large to process.") from exc
    except Exception as exc:
        raise ParseError(f"Could not parse {path.name}: {type(exc).__name__}: {exc}") from exc

    title = clean_text(title or "") or _title_from_filename(path)
    sections = build_sections(blocks, title, max_chars)
    if not sections:
        raise ParseError(f"No readable text or images were found in {path.name}.")
    return ParsedDocument(title, clean_text(author or "") or None, fmt, sections, warnings)
