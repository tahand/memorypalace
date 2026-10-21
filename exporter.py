"""
exporter.py - Compile knowledge-base sections into PDF, DOCX, Markdown or TXT.

The input is a list of ``ExportGroup`` objects (see ``db.get_export_groups``):
one group per document, each holding its sections in reading order, with
absolute image paths. The functions here only read that data and never touch
the database, so they can safely run on a worker thread.
"""
from __future__ import annotations

import html
import re
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Sequence

EXPORT_FORMATS = {"pdf": ".pdf", "docx": ".docx", "md": ".md", "txt": ".txt"}

ProgressFn = Callable[[int, int], None]


class ExportError(Exception):
    """Raised with a user-readable message when an export cannot be produced."""


@dataclass
class ExportOptions:
    title: str = "Memory Palace Export"
    include_images: bool = True
    include_tags: bool = True
    include_headings: bool = True


# Characters that are illegal in XML (and therefore in DOCX) or break ReportLab.
_XML_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff\ufffe\uffff]")


def _safe(text: str) -> str:
    return _XML_ILLEGAL.sub("", text or "")


def _paragraphs(text: str) -> list[str]:
    return [p.strip() for p in re.split(r"\n\s*\n", _safe(text)) if p.strip()]


def _count(groups: Sequence) -> int:
    return sum(len(g.sections) for g in groups)


def _prepare_image(path: Path, tmpdir: Path, allowed: set[str]) -> tuple[Path, int, int] | None:
    """Return (path, width_px, height_px) in a format the target accepts.

    Unsupported formats (WebP, TIFF for PDF, ...) are converted to PNG in
    ``tmpdir``. Returns None if the image cannot be read (e.g. SVG).
    """
    if not path.is_file():
        return None
    try:
        from PIL import Image
    except ImportError:  # no Pillow: pass through formats we know are fine
        return (path, 0, 0) if path.suffix.lower().lstrip(".") in allowed else None
    try:
        with Image.open(path) as im:
            fmt = (im.format or "").lower()
            width, height = im.size
            if fmt in allowed:
                return path, width, height
            if im.mode not in ("RGB", "RGBA", "L", "LA", "P"):
                im = im.convert("RGB")
            out = tmpdir / f"{path.parent.name}_{path.stem}.png"
            im.save(out, "PNG")
            return out, width, height
    except Exception:
        return None


def _default_out_path(out_path: str | Path, fmt: str | None) -> tuple[Path, str]:
    out = Path(out_path)
    fmt = (fmt or out.suffix.lstrip(".")).lower()
    if fmt == "markdown":
        fmt = "md"
    if fmt not in EXPORT_FORMATS:
        raise ExportError(f"Unsupported export format '{fmt}'. Choose one of: {', '.join(EXPORT_FORMATS)}")
    if out.suffix.lower() != EXPORT_FORMATS[fmt]:
        out = out.with_suffix(EXPORT_FORMATS[fmt])
    return out, fmt


def export(
    groups: Sequence,
    out_path: str | Path,
    fmt: str | None = None,
    options: ExportOptions | None = None,
    progress: ProgressFn | None = None,
) -> Path:
    """Write ``groups`` to ``out_path``. Returns the path actually written."""
    out, fmt = _default_out_path(out_path, fmt)
    options = options or ExportOptions()
    if not groups or _count(groups) == 0:
        raise ExportError("Nothing to export: no sections match the current selection.")
    handler = {"pdf": _export_pdf, "docx": _export_docx, "md": _export_markdown, "txt": _export_txt}[fmt]
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        handler(groups, out, options, progress or (lambda done, total: None))
    except ExportError:
        raise
    except PermissionError as exc:
        raise ExportError(f"Cannot write to {out}. Is the file open in another program?") from exc
    except Exception as exc:
        raise ExportError(f"Export failed: {type(exc).__name__}: {exc}") from exc
    return out


def _subtitle(groups: Sequence) -> str:
    n_docs, n_secs = len(groups), _count(groups)
    return (
        f"Exported {datetime.now():%Y-%m-%d %H:%M} · "
        f"{n_docs} document{'s' * (n_docs != 1)} · {n_secs} section{'s' * (n_secs != 1)}"
    )


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------
def _export_docx(groups, out: Path, opts: ExportOptions, progress: ProgressFn) -> None:
    try:
        import docx
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt, RGBColor
    except ImportError as exc:
        raise ExportError("DOCX export needs python-docx:  pip install python-docx") from exc

    document = docx.Document()
    document.core_properties.title = _safe(opts.title)
    document.add_heading(_safe(opts.title), 0)
    sub = document.add_paragraph(_subtitle(groups))
    sub.runs[0].italic = True

    page = document.sections[0]
    max_w_in = (page.page_width - page.left_margin - page.right_margin) / 914400  # EMU -> inch
    max_h_in = (page.page_height - page.top_margin - page.bottom_margin) / 914400 * 0.75
    total, done = _count(groups), 0

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for gi, group in enumerate(groups):
            if gi:
                document.add_page_break()
            document.add_heading(_safe(group.document["title"]), 1)
            if group.document.get("author"):
                p = document.add_paragraph(_safe(group.document["author"]))
                p.runs[0].italic = True

            for section in group.sections:
                if opts.include_headings and section.heading:
                    document.add_heading(_safe(section.heading), 2)
                for para in _paragraphs(section.content_text):
                    document.add_paragraph(para)

                if opts.include_images:
                    for img_path in section.image_paths:
                        prepared = _prepare_image(img_path, tmpdir, {"png", "jpeg", "gif", "bmp", "tiff"})
                        if not prepared:
                            document.add_paragraph(f"[image not embeddable: {img_path.name}]")
                            continue
                        path, w_px, h_px = prepared
                        # Assume 96 dpi; never upscale; fit inside the page.
                        w_in = min(max_w_in, (w_px or 480) / 96)
                        if w_px and h_px and w_in * h_px / w_px > max_h_in:
                            w_in = max_h_in * w_px / h_px
                        try:
                            document.add_picture(str(path), width=Inches(w_in))
                            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        except Exception:
                            document.add_paragraph(f"[image could not be embedded: {img_path.name}]")

                if opts.include_tags and section.tags:
                    p = document.add_paragraph()
                    run = p.add_run("Tags: " + "  ".join(f"#{t}" for t in section.tags))
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

                done += 1
                progress(done, total)
        document.save(str(out))


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
_FONT_CANDIDATES = [
    # (regular, bold, italic)
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/ariali.ttf"),
    ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/segoeuii.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
    ("/usr/share/fonts/TTF/DejaVuSans.ttf", "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/TTF/DejaVuSans-Oblique.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
     "/System/Library/Fonts/Supplemental/Arial Italic.ttf"),
]


def _register_fonts() -> tuple[str, str, str]:
    """Register a Unicode TrueType font if one exists (Helvetica is Latin-1 only)."""
    from reportlab.lib.fonts import addMapping
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    if "MP-Regular" in pdfmetrics.getRegisteredFontNames():
        return "MP-Regular", "MP-Bold", "MP-Italic"
    for regular, bold, italic in _FONT_CANDIDATES:
        if not Path(regular).is_file():
            continue
        try:
            pdfmetrics.registerFont(TTFont("MP-Regular", regular))
            pdfmetrics.registerFont(TTFont("MP-Bold", bold if Path(bold).is_file() else regular))
            pdfmetrics.registerFont(TTFont("MP-Italic", italic if Path(italic).is_file() else regular))
            addMapping("MP-Regular", 0, 0, "MP-Regular")
            addMapping("MP-Regular", 1, 0, "MP-Bold")
            addMapping("MP-Regular", 0, 1, "MP-Italic")
            addMapping("MP-Regular", 1, 1, "MP-Bold")
            return "MP-Regular", "MP-Bold", "MP-Italic"
        except Exception:
            continue
    return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


def _rl(text: str) -> str:
    """Escape text for a ReportLab Paragraph (mini-HTML markup)."""
    return html.escape(_safe(text), quote=False).replace("\n", "<br/>")


def _export_pdf(groups, out: Path, opts: ExportOptions, progress: ProgressFn) -> None:
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer
    except ImportError as exc:
        raise ExportError("PDF export needs reportlab:  pip install reportlab") from exc

    regular, bold, italic = _register_fonts()
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle("MPTitle", parent=base["Title"], fontName=bold, fontSize=24, leading=30),
        "subtitle": ParagraphStyle("MPSub", parent=base["Normal"], fontName=italic, fontSize=10,
                                   textColor=colors.grey, alignment=TA_CENTER, spaceAfter=18),
        "h1": ParagraphStyle("MPH1", parent=base["Heading1"], fontName=bold, fontSize=18, leading=22, spaceAfter=6),
        "author": ParagraphStyle("MPAuthor", parent=base["Normal"], fontName=italic, textColor=colors.grey,
                                 spaceAfter=12),
        "h2": ParagraphStyle("MPH2", parent=base["Heading2"], fontName=bold, fontSize=13, leading=16,
                             spaceBefore=12, spaceAfter=4, textColor=colors.HexColor("#1f4e79")),
        "body": ParagraphStyle("MPBody", parent=base["BodyText"], fontName=regular, fontSize=10.5, leading=14.5,
                               spaceAfter=6),
        "tags": ParagraphStyle("MPTags", parent=base["Normal"], fontName=regular, fontSize=8,
                               textColor=colors.grey, spaceAfter=6),
    }

    class _Doc(SimpleDocTemplate):
        """Adds PDF bookmarks (outline) for documents and section headings."""

        def afterFlowable(self, flowable):
            bookmark = getattr(flowable, "_mp_bookmark", None)
            if bookmark:
                level, text, key = bookmark
                self.canv.bookmarkPage(key)
                self.canv.addOutlineEntry(text, key, level=level, closed=level > 0)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(regular, 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"{_safe(opts.title)} · page {doc.page}")
        canvas.restoreState()

    doc = _Doc(
        str(out), pagesize=A4, title=_safe(opts.title), author="Memory Palace",
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
    )
    max_w, max_h = doc.width, doc.height * 0.7

    def heading(text: str, style: str, level: int) -> Paragraph:
        p = Paragraph(_rl(text), styles[style])
        p._mp_bookmark = (level, _safe(text)[:120], f"bm{id(p)}")
        return p

    story = [Paragraph(_rl(opts.title), styles["title"]), Paragraph(_rl(_subtitle(groups)), styles["subtitle"])]
    total, done = _count(groups), 0
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for gi, group in enumerate(groups):
            if gi:
                story.append(PageBreak())
            story.append(heading(group.document["title"], "h1", 0))
            if group.document.get("author"):
                story.append(Paragraph(_rl(group.document["author"]), styles["author"]))

            for section in group.sections:
                if opts.include_headings and section.heading:
                    story.append(heading(section.heading, "h2", 1))
                for para in _paragraphs(section.content_text):
                    story.append(Paragraph(_rl(para), styles["body"]))

                if opts.include_images:
                    for img_path in section.image_paths:
                        prepared = _prepare_image(img_path, tmpdir, {"png", "jpeg", "gif"})
                        if not prepared:
                            story.append(Paragraph(_rl(f"[image not embeddable: {img_path.name}]"), styles["tags"]))
                            continue
                        path, w_px, h_px = prepared
                        w_px, h_px = w_px or 400, h_px or 300
                        # 96 dpi to points, never upscale, fit inside the frame.
                        scale = min(1.0, max_w / (w_px * 0.75), max_h / (h_px * 0.75))
                        img = Image(str(path), width=w_px * 0.75 * scale, height=h_px * 0.75 * scale)
                        img.hAlign = "CENTER"
                        story += [Spacer(1, 4), img, Spacer(1, 8)]

                if opts.include_tags and section.tags:
                    story.append(Paragraph(_rl("Tags: " + "  ".join(f"#{t}" for t in section.tags)), styles["tags"]))

                done += 1
                progress(done, total)
        doc.build(story, onFirstPage=footer, onLaterPages=footer)


# ---------------------------------------------------------------------------
# Markdown / TXT
# ---------------------------------------------------------------------------
def _export_markdown(groups, out: Path, opts: ExportOptions, progress: ProgressFn) -> None:
    media_dir = out.with_name(out.stem + "_media")
    lines = [f"# {opts.title}", "", f"*{_subtitle(groups)}*", ""]
    total, done, copied = _count(groups), 0, 0

    for group in groups:
        lines += [f"## {group.document['title']}", ""]
        if group.document.get("author"):
            lines += [f"*{group.document['author']}*", ""]
        for section in group.sections:
            if opts.include_headings and section.heading:
                lines += [f"### {section.heading}", ""]
            for para in _paragraphs(section.content_text):
                lines += [para, ""]
            if opts.include_images:
                for img in section.image_paths:
                    media_dir.mkdir(exist_ok=True)
                    copied += 1
                    target = media_dir / f"{copied:04d}_{img.name}"
                    shutil.copy2(img, target)
                    lines += [f"![{img.stem}]({media_dir.name}/{target.name})", ""]
            if opts.include_tags and section.tags:
                lines += [" ".join(f"`#{t}`" for t in section.tags), ""]
            done += 1
            progress(done, total)
    out.write_text("\n".join(lines), encoding="utf-8")


def _export_txt(groups, out: Path, opts: ExportOptions, progress: ProgressFn) -> None:
    lines = [opts.title, "=" * len(opts.title), _subtitle(groups), ""]
    total, done = _count(groups), 0
    for group in groups:
        title = group.document["title"]
        lines += ["", title.upper(), "-" * len(title), ""]
        for section in group.sections:
            if opts.include_headings and section.heading:
                lines += [f"## {section.heading}", ""]
            for para in _paragraphs(section.content_text):
                lines += [para, ""]
            if opts.include_images:
                lines += [f"[Image: {img}]" for img in section.image_paths]
            if opts.include_tags and section.tags:
                lines += ["Tags: " + ", ".join(section.tags)]
            lines.append("")
            done += 1
            progress(done, total)
    out.write_text("\n".join(lines), encoding="utf-8")
