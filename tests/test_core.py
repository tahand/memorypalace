"""Core tests: SM-2, chunking, every parser, tagging, review queue, exports.

Run with:  python -m pytest -q
Fixture documents are generated on the fly with the same libraries the app
uses, so no binary test files live in the repository.
"""
from __future__ import annotations

import io
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import db as dbmod  # noqa: E402
import exporter  # noqa: E402
import parser as mp_parser  # noqa: E402
from db import KnowledgeBase, ReviewState, sm2  # noqa: E402


# -- helpers ---------------------------------------------------------------
def png_bytes(color=(200, 30, 30), size=(120, 80)) -> bytes:
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", size, color).save(buf, "PNG")
    return buf.getvalue()


@pytest.fixture
def kb(tmp_path):
    base = KnowledgeBase(tmp_path / "data")
    yield base
    base.close()


def parse_into(kb: KnowledgeBase, path: Path):
    media_abs, media_rel = kb.new_media_dir()
    parsed = mp_parser.parse_file(path, media_abs, media_rel)
    doc_id, _ = kb.add_document(parsed, str(path), None, media_rel)
    return doc_id, parsed


# -- SM-2 ------------------------------------------------------------------
def test_sm2_progression_and_reset():
    s = sm2(ReviewState(), 4)
    assert (s.interval, s.repetitions) == (1, 1)
    s = sm2(s, 4)
    assert (s.interval, s.repetitions) == (6, 2)
    s = sm2(s, 4)
    assert s.interval == round(6 * 2.5) and s.repetitions == 3
    failed = sm2(s, 1)
    assert failed.repetitions == 0 and failed.interval == 0
    assert failed.ease_factor < s.ease_factor


def test_sm2_ease_floor():
    s = ReviewState()
    for _ in range(20):
        s = sm2(s, 0)
    assert s.ease_factor == dbmod.MIN_EASE


# -- chunking --------------------------------------------------------------
def test_build_sections_split_and_breadcrumb():
    B = mp_parser.Block
    blocks = [
        B("heading", "Biology", 1),
        B("heading", "Cells", 2),
        B("text", "A" * 900),
        B("image", path="media/x/img.png"),
        B("text", "B" * 900),
        B("text", "C" * 50),
        B("heading", "Genetics", 2),
        B("text", "DNA stuff."),
    ]
    secs = mp_parser.build_sections(blocks, "Fallback", max_chars=1000)
    assert [s.heading for s in secs] == ["Biology › Cells", "Biology › Cells (part 2)", "Biology › Genetics"]
    assert secs[0].image_paths == ["media/x/img.png"]
    assert secs[1].content_text.startswith("B")


def test_long_paragraph_is_split_on_sentences():
    text = " ".join(f"Sentence number {i} is here." for i in range(200))
    pieces = mp_parser._split_long_paragraph(text, 500)
    assert all(len(p) <= 500 for p in pieces)
    assert " ".join(pieces) == text


# -- parsers ---------------------------------------------------------------
def test_markdown_parser_with_image(tmp_path, kb):
    (tmp_path / "pic.png").write_bytes(png_bytes())
    md = tmp_path / "notes.md"
    md.write_text(
        "---\ntitle: My Notes\nauthor: Me\n---\n# Physics\n\nIntro text.\n\n## Newton\n\n"
        "F = ma is Newton's second law.\n\n![diagram](pic.png)\n\n```python\n# not a heading\nx = 1\n```\n",
        encoding="utf-8",
    )
    doc_id, parsed = parse_into(kb, md)
    assert parsed.title == "My Notes" and parsed.author == "Me"
    newton = [s for s in parsed.sections if "Newton" in s.heading][0]
    assert len(newton.image_paths) == 1
    assert "# not a heading" in newton.content_text
    assert kb.abs_path(newton.image_paths[0]).is_file()


def test_txt_parser_chapters(tmp_path, kb):
    txt = tmp_path / "story.txt"
    txt.write_text(
        "CHAPTER 1\n\nIt was a dark night.\nThe wind howled.\n\n"
        "Part of the reason was the storm.\n\nChapter 2: Morning\n\nThe sun rose.",
        encoding="utf-8",
    )
    _, parsed = parse_into(kb, txt)
    headings = [s.heading for s in parsed.sections]
    assert headings == ["CHAPTER 1", "Chapter 2: Morning"]
    assert "Part of the reason" in parsed.sections[0].content_text  # not mistaken for a heading


def test_docx_parser_headings_tables_images(tmp_path, kb):
    import docx
    from docx.shared import Inches
    d = docx.Document()
    d.core_properties.title = "Word Book"
    d.add_heading("Chapter One", 1)
    d.add_paragraph("First paragraph about mitochondria.")
    (tmp_path / "a.png").write_bytes(png_bytes())
    d.add_picture(str(tmp_path / "a.png"), width=Inches(1))
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "key", "value"
    d.add_heading("Chapter Two", 1)
    d.add_paragraph("Second chapter text.")
    path = tmp_path / "book.docx"
    d.save(path)

    _, parsed = parse_into(kb, path)
    assert parsed.title == "Word Book"
    assert [s.heading for s in parsed.sections] == ["Chapter One", "Chapter Two"]
    assert len(parsed.sections[0].image_paths) == 1
    assert "key | value" in parsed.sections[0].content_text


def test_epub_parser(tmp_path, kb):
    from ebooklib import epub
    book = epub.EpubBook()
    book.set_identifier("id1")
    book.set_title("Epub Title")
    book.add_author("An Author")
    img = epub.EpubItem(uid="img1", file_name="images/pic.png", media_type="image/png", content=png_bytes())
    book.add_item(img)
    c1 = epub.EpubHtml(title="One", file_name="text/ch1.xhtml")
    c1.content = (
        "<html><body><h1>Chapter 1</h1><p>Hello <em>world</em> of cells.</p>"
        "<p><img src='../images/pic.png'/></p><h2>Sub</h2><ul><li>item one</li></ul></body></html>"
    )
    c2 = epub.EpubHtml(title="Two", file_name="text/ch2.xhtml")
    c2.content = "<html><body><h1>Chapter 2</h1><div>Loose div text.</div></body></html>"
    for c in (c1, c2):
        book.add_item(c)
    book.toc = (c1, c2)
    book.spine = ["nav", c1, c2]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    path = tmp_path / "b.epub"
    epub.write_epub(str(path), book)

    _, parsed = parse_into(kb, path)
    assert parsed.title == "Epub Title" and parsed.author == "An Author"
    by_heading = {s.heading: s for s in parsed.sections}
    assert "Hello world of cells." in by_heading["Chapter 1"].content_text
    assert len(by_heading["Chapter 1"].image_paths) == 1
    assert "• item one" in by_heading["Chapter 1 › Sub"].content_text
    assert "Loose div text." in by_heading["Chapter 2"].content_text


def _make_pdf(path: Path, img_path: Path, with_toc: bool):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setTitle("PDF Title")
    y = 780
    c.setFont("Helvetica-Bold", 22)
    c.drawString(72, y, "Introduction")
    if with_toc:
        c.bookmarkPage("p1")
        c.addOutlineEntry("Introduction", "p1", level=0)
    c.setFont("Helvetica", 11)
    for i in range(30):
        y -= 15
        c.drawString(72, y, f"Body line {i} about photosynthesis and chlorophyll in leaves.")
    c.drawImage(str(img_path), 72, 200, width=150, height=100)
    c.showPage()
    c.setFont("Helvetica-Bold", 22)
    c.drawString(72, 780, "Methods")
    if with_toc:
        c.bookmarkPage("p2")
        c.addOutlineEntry("Methods", "p2", level=0)
    c.setFont("Helvetica", 11)
    c.drawString(72, 760, "We measured things carefully.")
    c.showPage()
    c.save()


@pytest.mark.parametrize("with_toc", [False, True])
def test_pdf_parser(tmp_path, kb, with_toc):
    img = tmp_path / "i.png"
    img.write_bytes(png_bytes(size=(300, 200)))
    path = tmp_path / "doc.pdf"
    _make_pdf(path, img, with_toc)
    _, parsed = parse_into(kb, path)
    assert parsed.title == "PDF Title"
    headings = [s.heading for s in parsed.sections]
    assert headings[0] == "Introduction" and headings[-1] == "Methods"
    assert sum(len(s.image_paths) for s in parsed.sections) == 1
    assert "photosynthesis" in parsed.sections[0].content_text


def test_bad_files(tmp_path):
    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"not really a pdf at all" * 20)
    with pytest.raises(mp_parser.ParseError):
        mp_parser.parse_file(bad, tmp_path / "m", "media/m")
    bad_docx = tmp_path / "bad.docx"
    bad_docx.write_bytes(b"garbage" * 50)
    with pytest.raises(mp_parser.ParseError):
        mp_parser.parse_file(bad_docx, tmp_path / "m", "media/m")
    with pytest.raises(mp_parser.ParseError):
        mp_parser.parse_file(tmp_path / "missing.txt", tmp_path / "m", "media/m")
    unsupported = tmp_path / "x.xyz"
    unsupported.write_text("hi")
    with pytest.raises(mp_parser.ParseError):
        mp_parser.parse_file(unsupported, tmp_path / "m", "media/m")


# -- knowledge base --------------------------------------------------------
def _sample(kb, tmp_path) -> int:
    md = tmp_path / "s.md"
    md.write_text(
        "# Bio\n\n## Cells\n\nThe Mitochondria is the powerhouse.\n\n## Plants\n\nPhotosynthesis in leaves.\n\n"
        "## Misc\n\nmitochondrial DNA is different.\n",
        encoding="utf-8",
    )
    doc_id, _ = parse_into(kb, md)
    return doc_id


def test_tagging_and_autotag(kb, tmp_path):
    doc_id = _sample(kb, tmp_path)
    secs = kb.get_sections(doc_id)
    tag_id = kb.add_tag_to_section(secs[0]["id"], "#mitochondria")
    assert kb.section_tags(secs[0]["id"]) == ["mitochondria"]
    # Whole-word keyword: "mitochondrial" must not match.
    assert kb.count_auto_tag_candidates(tag_id) == 0

    rx = kb.create_tag("mito-any", r"mitochondri\w*", is_regex=True)
    assert kb.auto_tag(rx) == 2
    assert kb.auto_tag(rx) == 0  # idempotent
    assert kb.clear_auto_tags(rx) == 2

    with pytest.raises(ValueError):
        kb.create_tag("broken", "(", is_regex=True)
    with pytest.raises(ValueError):
        kb.create_tag("MITOCHONDRIA")  # case-insensitive uniqueness

    # Tags created earlier are applied automatically to new imports.
    kb.create_tag("leaves")
    md2 = tmp_path / "t.md"
    md2.write_text("# Trees\n\nTrees drop their leaves.\n", encoding="utf-8")
    _, applied = kb.add_document(mp_parser.parse_file(md2, tmp_path / "mm", "media/mm"))
    assert applied >= 1


def test_review_queue_and_undo(kb, tmp_path):
    doc_id = _sample(kb, tmp_path)
    queue = kb.get_review_queue()
    assert len(queue) == 3
    sid = queue[0]
    prev = kb.record_answer(sid, known=True)
    r = kb.get_review(sid)
    assert r["repetitions"] == 1 and r["status"] == "learning"
    assert r["next_review_date"] == (date.today() + timedelta(days=1)).isoformat()
    assert sid not in kb.get_review_queue()

    kb.restore_review(sid, prev)
    assert kb.get_review(sid)["status"] == "new"

    kb.record_answer(queue[1], known=False)
    assert queue[1] in kb.get_review_queue()  # still due today

    kb.soft_delete_section(queue[2])
    assert queue[2] not in kb.get_review_queue()
    kb.restore_section(queue[2])
    assert queue[2] in kb.get_review_queue()

    # Mastery after enough successful reviews.
    t = datetime.now()
    for _ in range(4):
        kb.record_answer(sid, known=True, now=t)
        t += timedelta(days=kb.get_review(sid)["interval"])
    assert kb.get_review(sid)["status"] == "mastered"
    assert kb.stats()["mastered"] == 1


def test_purge_and_delete_document(kb, tmp_path):
    (tmp_path / "p.png").write_bytes(png_bytes())
    md = tmp_path / "img.md"
    md.write_text("# A\n\nText ![x](p.png)\n\n# B\n\nMore\n", encoding="utf-8")
    doc_id, _ = parse_into(kb, md)
    sec = kb.get_sections(doc_id)[0]
    img = kb.abs_path(sec["image_paths"][0])
    assert img.is_file()
    kb.purge_section(sec["id"])
    assert not img.exists()
    assert kb.get_section(sec["id"]) is None
    media_dir = kb.abs_path(kb.get_document(doc_id)["media_dir"])
    kb.delete_document(doc_id)
    assert kb.get_document(doc_id) is None
    assert not media_dir.exists()


# -- export ----------------------------------------------------------------
@pytest.mark.parametrize("fmt", ["pdf", "docx", "md", "txt"])
def test_exports(kb, tmp_path, fmt):
    (tmp_path / "p.png").write_bytes(png_bytes(size=(2000, 1200)))
    md = tmp_path / "e.md"
    md.write_text(
        "# Unicode ✓ café Ωmega\n\nSome <tricky> & text\x0b here.\n\n![x](p.png)\n\n# Second\n\nMore text.\n",
        encoding="utf-8",
    )
    doc_id, _ = parse_into(kb, md)
    secs = kb.get_sections(doc_id)
    kb.add_tag_to_section(secs[0]["id"], "special")
    kb.record_answer(secs[0]["id"], known=True)

    groups = kb.get_export_groups()
    out = exporter.export(groups, tmp_path / f"out.{fmt}", fmt)
    assert out.is_file() and out.stat().st_size > 0

    known = kb.get_export_groups(review_filter="known")
    assert sum(len(g.sections) for g in known) == 1
    picked = kb.get_export_groups(section_ids=[secs[1]["id"]])
    assert [s.id for g in picked for s in g.sections] == [secs[1]["id"]]

    if fmt == "docx":
        import docx
        d = docx.Document(str(out))
        assert len(d.inline_shapes) == 1
    if fmt == "pdf":
        import pymupdf
        with pymupdf.open(str(out)) as d:
            text = "".join(p.get_text() for p in d)
            assert "café" in text and "Ωmega" in text
            assert sum(len(p.get_images()) for p in d) == 1


def test_export_nothing_raises(kb, tmp_path):
    with pytest.raises(exporter.ExportError):
        exporter.export([], tmp_path / "x.pdf")


# -- your own books and cards -----------------------------------------------
def test_manual_book_cards_images_and_settings(kb):
    with pytest.raises(ValueError):
        kb.create_document("   ")
    doc_id = kb.create_document("My notes", "Me")
    doc = kb.get_document(doc_id)
    assert doc["format"] == dbmod.MANUAL_FORMAT and doc["author"] == "Me"

    kb.create_tag("mitochondria")
    with pytest.raises(ValueError):
        kb.add_section(doc_id, "Empty", "  ")
    first = kb.add_section(doc_id, "Q1", "What makes ATP? The mitochondria.")
    rel = kb.save_image(doc_id, png_bytes(), "png")
    second = kb.add_section(doc_id, "", "", [rel])
    sections = kb.get_sections(doc_id)
    assert [s["id"] for s in sections] == [first, second]
    assert [s["sequence_order"] for s in sections] == [0, 1]
    assert kb.get_section(first)["tags"] == ["mitochondria"]       # new cards are auto-tagged
    assert kb.get_section(second)["image_paths"] == [rel] and kb.abs_path(rel).is_file()
    assert set(kb.get_review_queue(document_id=doc_id)) == {first, second}

    # Replacing a card's images deletes files no other card uses.
    kb.update_section(second, "Diagram", "Labelled", [])
    assert kb.get_section(second)["image_paths"] == [] and not kb.abs_path(rel).exists()

    # Splitting a card puts the second half right after it, with the same heading and manual tags.
    third = kb.add_section(doc_id, "Q3", "Last")
    kb.add_tag_to_section(first, "biology")
    with pytest.raises(ValueError):
        kb.split_section(first, "only one half", "  ")
    new = kb.split_section(first, "What makes ATP?", "The mitochondria.")
    assert [s["id"] for s in kb.get_sections(doc_id)] == [first, new, second, third]
    split = kb.get_section(new)
    assert split["heading"] == "Q1" and split["content_text"] == "The mitochondria."
    assert set(split["tags"]) == {"biology", "mitochondria"}
    assert kb.get_section(first)["content_text"] == "What makes ATP?"
    assert new in kb.get_review_queue(document_id=doc_id)

    # A new card can be placed right after an existing one instead of at the end.
    inserted = kb.add_section(doc_id, "Q1b", "Inserted", after=first)
    assert [s["id"] for s in kb.get_sections(doc_id)] == [first, inserted, new, second, third]

    kb.set_setting("tts_rate", "3")
    kb.set_setting("tts_rate", "-2")
    assert kb.get_setting("tts_rate") == "-2" and kb.get_setting("missing", "x") == "x"


def test_split_text_and_ocr_paragraphs():
    sections = mp_parser.split_text("CHAPTER 1 Cells\n\nCells are small.\n\nChapter 2 Tissue\n\nTissues.", "Notes")
    assert [s.heading for s in sections] == ["CHAPTER 1 Cells", "Chapter 2 Tissue"]

    import ocr
    # Two tight lines, then a gap: two paragraphs; the hyphen across lines is joined.
    text = ocr.lines_to_paragraphs([("Plant cells have a cell", 0, 20), ("wall made of cellu-", 26, 20),
                                    ("lose.", 52, 20), ("New idea here", 120, 20)])
    assert text == "Plant cells have a cell wall made of cellulose.\n\nNew idea here"


def test_photo_for_card_rotates_and_shrinks(tmp_path):
    from PIL import Image
    import ocr
    src = tmp_path / "phone.jpg"
    exif = Image.Exif()
    exif[0x0112] = 6  # orientation: rotate 90° clockwise
    Image.new("RGB", (4000, 3000), "white").save(src, "JPEG", exif=exif)
    data, ext = ocr.photo_for_card(src)
    assert ext == "jpg"
    with Image.open(io.BytesIO(data)) as im:
        assert im.size == (1500, 2000)  # upright (portrait) and capped at 2000 px


def test_speakable_text():
    from speech import speakable
    assert speakable("Ch. 2 › Cells", "**Bold** text.") == "Ch. 2. Cells.\n\nBold text"
    assert speakable(None, "") == ""


def test_lines_split_and_remove():
    from lines import heading_path, remove_line, split_lines

    text = ("Cells are small. Every cell has a membrane, e.g. lipids! Plant cells differ?\n\n"
            "1. First item here.\n2. Second item.\n\nMr. Smith said so. The end")
    assert split_lines(text) == [
        "Cells are small.", "Every cell has a membrane, e.g. lipids!", "Plant cells differ?",
        "1. First item here.", "2. Second item.", "Mr. Smith said so.", "The end",
    ]
    # Removing a sentence leaves the rest of the text exactly as it was.
    assert remove_line(text, 2).startswith("Cells are small. Every cell has a membrane, e.g. lipids!\n\n1.")
    assert "1. First item here.\n" not in remove_line(text, 3)
    assert remove_line(text, 3).count("\n\n") == 2
    assert remove_line("Only one.", 0) == ""
    assert heading_path("Ch 2 › Cells (part 3)", "Book") == ("Ch 2", "Cells")
    assert heading_path(None, "Book") == ("Book",)


# -- convert to Word -------------------------------------------------------
def _word_text(path: Path) -> str:
    import docx
    d = docx.Document(str(path))
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        parts += [c.text for row in t.rows for c in row.cells]
    return "\n".join(parts)


def test_convert_epub_keeps_formatting_and_images(tmp_path):
    import docx
    import converter
    from ebooklib import epub
    book = epub.EpubBook()
    book.set_identifier("id1")
    book.set_title("Epub Title")
    book.add_author("An Author")
    book.add_item(epub.EpubItem(uid="img1", file_name="images/pic.png", media_type="image/png",
                                content=png_bytes()))
    c1 = epub.EpubHtml(title="One", file_name="text/ch1.xhtml")
    c1.content = (
        "<html><body><h1>Chapter 1</h1><p>Hello <em>world</em> of <b>cells</b>.</p>"
        "<p><img src='../images/pic.png'/></p><ul><li>item one</li><li>item two</li></ul>"
        "<table><tr><th>key</th><td>value</td></tr></table></body></html>"
    )
    c2 = epub.EpubHtml(title="Two", file_name="text/ch2.xhtml")
    c2.content = "<html><body><h1>Chapter 2</h1><blockquote>A quote.</blockquote></body></html>"
    for c in (c1, c2):
        book.add_item(c)
    book.spine = ["nav", c1, c2]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    src = tmp_path / "b.epub"
    epub.write_epub(str(src), book)

    result = converter.convert_to_docx(src, tmp_path / "out")
    assert result.output == tmp_path / "out" / "b.docx"
    d = docx.Document(str(result.output))
    assert d.core_properties.title == "Epub Title" and d.core_properties.author == "An Author"
    styles = {p.text: p.style.name for p in d.paragraphs if p.text}
    assert styles["Chapter 1"] == "Heading 1" and styles["Chapter 2"] == "Heading 1"
    assert styles["item one"] == "List Bullet" and styles["A quote."] == "Quote"
    hello = next(p for p in d.paragraphs if p.text.startswith("Hello"))
    assert hello.text == "Hello world of cells."
    assert any(r.italic for r in hello.runs if r.text == "world")
    assert any(r.bold for r in hello.runs if r.text == "cells")
    assert len(d.inline_shapes) == 1
    assert d.tables[0].cell(0, 0).text == "key" and d.tables[0].cell(0, 1).text == "value"

    # The Word copy imports with the same chapters.
    again = converter.convert_to_docx(src, tmp_path / "out")
    assert again.output.name == "b (2).docx"
    parsed = mp_parser.parse_file(result.output, tmp_path / "media", "media")
    assert [s.heading for s in parsed.sections] == ["Chapter 1", "Chapter 2"]
    assert len(parsed.sections[0].image_paths) == 1


def test_convert_text_markdown_and_docx(tmp_path):
    import converter
    txt = tmp_path / "notes.txt"
    txt.write_text("CHAPTER ONE\n\nSome notes\nwrapped here.\n\nMore notes.\n", encoding="utf-8")
    out = converter.convert_to_docx(txt, tmp_path / "w").output
    assert "Some notes wrapped here." in _word_text(out)

    (tmp_path / "p.png").write_bytes(png_bytes())
    md = tmp_path / "m.md"
    md.write_text("# Title\n\nText ![x](p.png)\n", encoding="utf-8")
    import docx
    d = docx.Document(str(converter.convert_to_docx(md, tmp_path / "w").output))
    assert d.paragraphs[0].style.name == "Heading 1" and len(d.inline_shapes) == 1

    original = tmp_path / "orig.docx"
    docx.Document().save(original)
    copied = converter.convert_to_docx(original, tmp_path / "w").output
    assert copied.read_bytes() == original.read_bytes() and copied != original

    with pytest.raises(mp_parser.ParseError):
        converter.convert_to_docx(tmp_path / "missing.pdf", tmp_path / "w")
    (tmp_path / "empty.txt").write_text("", encoding="utf-8")
    with pytest.raises(mp_parser.ParseError):
        converter.convert_to_docx(tmp_path / "empty.txt", tmp_path / "w")
    assert sorted(p.name for p in (tmp_path / "w").iterdir()) == ["m.docx", "notes.docx", "orig.docx"]


def test_mark_headings_from_fonts_and_toc(tmp_path):
    import docx
    from docx.shared import Pt
    import converter

    def build() -> Path:
        d = docx.Document()
        for text, size, bold in [("Big Title", 24, False), ("Contents\t1", 11, True),
                                 ("Body text that runs on for a while.", 11, False),
                                 ("Bold Section", 11, True), ("More body text, quite ordinary.", 11, False),
                                 ("Ends with a full stop.", 11, True)]:
            run = d.add_paragraph().add_run(text)
            run.font.size, run.bold = Pt(size), bold
        path = tmp_path / "h.docx"
        d.save(path)
        return path

    def levels(path: Path) -> dict[str, str]:
        out = {}
        for p in docx.Document(str(path)).paragraphs:
            v = p._p.xpath("./w:pPr/w:outlineLvl/@w:val")
            if v:
                out[p.text] = v[0]
        return out

    path = build()
    assert converter.mark_headings(path) == 2
    assert levels(path) == {"Big Title": "0", "Bold Section": "1"}
    parsed = mp_parser.parse_file(path, tmp_path / "media", "media")
    assert "Bold Section" in [s.heading.split(" › ")[-1] for s in parsed.sections]

    path = build()  # bookmarks win over guessing
    assert converter.mark_headings(path, [(1, "Bold Section")]) == 1
    assert levels(path) == {"Bold Section": "0"}


def test_convert_pdf(tmp_path):
    pytest.importorskip("pdf2docx")
    import converter
    (tmp_path / "i.png").write_bytes(png_bytes())
    pdf = tmp_path / "doc.pdf"
    _make_pdf(pdf, tmp_path / "i.png", with_toc=True)
    result = converter.convert_to_docx(pdf, tmp_path / "w")
    text = _word_text(result.output)
    assert "Introduction" in text and "We measured things carefully." in text
    parsed = mp_parser.parse_file(result.output, tmp_path / "media", "media")
    headings = [s.heading for s in parsed.sections]
    assert "Introduction" in headings and "Methods" in headings
    assert parsed.title == "PDF Title"
    assert any(s.image_paths for s in parsed.sections)


def test_replace_document_keeps_progress_of_unchanged_cards(kb, tmp_path):
    src = tmp_path / "book.md"
    src.write_text("# One\n\nKeep me.\n\n# Two\n\nCut me.\n", encoding="utf-8")
    old_id, _ = parse_into(kb, src)
    kb.update_document(old_id, title="Renamed")
    keep, cut = kb.get_sections(old_id)
    kb.record_answer(keep["id"], known=True)
    kb.add_tag_to_section(keep["id"], "important")
    assert [d["id"] for d in kb.find_documents_by_path(tmp_path / "." / "book.md")] == [old_id]

    src.write_text("# One\n\nKeep me.\n\n# Three\n\nNew text.\n", encoding="utf-8")
    new_id, _ = parse_into(kb, src)
    assert kb.replace_document(old_id, new_id) == 1
    assert kb.get_document(old_id) is None
    assert kb.get_document(new_id)["title"] == "Renamed"
    first, second = kb.get_sections(new_id)
    assert kb.get_review(first["id"])["repetitions"] == 1
    assert kb.section_tags(first["id"]) == ["important"]
    assert kb.get_review(second["id"])["repetitions"] == 0
