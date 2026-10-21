# memorypalace
Second Brain to store information on a local data structure and use as a learning aid.

Memory Palace is a local desktop app. It imports your books and notes, splits them into bite-sized
cards, and quizzes you on them with spaced repetition (SM-2). Everything stays on your machine
in a single SQLite file plus a folder of extracted images.

## Quick start

```bash
python -m venv .venv
.venv\Scripts\activate          # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

Python 3.10+ is required. Tested on 3.14.

## Features

| Area | What it does |
|---|---|
| **Import** | `.pdf` `.epub` `.docx` `.doc` `.txt` `.md`. Text is split at headings (Markdown `#`, EPUB `h1–h3`, Word heading styles, PDF bookmarks or font-size detection, "Chapter N" lines in text files). Long sections are split into parts of about 1,800 characters. Embedded images are extracted and attached to the card they appeared in. |
| **Your own cards** | **New book…** creates an empty book. **＋ Add card** (Library, on a book) opens a card editor: a heading (the question), text (the answer) and any number of images. *Save & add another* keeps the editor open for fast entry. Any card, imported or your own, can be edited the same way, including adding or removing images. |
| **Cards from photos** | Pick one photo or a whole series (notes, book pages, slides, whiteboards). Text is read from each photo with **offline OCR**, and you can correct it before saving. Make *one card per photo*, *all photos → one card*, or *pages → split at headings*: photos of consecutive book pages are joined, and a new card starts at each heading. The photo itself can be shown on the card, turned upright using its EXIF orientation and downscaled to 2,000 px. The card editor's **📷 Extract text from images** does the same for a single card. |
| **Read aloud** | **🔊 Read** on a card (or **R** in Review) speaks the heading and text with your computer's built-in voices. Tick *Read aloud* in Review to hear every card as it appears. In recall mode only the question is read until you reveal the answer. **🔊 Voice…** in the sidebar picks the voice and speed. |
| **Review** | One card at a time. **Swipe/drag right** or **→** = *I know* (SM-2 interval grows). **Left** or **←** = *I don't know* (reset; card goes to the back of today's queue). **Down** or **↓/Del** = *Delete* (soft delete, restorable). **Ctrl+Z** undoes. *Recall mode* hides the text until you press **Space**. Filter by document or tag; cap new cards per session. |
| **Library** | Browse and search documents and sections; rename books, edit card text, restore or purge deleted cards, delete documents. |
| **Tags** | Tag a section or a whole document. When you create a **new** tag, the app offers to apply it to every other section that mentions it. Tags can match a keyword (whole word, any case) or a **regex**. *Run auto-tagging* re-applies every tag to the whole library, and new imports are auto-tagged with your existing tags. |
| **Export** | PDF (with bookmarks and Unicode fonts), DOCX, Markdown (images copied to a `_media` folder) or TXT. Export everything, or tick books and sections in the tree. Filter by *reviewed/known*, *mastered* or tag. Images are embedded. |

## Build a standalone .exe (Windows)

```powershell
.\build.ps1            # -> dist\MemoryPalace.exe  (one file, ~47 MB, no Python needed)
.\build.ps1 -OneDir    # -> dist\MemoryPalace\    (folder; starts in ~2 s instead of ~10 s)
```

If PowerShell blocks the script, run `powershell -ExecutionPolicy Bypass -File .\build.ps1`.
The exe keeps its library in a `data` folder **next to the exe**, so put it somewhere writable
(e.g. `Documents\MemoryPalace\`), not in Program Files. The CLI commands below also work with
the exe (`MemoryPalace.exe import book.pdf`), but print nothing in a normal console window
because the exe is built without a console.

## Command line

```bash
python main.py import book.epub notes.md     # import without the GUI
python main.py export all.pdf                 # everything
python main.py export known.docx --reviewed known --tag biology
python main.py autotag                        # re-run all tags
python main.py stats
python main.py --data D:\palace ...           # use another data folder
```

## How it's built

```
main.py      entry point, dependency check, CLI
ui.py        CustomTkinter desktop UI (Review / Library / Tags / Export + card, photo and voice dialogs)
db.py        SQLite schema, queries, tagging (regex runs inside SQLite), SM-2, settings
parser.py    format extractors → common block stream → section chunker
ocr.py       photo clean-up (EXIF rotation, resize) and OCR: Windows OCR, else Tesseract
speech.py    offline text-to-speech: Windows SAPI, macOS `say`, Linux espeak-ng
exporter.py  PDF (reportlab), DOCX (python-docx), Markdown, TXT
tests/       pytest suite; generates its own sample PDF/EPUB/DOCX files
data/        created on first run: memorypalace.db + media/<document>/img_####.*
```

**Schema:** `documents`, `sections` (with `heading`, `content_text`, JSON `image_paths`, `status`),
`tags` (with optional `pattern` / `is_regex`), `section_tags` (records whether a tag was applied
manually or automatically), and `reviews` (`ease_factor`, `interval`, `repetitions`,
`next_review_date`, `last_reviewed_at`, `status` = new / learning / mastered).

**Scheduling (SM-2):** "I know" is grade 4 and "I don't know" is grade 1. The first intervals are
1 day, then 6 days, then the previous interval × ease. A miss resets repetitions and makes the
card due again today. Cards with an interval of 21 days or more count as *mastered*.

## Notes and limits

- **Legacy `.doc`** files are converted through LibreOffice (`soffice` on PATH) or, on Windows,
  Microsoft Word through `pywin32`. If neither is available, re-save the file as `.docx`.
- **Scanned PDFs** have no text layer. Their page images are imported, but run OCR first if you
  want text cards.
- DRM-protected EPUBs can't be read.
- **OCR** uses the engine built into Windows 10/11 (the `winrt-*` packages in `requirements.txt`) and
  reads the languages installed in *Settings › Time & language › Language & region*. On macOS/Linux,
  or as a fallback, install [Tesseract](https://github.com/tesseract-ocr/tesseract) and put it on PATH.
  Handwriting and photos taken at a steep angle are read less reliably, so check the text before saving.
- **Read aloud** uses Windows' SAPI voices (through `pywin32`), `say` on macOS, or `espeak-ng` on Linux.
  Nothing is sent online.
- Back up your library by copying the `data/` folder (the app should be closed).

## Tests

```bash
python -m pytest -q
```
