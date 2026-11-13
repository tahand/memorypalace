"""
main.py - Entry point for Memory Palace.

    python main.py                      # launch the desktop app
    python main.py convert FILE [FILE…] [--out DIR]  # make editable Word copies
    python main.py import FILE [FILE…]  # import from the command line
    python main.py export OUT.pdf [--format pdf|docx|md|txt] [--reviewed known|mastered]
                                  [--document ID …] [--tag NAME] [--title TEXT] [--no-images]
    python main.py autotag              # run every tag over the whole library
    python main.py stats

Global option:  --data DIR  (default: ./data next to this file)
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

# When bundled by PyInstaller (see build.ps1), __file__ points into a temporary
# extraction folder that is wiped on exit, so user data must live next to the .exe.
FROZEN = getattr(sys, "frozen", False)
APP_DIR = Path(sys.executable).resolve().parent if FROZEN else HERE

# (import name, pip name, needed for)
DEPENDENCIES = [
    ("customtkinter", "customtkinter", "the desktop interface"),
    ("PIL", "Pillow", "showing and converting images"),
    ("pymupdf", "pymupdf", "PDF import (pypdf is used as a fallback)"),
    ("ebooklib", "EbookLib", "EPUB import"),
    ("bs4", "beautifulsoup4", "EPUB import"),
    ("docx", "python-docx", "DOCX import and export"),
    ("pdf2docx", "pdf2docx", "converting PDFs to Word with their layout"),
    ("reportlab", "reportlab", "PDF export"),
]
if sys.platform == "win32":
    DEPENDENCIES += [
        ("win32com", "pywin32", "reading cards aloud"),
        ("winrt.windows.media.ocr", "winrt-Windows.Media.Ocr", "reading text from photos (OCR)"),
    ]


def missing_dependencies() -> list[tuple[str, str]]:
    return [(pip, why) for mod, pip, why in DEPENDENCIES if importlib.util.find_spec(mod) is None]


def warn_missing(require_gui: bool) -> bool:
    """Print missing packages. Returns False if the app cannot start at all."""
    if FROZEN:  # everything needed was bundled at build time
        return True
    missing = missing_dependencies()
    if missing:
        print("Some optional packages are missing; these features will be unavailable:", file=sys.stderr)
        for pip, why in missing:
            print(f"  - {pip:<16} ({why})", file=sys.stderr)
        print(f"Install everything with:  pip install -r {HERE / 'requirements.txt'}\n", file=sys.stderr)
    if require_gui and any(pip == "customtkinter" for pip, _ in missing):
        print("customtkinter is required for the desktop interface.", file=sys.stderr)
        return False
    return True


def open_kb(data_dir: Path):
    from db import DatabaseError, KnowledgeBase
    try:
        return KnowledgeBase(data_dir)
    except DatabaseError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        if FROZEN:  # no console in the windowed exe: show the error in a dialog
            from tkinter import Tk, messagebox
            root = Tk()
            root.withdraw()
            messagebox.showerror("Memory Palace", str(exc))
            root.destroy()
        sys.exit(2)


# ---------------------------------------------------------------------------
# CLI commands
# ---------------------------------------------------------------------------
def cmd_gui(args) -> int:
    if not warn_missing(require_gui=True):
        return 1
    kb = open_kb(args.data)
    try:
        import customtkinter as ctk
        from ui import App
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        app = App(kb)
        app.mainloop()
    finally:
        kb.close()
    return 0


def cmd_convert(args) -> int:
    from converter import convert_to_docx
    from parser import ParseError

    out_dir = args.out
    if out_dir is None:
        kb = open_kb(args.data)
        saved = kb.get_setting("word_folder")
        kb.close()
        out_dir = Path(saved) if saved else args.data.resolve().parent / "Word copies"
    failures = 0
    for name in args.files:
        path = Path(name)
        try:
            result = convert_to_docx(path, out_dir)
        except (ParseError, OSError) as exc:
            failures += 1
            print(f"FAIL  {path.name}: {exc}", file=sys.stderr)
            continue
        print(f"ok    {path.name} -> {result.output}  ({result.method})")
        for w in result.warnings:
            print(f"      note: {w}")
    return 1 if failures else 0


def cmd_import(args) -> int:
    import shutil
    from parser import ParseError, parse_file
    from ui import file_sha256  # pure helper, no Tk window is created

    kb = open_kb(args.data)
    failures = 0
    for name in args.files:
        path = Path(name)
        media_abs, media_rel = kb.new_media_dir()
        try:
            digest = file_sha256(path)
            if kb.find_document_by_hash(digest) and not args.force:
                print(f"skip  {path.name}: already imported (use --force)")
                continue
            parsed = parse_file(path, media_abs, media_rel)
            _, auto = kb.add_document(parsed, str(path.resolve()), digest, media_rel)
            print(f"ok    {path.name}: “{parsed.title}”, {len(parsed.sections)} sections, {auto} auto-tags")
            for w in parsed.warnings:
                print(f"      note: {w}")
        except (ParseError, OSError) as exc:
            shutil.rmtree(media_abs, ignore_errors=True)
            failures += 1
            print(f"FAIL  {path.name}: {exc}", file=sys.stderr)
    kb.close()
    return 1 if failures else 0


def cmd_export(args) -> int:
    from exporter import ExportError, ExportOptions, export

    kb = open_kb(args.data)
    tag_id = None
    if args.tag:
        tag = kb.get_tag(args.tag)
        if not tag:
            print(f"Unknown tag: {args.tag}", file=sys.stderr)
            return 1
        tag_id = tag["id"]
    groups = kb.get_export_groups(document_ids=args.document or None, review_filter=args.reviewed, tag_id=tag_id)
    kb.close()
    try:
        out = export(groups, args.out, args.format,
                     ExportOptions(title=args.title, include_images=not args.no_images))
    except ExportError as exc:
        print(f"Export failed: {exc}", file=sys.stderr)
        return 1
    print(f"Wrote {sum(len(g.sections) for g in groups)} sections to {out}")
    return 0


def cmd_autotag(args) -> int:
    kb = open_kb(args.data)
    results = kb.run_auto_tagging()
    kb.close()
    for name, n in results.items():
        print(f"#{name}: +{n}")
    if not results:
        print("No tags defined yet.")
    return 0


def cmd_stats(args) -> int:
    kb = open_kb(args.data)
    s = kb.stats()
    docs = kb.list_documents()
    kb.close()
    print(f"Documents: {len(docs)}")
    print(f"Cards: {s['total']}  (new {s['new']}, learning {s['learning']}, mastered {s['mastered']})")
    print(f"Due now: {s['due']}")
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="memorypalace", description="Local knowledge base with spaced repetition.")
    p.add_argument("--data", type=Path, default=APP_DIR / "data",
                   help="data directory (default: 'data' next to main.py or the .exe)")
    sub = p.add_subparsers(dest="command")

    conv = sub.add_parser("convert", help="make editable Word copies of documents")
    conv.add_argument("files", nargs="+")
    conv.add_argument("--out", type=Path, help="folder for the Word copies (default: the app's 'Word copies' folder)")
    conv.set_defaults(func=cmd_convert)

    imp = sub.add_parser("import", help="import documents")
    imp.add_argument("files", nargs="+")
    imp.add_argument("--force", action="store_true", help="import even if already imported")
    imp.set_defaults(func=cmd_import)

    exp = sub.add_parser("export", help="export sections")
    exp.add_argument("out", type=Path)
    exp.add_argument("--format", choices=["pdf", "docx", "md", "txt"])
    exp.add_argument("--reviewed", choices=["all", "known", "mastered"], default="all")
    exp.add_argument("--document", type=int, action="append", help="document id (repeatable)")
    exp.add_argument("--tag")
    exp.add_argument("--title", default="Memory Palace Export")
    exp.add_argument("--no-images", action="store_true")
    exp.set_defaults(func=cmd_export)

    sub.add_parser("autotag", help="run all tags over the library").set_defaults(func=cmd_autotag)
    sub.add_parser("stats", help="show card counts").set_defaults(func=cmd_stats)
    return p


def _redirect_output_to_log(data_dir: Path) -> None:
    """The windowed exe has no console (stdout/stderr are None): log to a file instead."""
    if sys.stderr is not None and sys.stdout is not None:
        return
    try:
        data_dir.mkdir(parents=True, exist_ok=True)
        log = open(data_dir / "memorypalace.log", "a", encoding="utf-8", buffering=1)
    except OSError:
        return
    sys.stdout = sys.stdout or log
    sys.stderr = sys.stderr or log


def main(argv: list[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    _redirect_output_to_log(args.data)
    if args.command is None:
        return cmd_gui(args)
    warn_missing(require_gui=False)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
