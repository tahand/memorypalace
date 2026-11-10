"""
ui.py - CustomTkinter desktop interface for Memory Palace.

Views (switched from the sidebar):
* Read    - the main screen: books by chapter and header; text shown in full or
            one sentence per line with a ✕ to delete it.
* Review  - one card at a time, swipe or buttons: right = I know,
            left = I don't know, down = delete. Arrow keys work too; Ctrl+Z undoes.
* Library - browse, search, edit, tag, delete and restore documents and sections.
* Tags    - create keyword/regex tags, preview matches, run auto-tagging.
* Export  - tree picker with checkboxes, review filters and PDF/DOCX/MD/TXT output.

Cards are edited in place on the Review and Library screens (``CardEditor``):
change the heading, text or images and press Save (or Ctrl+S).

Dialogs: New book, New card (text and images, OCR an image into the text), Cards from photos (OCR a series of photos into cards) and Voice
(read-aloud settings). Read aloud goes through ``App.speak`` so only one card
is spoken at a time and the button that started it shows "Stop".

Threading: parsing and exporting run on worker threads, which never touch Tk
or SQLite. They hand results back by posting callables to ``App.ui_queue``,
and the main loop runs those every 80 ms.
"""
from __future__ import annotations

import hashlib
import os
import queue
import shutil
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from collections import deque
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

import customtkinter as ctk

try:
    from PIL import Image, ImageOps
except ImportError:  # images are shown as placeholders without Pillow
    Image = ImageOps = None

from db import MANUAL_FORMAT, KnowledgeBase
from exporter import EXPORT_FORMATS, ExportError, ExportOptions, export
from ocr import PHOTO_EXTENSIONS, OcrError, engine_name, extract_text, photo_for_card, unavailable_message
from parser import SUPPORTED_EXTENSIONS, Block, ParseError, build_sections, parse_file, text_to_blocks
from lines import HEADING_SEP, heading_path, remove_line, split_lines
from speech import Speaker, speakable

APP_NAME = "Memory Palace"
GREEN = "#2e9e5b"
RED = "#d9534f"
AMBER = "#d08b1e"
MUTED = ("gray40", "gray65")
FONT_FAMILY = "Segoe UI" if sys.platform == "win32" else "Helvetica"


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def open_with_system(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError as exc:
        messagebox.showerror(APP_NAME, f"Could not open {path}:\n{exc}")


def load_ctk_image(path: Path, max_w: int, max_h: int) -> ctk.CTkImage | None:
    """Load an image as a thumbnail, or return None if it can't be shown."""
    if Image is None:
        return None
    try:
        with Image.open(path) as im:
            im.load()
            im = ImageOps.exif_transpose(im)  # phone photos store their rotation in EXIF
    except Exception:
        return None
    im.thumbnail((max_w, max_h))
    return ctk.CTkImage(light_image=im, dark_image=im, size=im.size)


def bind_tree(widget: tk.Misc, sequence: str, func: Callable) -> None:
    """Bind an event on a widget and all its descendants.

    Tk events don't bubble up to parents, so a swipe that starts on the text
    box or an image must be bound on those widgets too. ``tk.Misc.bind`` is
    called directly to skip CustomTkinter's restricted ``bind`` override.
    """
    tk.Misc.bind(widget, sequence, func, "+")
    for child in widget.winfo_children():
        bind_tree(child, sequence, func)


def is_inside(widget: Any, classes: tuple[type, ...]) -> bool:
    while widget is not None:
        if isinstance(widget, classes):
            return True
        widget = getattr(widget, "master", None)
    return False


def apply_treeview_style(root: tk.Misc) -> None:
    """ttk.Treeview doesn't follow CustomTkinter themes; style it by hand."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    dark = ctk.get_appearance_mode() == "Dark"
    bg = "#242424" if dark else "#ffffff"
    fg = "#e6e6e6" if dark else "#1a1a1a"
    head = "#2f2f2f" if dark else "#ececec"
    # ttk ignores CustomTkinter's DPI scaling, so scale pixel sizes by hand.
    try:
        scale = ctk.ScalingTracker.get_widget_scaling(root)
    except Exception:
        scale = 1.0
    style.configure("Treeview", background=bg, fieldbackground=bg, foreground=fg,
                    rowheight=round(28 * scale), borderwidth=0, bordercolor=bg, lightcolor=bg, darkcolor=bg,
                    font=(FONT_FAMILY, -round(14 * scale)))
    style.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background=head, foreground=fg, relief="flat", bordercolor=head,
                    lightcolor=head, darkcolor=head, font=(FONT_FAMILY, -round(13 * scale), "bold"))
    style.map("Treeview.Heading", background=[("active", head)])


def make_tree(parent, columns: dict[str, tuple[str, int]], tree_heading: str, tree_width: int = 320):
    """A Treeview with a CTk scrollbar inside a frame. ``columns``: id -> (title, width)."""
    frame = ctk.CTkFrame(parent)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    tree = ttk.Treeview(frame, columns=list(columns), selectmode="browse")
    tree.heading("#0", text=tree_heading, anchor="w")
    tree.column("#0", width=tree_width, stretch=True)
    for cid, (title, width) in columns.items():
        tree.heading(cid, text=title, anchor="w")
        tree.column(cid, width=width, stretch=False, anchor="w")
    scroll = ctk.CTkScrollbar(frame, command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
    scroll.grid(row=0, column=1, sticky="ns", pady=4)
    tree.tag_configure("deleted", foreground="#888888")
    try:
        scale = ctk.ScalingTracker.get_widget_scaling(parent)
    except Exception:
        scale = 1.0
    tree.tag_configure("doc", font=(FONT_FAMILY, -round(14 * scale), "bold"))
    return frame, tree


class ImageViewer(ctk.CTkToplevel):
    """Full-size image window (opened by double-clicking a thumbnail)."""

    def __init__(self, master, path: Path):
        super().__init__(master)
        self.title(path.name)
        max_w = int(self.winfo_screenwidth() * 0.85)
        max_h = int(self.winfo_screenheight() * 0.85)
        img = load_ctk_image(path, max_w, max_h)
        if img is None:
            ctk.CTkLabel(self, text=f"Cannot display {path.name}").pack(padx=40, pady=40)
        else:
            self._img = img
            ctk.CTkLabel(self, image=img, text="").pack(padx=8, pady=8)
        ctk.CTkButton(self, text="Open in system viewer", command=lambda: open_with_system(path)).pack(pady=(0, 10))
        self.after(150, self.focus_force)


def grab_later(dialog: ctk.CTkToplevel, focus: tk.Misc | None = None) -> None:
    """Make a CTkToplevel modal once it is mapped (grabbing earlier fails on Windows)."""
    def grab():
        try:
            dialog.grab_set()
            (focus or dialog).focus_set()
        except tk.TclError:
            pass
    dialog.after(150, grab)


def split_heading(text: str) -> tuple[str, str]:
    """Use a short first paragraph of OCR text as the card heading."""
    first, _, rest = text.strip().partition("\n\n")
    if rest and len(first) <= 80 and not first.endswith((".", ",", ";", ":", "!", "?")):
        return first, rest.strip()
    return "", text.strip()


def photo_filetypes() -> list[tuple[str, str]]:
    return [("Images", " ".join(f"*{e}" for e in PHOTO_EXTENSIONS)), ("All files", "*.*")]


class NewBookDialog(ctk.CTkToplevel):
    """Ask for the title and author of a new, empty book."""

    def __init__(self, master, on_create: Callable[[str, str], bool]):
        super().__init__(master)
        self.title("New book")
        self.geometry("460x250")
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.on_create = on_create
        self.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self, text="Title").grid(row=0, column=0, sticky="w", padx=16, pady=(14, 0))
        self.name = ctk.CTkEntry(self, placeholder_text="e.g. Biology notes")
        self.name.grid(row=1, column=0, sticky="ew", padx=16)
        ctk.CTkLabel(self, text="Author (optional)").grid(row=2, column=0, sticky="w", padx=16, pady=(10, 0))
        self.author = ctk.CTkEntry(self)
        self.author.grid(row=3, column=0, sticky="ew", padx=16)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=4, column=0, sticky="e", padx=16, pady=16)
        ctk.CTkButton(bar, text="Cancel", fg_color="gray50", width=90, command=self.destroy).pack(side="left", padx=6)
        ctk.CTkButton(bar, text="Create book", width=110, command=self._create).pack(side="left")
        tk.Misc.bind(self, "<Return>", lambda e: self._create(), "+")
        tk.Misc.bind(self, "<Escape>", lambda e: self.destroy(), "+")
        grab_later(self, self.name)

    def _create(self):
        if not self.name.get().strip():
            self.name.focus_set()
            return
        title, author = self.name.get(), self.author.get()
        self.destroy()
        self.on_create(title, author)


class CardDialog(ctk.CTkToplevel):
    """Create new cards: heading, text and images. (Existing cards are edited
    in place with ``CardEditor``.)

    Images can be added from files, and "Extract text" runs OCR over the
    card's images and appends what it reads to the text.
    """

    THUMB = (150, 110)

    def __init__(self, master, app: "App", document_id: int,
                 on_saved: Callable[[int], None] | None = None, after: int | None = None):
        super().__init__(master)
        self.app, self.kb = app, app.kb
        self.document_id = document_id
        self.on_saved = on_saved
        self.after = after  # new cards go right after this section; None = end of the book
        # A str is an image already stored in the library; a Path is a new file to add on save.
        self.images: list[str | Path] = []
        self._thumbs: list = []
        self._ocr_running = False
        self.saved_count = 0

        doc = self.kb.get_document(document_id) or {"title": ""}
        self.title(f"New card — {doc['title']}")
        self.geometry("860x680")
        self.minsize(620, 480)
        self.transient(master.winfo_toplevel())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        small = app.fonts["small"]

        ctk.CTkLabel(self, text="Heading (the question or title shown on the card)").grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 0))
        self.heading = ctk.CTkEntry(self)
        self.heading.grid(row=1, column=0, sticky="ew", padx=14)

        text_bar = ctk.CTkFrame(self, fg_color="transparent")
        text_bar.grid(row=2, column=0, sticky="ew", padx=14, pady=(10, 2))
        ctk.CTkLabel(text_bar, text="Text (the answer or note)").pack(side="left")
        self.listen_btn = ctk.CTkButton(text_bar, text="🔊 Listen", width=90, height=26)
        self.listen_btn.configure(command=lambda: self.app.toggle_speech(
            self.heading.get(), self.text.get("1.0", "end"), self.listen_btn))
        self.listen_btn.pack(side="right")
        self.text = ctk.CTkTextbox(self, wrap="word", font=ctk.CTkFont(FONT_FAMILY, 14))
        self.text.grid(row=3, column=0, sticky="nsew", padx=14)

        img_bar = ctk.CTkFrame(self, fg_color="transparent")
        img_bar.grid(row=4, column=0, sticky="ew", padx=14, pady=(10, 2))
        ctk.CTkLabel(img_bar, text="Images").pack(side="left")
        ctk.CTkButton(img_bar, text="🖼 Add images…", width=120, height=26, command=self._add_images).pack(
            side="left", padx=10)
        self.ocr_btn = ctk.CTkButton(img_bar, text="📷 Extract text from images", width=190, height=26,
                                     command=self._extract_text)
        self.ocr_btn.pack(side="left")
        self.strip_holder = ctk.CTkFrame(self, fg_color="transparent", height=self.THUMB[1] + 50)
        self.strip_holder.grid(row=5, column=0, sticky="ew", padx=14)
        self.strip_holder.grid_columnconfigure(0, weight=1)

        self.status = ctk.CTkLabel(self, text="", text_color=MUTED, font=small, anchor="w")
        self.status.grid(row=6, column=0, sticky="ew", padx=14)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=7, column=0, sticky="e", padx=14, pady=12)
        ctk.CTkButton(bar, text="Close", fg_color="gray50", width=90,
                      command=self.destroy).pack(side="left", padx=6)
        ctk.CTkButton(bar, text="Save & add another", width=150,
                      command=lambda: self._save(another=True)).pack(side="left", padx=(0, 6))
        ctk.CTkButton(bar, text="Save", width=90, command=self._save).pack(side="left")
        tk.Misc.bind(self, "<Control-Return>", lambda e: self._save(another=True), "+")
        self._render_images()
        grab_later(self, self.heading)

    def destroy(self):
        self.app.stop_speech(self.listen_btn)
        super().destroy()

    # -- images ----------------------------------------------------------
    def _image_path(self, img: str | Path) -> Path:
        return self.kb.abs_path(img) if isinstance(img, str) else img

    def _render_images(self):
        for w in self.strip_holder.winfo_children():
            w.destroy()
        self._thumbs.clear()
        if not self.images:
            ctk.CTkLabel(self.strip_holder, text="No images. Add a photo or diagram to show it on the card.",
                         text_color=MUTED, font=self.app.fonts["small"]).grid(row=0, column=0, sticky="w", pady=6)
            return
        strip = ctk.CTkScrollableFrame(self.strip_holder, orientation="horizontal", height=self.THUMB[1] + 36,
                                       fg_color="transparent")
        strip.grid(row=0, column=0, sticky="ew")
        for img in self.images:
            path = self._image_path(img)
            cell = ctk.CTkFrame(strip, fg_color="transparent")
            cell.pack(side="left", padx=4)
            thumb = load_ctk_image(path, *self.THUMB)
            if thumb is None:
                ctk.CTkLabel(cell, text=f"🖼 {path.name}", width=self.THUMB[0], height=self.THUMB[1],
                             fg_color=("gray85", "gray25"), corner_radius=8).pack()
            else:
                self._thumbs.append(thumb)
                ctk.CTkLabel(cell, image=thumb, text="").pack()
            ctk.CTkButton(cell, text="✕ Remove", height=22, width=80, fg_color="gray45", hover_color=RED,
                          command=lambda i=img: self._remove_image(i)).pack(pady=(4, 0))

    def _add_images(self):
        paths = filedialog.askopenfilenames(parent=self, title="Add images", filetypes=photo_filetypes())
        if paths:
            self.images += [Path(p) for p in paths]
            self._render_images()
            if engine_name() and messagebox.askyesno(
                    "Extract text", "Read the text in the new image(s) and add it to the card?", parent=self):
                self._extract_text([Path(p) for p in paths])

    def _remove_image(self, img: str | Path):
        self.images.remove(img)
        self._render_images()

    def _extract_text(self, paths: list[Path] | None = None):
        if self._ocr_running:
            return
        paths = paths or [self._image_path(i) for i in self.images]
        if not paths:
            messagebox.showinfo("Extract text", "Add an image to the card first.", parent=self)
            return
        if not engine_name():
            messagebox.showwarning("Extract text", unavailable_message(), parent=self)
            return
        self._ocr_running = True
        self.ocr_btn.configure(state="disabled")
        self.status.configure(text=f"Reading text from {len(paths)} image(s)…")

        def done(texts: list[str]):
            if not self.winfo_exists():
                return
            self._ocr_running = False
            self.ocr_btn.configure(state="normal")
            found = "\n\n".join(t for t in texts if t)
            if not found:
                self.status.configure(text="No text was found in the image(s).")
                return
            if not self.heading.get().strip():
                heading, found = split_heading(found)
                self.heading.insert(0, heading)
            current = self.text.get("1.0", "end").strip()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", f"{current}\n\n{found}".strip())
            self.status.configure(text="Text extracted. Check it for recognition mistakes before saving.")

        def failed(exc: BaseException, _tb: str):
            if not self.winfo_exists():
                return
            self._ocr_running = False
            self.ocr_btn.configure(state="normal")
            self.status.configure(text="")
            messagebox.showerror("Extract text", str(exc), parent=self)

        self.app.run_background(lambda: [extract_text(p) for p in paths], done, failed)

    # -- save ------------------------------------------------------------
    def _store_images(self) -> list[str]:
        """Copy newly added images into the library; returns all relative paths."""
        for i, img in enumerate(self.images):
            if isinstance(img, Path):  # replaced in place so a retry doesn't store it twice
                self.images[i] = self.kb.save_image(self.document_id, *photo_for_card(img))
        return [str(i) for i in self.images]

    def _save(self, another: bool = False):
        heading = self.heading.get()
        text = self.text.get("1.0", "end").strip()
        if not text and not self.images:
            messagebox.showinfo("Card", "Add some text or an image first.", parent=self)
            return False
        try:
            rels = self._store_images()
            section_id = self.kb.add_section(self.document_id, heading, text, rels, after=self.after)
            if self.after is not None:  # keep "add another" cards in the order they were made
                self.after = section_id
        except (OcrError, OSError, ValueError) as exc:
            messagebox.showerror("Card", f"Could not save the card:\n{exc}", parent=self)
            return False
        self.saved_count += 1
        if self.on_saved:
            self.on_saved(section_id)
        if not another:
            self.destroy()
            return
        self.app.stop_speech(self.listen_btn)
        self.heading.delete(0, "end")
        self.text.delete("1.0", "end")
        self.images = []
        self._render_images()
        self.status.configure(text=f"Saved {self.saved_count} card(s). Add the next one.")
        self.heading.focus_set()


class CardEditor(ctk.CTkFrame):
    """An existing card edited in place: heading, text and images, with a Save
    button. Nothing is written until Save (or Ctrl+S); the button is enabled
    only while there are unsaved changes.

    ``show_body=False`` shows only the heading (recall mode before reveal).
    ``on_saved(section_id)`` runs after a save or a split.
    """

    def __init__(self, master, app: "App", section: dict, on_saved: Callable[[int], None],
                 heading_font=None, img_size: tuple[int, int] = (320, 200), show_body: bool = True,
                 transparent_text: bool = False):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.section = section
        self.on_saved = on_saved
        self.img_size = img_size
        self.show_body = show_body
        # A str is an image already stored in the library; a Path is a new file to add on save.
        self.images: list[str | Path] = list(section["image_paths"])
        self._thumbs: list = []
        self._ocr_running = False
        self.grid_columnconfigure(0, weight=1)
        small = app.fonts["small"]

        self.heading = ctk.CTkEntry(self, font=heading_font or app.fonts["h2"],
                                    placeholder_text="Heading (the question or title shown on the card)")
        if section.get("heading"):
            self.heading.insert(0, section["heading"])
        self.heading.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.text: ctk.CTkTextbox | None = None
        self.strip_holder: ctk.CTkFrame | None = None
        if show_body:
            kwargs = {"fg_color": "transparent", "border_width": 1} if transparent_text else {}
            self.text = ctk.CTkTextbox(self, wrap="word", font=app.fonts["body"], border_spacing=6, **kwargs)
            self.text.insert("1.0", section.get("content_text") or "")
            self.text.grid(row=1, column=0, sticky="nsew")
            self.text.edit_modified(False)
            self.strip_holder = ctk.CTkFrame(self, fg_color="transparent")
            self.strip_holder.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
            self.strip_holder.grid_columnconfigure(0, weight=1)
            self.strip_holder.grid_rowconfigure(0, weight=1)
            # Text gets the space; an image-only card gives it to the images instead.
            self.grid_rowconfigure(1 if section.get("content_text") or not self.images else 2, weight=1)

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        if show_body:
            ctk.CTkButton(bar, text="🖼 Add images…", width=10, height=28,
                          command=self._add_images).pack(side="left", padx=(0, 6))
            self.ocr_btn = ctk.CTkButton(bar, text="📷 Extract text", width=10, height=28,
                                         command=self._extract_text)
            self.ocr_btn.pack(side="left", padx=(0, 6))
            ctk.CTkButton(bar, text="✂ Split at cursor", width=10, height=28,
                          command=self._split).pack(side="left", padx=(0, 6))
        self.save_btn = ctk.CTkButton(bar, text="💾 Save", width=90, height=28, fg_color=GREEN,
                                      hover_color="#237a46", command=self.save)
        self.save_btn.pack(side="right")
        self.status = ctk.CTkLabel(bar, text="", text_color=MUTED, font=small, anchor="e")
        self.status.pack(side="right", padx=8)

        for w in (self.heading, self.text):
            if w is not None:
                tk.Misc.bind(w, "<KeyRelease>", lambda e: self._update_state(), "+")
                tk.Misc.bind(w, "<Control-s>", lambda e: (self.save(), "break")[1], "+")
                tk.Misc.bind(w, "<Escape>", lambda e: self.focus_set(), "+")
        if show_body:
            self._render_images()
        self._update_state()

    # -- state -------------------------------------------------------------
    def get_heading(self) -> str:
        return self.heading.get()

    def get_text(self) -> str:
        return self.text.get("1.0", "end").strip() if self.text else (self.section.get("content_text") or "")

    def dirty(self) -> bool:
        return (self.get_heading().strip() != (self.section.get("heading") or "").strip()
                or self.get_text() != (self.section.get("content_text") or "").strip()
                or [str(i) for i in self.images] != list(self.section["image_paths"]))

    def _update_state(self):
        if self.winfo_exists():
            changed = self.dirty()
            self.save_btn.configure(state="normal" if changed else "disabled")
            if changed:
                self.status.configure(text="Unsaved changes")
            elif self.status.cget("text") == "Unsaved changes":
                self.status.configure(text="")

    def confirm_leave(self, cancel: bool = True) -> bool:
        """Before moving away: offer to save unsaved changes. False = stay here."""
        if not self.winfo_exists() or not self.dirty():
            return True
        ask = messagebox.askyesnocancel if cancel else messagebox.askyesno
        answer = ask("Unsaved changes", "Save your changes to this card?", parent=self)
        if answer is None:
            return False
        return self.save(notify=False) is not False if answer else True

    # -- images ------------------------------------------------------------
    def _image_path(self, img: str | Path) -> Path:
        return self.kb.abs_path(img) if isinstance(img, str) else img

    def _render_images(self):
        if self.strip_holder is None:
            return
        for w in self.strip_holder.winfo_children():
            w.destroy()
        self._thumbs.clear()
        if not self.images:
            return
        max_w, max_h = self.img_size
        strip = ctk.CTkScrollableFrame(self.strip_holder, orientation="horizontal", height=max_h + 40,
                                       fg_color="transparent")
        strip.grid(row=0, column=0, sticky="nsew")
        for img in self.images:
            path = self._image_path(img)
            cell = ctk.CTkFrame(strip, fg_color="transparent")
            cell.pack(side="left", padx=4)
            thumb = load_ctk_image(path, max_w, max_h)
            if thumb is None:
                lbl = ctk.CTkLabel(cell, text=f"🖼 {path.name}\n(preview unavailable)", width=160, height=max_h,
                                   fg_color=("gray85", "gray25"), corner_radius=8)
            else:
                self._thumbs.append(thumb)
                lbl = ctk.CTkLabel(cell, image=thumb, text="", cursor="hand2")
            lbl.pack()
            for w in (lbl, *lbl.winfo_children()):
                tk.Misc.bind(w, "<Double-Button-1>", lambda e, p=path: ImageViewer(self.winfo_toplevel(), p), "+")
            ctk.CTkButton(cell, text="✕ Remove", height=22, width=80, fg_color="gray45", hover_color=RED,
                          command=lambda i=img: self._remove_image(i)).pack(pady=(4, 0))

    def _add_images(self):
        paths = filedialog.askopenfilenames(parent=self, title="Add images", filetypes=photo_filetypes())
        if not paths:
            return
        self.images += [Path(p) for p in paths]
        self._render_images()
        self._update_state()
        if engine_name() and messagebox.askyesno(
                "Extract text", "Read the text in the new image(s) and add it to the card?", parent=self):
            self._extract_text([Path(p) for p in paths])

    def _remove_image(self, img: str | Path):
        self.images.remove(img)
        self._render_images()
        self._update_state()

    def _extract_text(self, paths: list[Path] | None = None):
        if self._ocr_running or self.text is None:
            return
        paths = paths or [self._image_path(i) for i in self.images]
        if not paths:
            messagebox.showinfo("Extract text", "Add an image to the card first.", parent=self)
            return
        if not engine_name():
            messagebox.showwarning("Extract text", unavailable_message(), parent=self)
            return
        self._ocr_running = True
        self.ocr_btn.configure(state="disabled")
        self.status.configure(text=f"Reading text from {len(paths)} image(s)…")

        def done(texts: list[str]):
            if not self.winfo_exists():
                return
            self._ocr_running = False
            self.ocr_btn.configure(state="normal")
            found = "\n\n".join(t for t in texts if t)
            if not found:
                self.status.configure(text="No text was found in the image(s).")
                return
            if not self.heading.get().strip():
                heading, found = split_heading(found)
                self.heading.insert(0, heading)
            current = self.get_text()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", f"{current}\n\n{found}".strip())
            self._update_state()
            self.status.configure(text="Text extracted. Check it, then Save.")

        def failed(exc: BaseException, _tb: str):
            if not self.winfo_exists():
                return
            self._ocr_running = False
            self.ocr_btn.configure(state="normal")
            self.status.configure(text="")
            messagebox.showerror("Extract text", str(exc), parent=self)

        self.app.run_background(lambda: [extract_text(p) for p in paths], done, failed)

    # -- save --------------------------------------------------------------
    def _store_images(self) -> list[str]:
        """Copy newly added images into the library; returns all relative paths."""
        for i, img in enumerate(self.images):
            if isinstance(img, Path):  # replaced in place so a retry doesn't store it twice
                self.images[i] = self.kb.save_image(self.section["document_id"], *photo_for_card(img))
        return [str(i) for i in self.images]

    def save(self, notify: bool = True):
        if not self.dirty():
            return None
        text = self.get_text()
        if not text and not self.images:
            messagebox.showinfo("Card", "Add some text or an image first.", parent=self)
            return False
        try:
            rels = self._store_images()
            self.kb.update_section(self.section["id"], self.get_heading(), text, rels)
        except (OcrError, OSError, ValueError) as exc:
            messagebox.showerror("Card", f"Could not save the card:\n{exc}", parent=self)
            return False
        self.section = {**self.section, "heading": self.get_heading(), "content_text": text, "image_paths": rels}
        self.app.set_status("Card saved.")
        self.status.configure(text="Saved ✓")
        self._update_state()
        if notify:
            self.on_saved(self.section["id"])
        return True

    def _split(self):
        """Save the text before the cursor to this card and the rest to a new
        card right after it. Images stay on this card."""
        before = self.text.get("1.0", "insert").strip()
        after = self.text.get("insert", "end").strip()
        if not before or not after:
            messagebox.showinfo("Split card", "Click in the text where the new card should start.", parent=self)
            return
        heading = self.get_heading()
        try:
            self.kb.update_section(self.section["id"], heading, before, self._store_images())
            self.kb.split_section(self.section["id"], before, after, heading)
        except (OcrError, OSError, ValueError) as exc:
            messagebox.showerror("Split card", f"Could not split the card:\n{exc}", parent=self)
            return
        self.section = {**self.section, "heading": heading, "content_text": before,
                        "image_paths": [str(i) for i in self.images]}
        self.on_saved(self.section["id"])


NEW_BOOK = "＋ New book"


class PhotoCardsDialog(ctk.CTkToplevel):
    """Turn one photo or a series of photos into cards.

    Text is read from every photo with OCR in the background. Each photo's
    heading and text can be corrected before the cards are created.
    """

    MODES = {
        "One card per photo": "each",
        "All photos → one card": "one",
        "Pages → split at headings": "split",
    }
    PREVIEW = (480, 300)

    def __init__(self, master, app: "App", document_id: int | None = None,
                 on_done: Callable[[int], None] | None = None):
        super().__init__(master)
        self.app, self.kb = app, app.kb
        self.on_done = on_done
        self.photos: list[dict] = []
        self.current: dict | None = None
        self._ocr_busy = False
        self._images: list = []
        self.engine = engine_name()

        self.title("Cards from photos")
        self.geometry("1120x760")
        self.minsize(900, 600)
        self.transient(master.winfo_toplevel())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        fonts = app.fonts

        # -- target book
        top = ctk.CTkFrame(self)
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        ctk.CTkLabel(top, text="Add cards to").pack(side="left", padx=(12, 6), pady=8)
        self._books = {NEW_BOOK: None}
        for d in self.kb.list_documents():
            label = d["title"] if d["title"] not in self._books else f"{d['title']} (#{d['id']})"
            self._books[label] = d["id"]
        self.book_menu = ctk.CTkOptionMenu(top, values=list(self._books), width=260, dynamic_resizing=False,
                                           command=lambda _v: self._on_book())
        self.book_menu.pack(side="left")
        self.new_title = ctk.CTkEntry(top, placeholder_text="Title of the new book", width=280)
        start = next((label for label, did in self._books.items() if did == document_id), NEW_BOOK)
        self.book_menu.set(start)
        self._on_book()
        engine = f"Text recognition: {self.engine}" if self.engine else "Text recognition unavailable"
        ctk.CTkLabel(top, text=engine, text_color=MUTED if self.engine else AMBER, font=fonts["small"]).pack(
            side="right", padx=12)

        # -- photo list (left) and editor (right)
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=12)
        body.grid_columnconfigure(0, weight=2, uniform="ph")
        body.grid_columnconfigure(1, weight=3, uniform="ph")
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)
        lbar = ctk.CTkFrame(left, fg_color="transparent")
        lbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(lbar, text="＋ Add photos…", width=120, command=self.add_photos).pack(side="left")
        ctk.CTkButton(lbar, text="▲", width=32, fg_color="gray45", command=lambda: self._move(-1)).pack(
            side="left", padx=(8, 2))
        ctk.CTkButton(lbar, text="▼", width=32, fg_color="gray45", command=lambda: self._move(1)).pack(side="left")
        ctk.CTkButton(lbar, text="Remove", width=70, fg_color="gray45", hover_color=RED,
                      command=self._remove).pack(side="right")
        self.listbox = ctk.CTkScrollableFrame(left, fg_color="transparent")
        self.listbox.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))

        right = ctk.CTkFrame(body)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(4, weight=1)
        self.preview = ctk.CTkLabel(right, text="Add photos of notes, book pages, slides or diagrams.",
                                    text_color=MUTED, height=self.PREVIEW[1])
        self.preview.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        tk.Misc.bind(self.preview, "<Double-Button-1>", lambda e: self._open_full(), "+")
        ctk.CTkLabel(right, text="Heading").grid(row=1, column=0, sticky="w", padx=12)
        self.heading = ctk.CTkEntry(right)
        self.heading.grid(row=2, column=0, sticky="ew", padx=12)
        tbar = ctk.CTkFrame(right, fg_color="transparent")
        tbar.grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 2))
        ctk.CTkLabel(tbar, text="Text read from the photo (edit to fix mistakes)").pack(side="left")
        self.rerun_btn = ctk.CTkButton(tbar, text="↻ Read again", width=100, height=26, command=self._rerun)
        self.rerun_btn.pack(side="right")
        self.listen_btn = ctk.CTkButton(tbar, text="🔊 Listen", width=90, height=26)
        self.listen_btn.configure(command=lambda: self.app.toggle_speech(
            self.heading.get(), self.text.get("1.0", "end"), self.listen_btn))
        self.listen_btn.pack(side="right", padx=6)
        self.text = ctk.CTkTextbox(right, wrap="word", font=ctk.CTkFont(FONT_FAMILY, 14))
        self.text.grid(row=4, column=0, sticky="nsew", padx=12, pady=(0, 12))

        # -- options and actions
        opts = ctk.CTkFrame(self)
        opts.grid(row=2, column=0, sticky="ew", padx=12, pady=6)
        ctk.CTkLabel(opts, text="Make").pack(side="left", padx=(12, 6), pady=8)
        self.mode = ctk.CTkSegmentedButton(opts, values=list(self.MODES))
        self.mode.set("One card per photo")
        self.mode.pack(side="left")
        self.attach = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts, text="Show the photo on the card", variable=self.attach).pack(side="left", padx=16)
        ctk.CTkLabel(self, text="“Pages → split at headings” joins the photos in order and starts a new card at "
                                "each heading (and splits long text), which suits photos of book pages.",
                     text_color=MUTED, font=fonts["small"], anchor="w").grid(row=3, column=0, sticky="ew", padx=24)

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=4, column=0, sticky="ew", padx=12, pady=(6, 12))
        self.status = ctk.CTkLabel(bottom, text="", text_color=MUTED, anchor="w")
        self.status.pack(side="left", padx=12)
        self.create_btn = ctk.CTkButton(bottom, text="Create cards", width=140, fg_color=GREEN,
                                        hover_color="#237a46", command=self._create)
        self.create_btn.pack(side="right")
        ctk.CTkButton(bottom, text="Cancel", width=90, fg_color="gray50", command=self.destroy).pack(
            side="right", padx=6)
        self._show(None)
        grab_later(self)

    def destroy(self):
        self.app.stop_speech(self.listen_btn)
        super().destroy()

    # -- book --------------------------------------------------------------
    def _on_book(self):
        if self.book_menu.get() == NEW_BOOK:
            self.new_title.pack(side="left", padx=8)
        else:
            self.new_title.pack_forget()

    # -- photo list --------------------------------------------------------
    def add_photos(self, paths: list[str] | None = None):
        if paths is None:
            paths = list(filedialog.askopenfilenames(parent=self, title="Choose photos",
                                                     filetypes=photo_filetypes()))
        known = {p["path"] for p in self.photos}
        for p in sorted(Path(x) for x in paths):  # camera names sort in shooting order
            if p in known:
                continue
            self.photos.append({"path": p, "heading": "", "text": "",
                                "state": "pending" if self.engine else "off", "error": "", "thumb": None})
        self._store_current()
        self._render_list()
        if self.current is None and self.photos:
            self._show(self.photos[0])
        self._run_ocr()

    def _render_list(self):
        for w in self.listbox.winfo_children():
            w.destroy()
        marks = {"pending": "… waiting", "running": "⟳ reading", "done": "✓ text", "failed": "✗ no text",
                 "off": ""}
        for i, p in enumerate(self.photos, start=1):
            if p["thumb"] is None:
                p["thumb"] = load_ctk_image(p["path"], 72, 56) or False
            selected = p is self.current
            label = f"{i}. {p['path'].name}\n{p['heading'][:40] or ''}  {marks[p['state']]}".rstrip()
            btn = ctk.CTkButton(
                self.listbox, text=label, image=p["thumb"] or None, compound="left", anchor="w", height=64,
                fg_color=("gray75", "gray28") if selected else "transparent", text_color=("gray10", "gray90"),
                hover_color=("gray80", "gray25"), command=lambda ph=p: self._select(ph))
            btn.pack(fill="x", pady=2)
        if not self.photos:
            ctk.CTkLabel(self.listbox, text="No photos yet.", text_color=MUTED).pack(pady=20)
        self.create_btn.configure(state="normal" if self.photos else "disabled")

    def _select(self, photo: dict):
        self._store_current()
        self._show(photo)
        self._render_list()

    def _move(self, step: int):
        if self.current is None:
            return
        i = self.photos.index(self.current)
        j = max(0, min(len(self.photos) - 1, i + step))
        self.photos[i], self.photos[j] = self.photos[j], self.photos[i]
        self._render_list()

    def _remove(self):
        if self.current is None:
            return
        i = self.photos.index(self.current)
        self.photos.pop(i)
        self.current = None
        self._show(self.photos[min(i, len(self.photos) - 1)] if self.photos else None)
        self._render_list()

    # -- editor ------------------------------------------------------------
    def _store_current(self):
        p = self.current
        if p is not None and p["state"] not in ("pending", "running"):
            p["heading"] = self.heading.get().strip()
            p["text"] = self.text.get("1.0", "end").strip()

    def _show(self, photo: dict | None):
        self.app.stop_speech(self.listen_btn)
        self.current = photo
        self._images.clear()
        self.heading.configure(state="normal")
        self.text.configure(state="normal")
        self.heading.delete(0, "end")
        self.text.delete("1.0", "end")
        if photo is None:
            self.preview.configure(image=None, text="Add photos of notes, book pages, slides or diagrams.")
            self.heading.configure(state="disabled")
            self.text.configure(state="disabled")
            self.rerun_btn.configure(state="disabled")
            return
        img = load_ctk_image(photo["path"], *self.PREVIEW)
        if img is None:
            self.preview.configure(image=None, text=f"Cannot preview {photo['path'].name}")
        else:
            self._images.append(img)
            self.preview.configure(image=img, text="")
        busy = photo["state"] in ("pending", "running")
        self.heading.insert(0, photo["heading"])
        if busy:
            self.text.insert("1.0", "Reading text from this photo…")
        elif photo["state"] == "failed":
            self.text.insert("1.0", photo["text"])
            self.status.configure(text=f"{photo['path'].name}: {photo['error']}")
        else:
            self.text.insert("1.0", photo["text"])
        self.heading.configure(state="disabled" if busy else "normal")
        self.text.configure(state="disabled" if busy else "normal")
        self.rerun_btn.configure(state="normal" if self.engine and not busy else "disabled")

    def _open_full(self):
        if self.current is not None:
            ImageViewer(self, self.current["path"])

    # -- OCR ---------------------------------------------------------------
    def _rerun(self):
        p = self.current
        if p is None or not self.engine or p["state"] in ("pending", "running"):
            return
        p["state"] = "pending"
        self._show(p)
        self._render_list()
        self._run_ocr()

    def _run_ocr(self):
        """Read the photos one at a time on a worker thread."""
        if self._ocr_busy:
            return
        photo = next((p for p in self.photos if p["state"] == "pending"), None)
        left = sum(p["state"] in ("pending", "running") for p in self.photos)
        if photo is None:
            done = sum(p["state"] == "done" for p in self.photos)
            if self.engine and self.photos:
                self.status.configure(text=f"Text read from {done} of {len(self.photos)} photo(s). "
                                           "Check it, then create the cards.")
            return
        self._ocr_busy = True
        photo["state"] = "running"
        self.status.configure(text=f"Reading text… {left} photo(s) left")
        self._render_list()

        def finish():
            self._ocr_busy = False
            if photo is self.current:
                self._show(photo)
            self._render_list()
            self._run_ocr()

        def ok(text: str):
            if not self.winfo_exists():
                return
            photo["heading"], photo["text"] = split_heading(text)
            photo["state"] = "done" if text else "failed"
            photo["error"] = "" if text else "no text found"
            finish()

        def failed(exc: BaseException, _tb: str):
            if not self.winfo_exists():
                return
            photo["state"], photo["error"] = "failed", str(exc)
            finish()

        self.app.run_background(lambda: extract_text(photo["path"]), ok, failed)

    # -- create ------------------------------------------------------------
    def _create(self):
        self._store_current()
        if not self.photos:
            return
        waiting = sum(p["state"] in ("pending", "running") for p in self.photos)
        if waiting and not messagebox.askyesno(
                "Cards from photos", f"Text is still being read from {waiting} photo(s).\n\n"
                                     "Create the cards now without that text?", parent=self):
            return
        attach = self.attach.get()
        mode = self.MODES[self.mode.get()]
        photos = [p for p in self.photos if attach or p["text"] or p["heading"]]
        if not photos:
            messagebox.showinfo("Cards from photos", "No text was found. Tick “Show the photo on the card” to "
                                                     "make picture cards, or type the text in.", parent=self)
            return

        doc_id = self._books.get(self.book_menu.get())
        if doc_id is None:
            title = self.new_title.get().strip() or f"Photo notes {date.today().isoformat()}"
            doc_id = self.kb.create_document(title)
        title = (self.kb.get_document(doc_id) or {}).get("title") or "Photos"

        try:
            rels = {id(p): self.kb.save_image(doc_id, *photo_for_card(p["path"])) for p in photos} if attach else {}
        except (OcrError, OSError) as exc:
            messagebox.showerror("Cards from photos", f"Could not store a photo:\n{exc}", parent=self)
            return
        images = lambda p: [rels[id(p)]] if attach else []  # noqa: E731

        cards: list[tuple[str, str, list[str]]] = []
        if mode == "each":
            for i, p in enumerate(photos, start=1):
                cards.append((p["heading"] or f"Photo {i}", p["text"], images(p)))
        elif mode == "one":
            parts = []
            for i, p in enumerate(photos):
                parts += [p["heading"]] if i and p["heading"] else []
                parts.append(p["text"])
            heading = photos[0]["heading"] or title
            cards.append((heading, "\n\n".join(x for x in parts if x), [r for p in photos for r in images(p)]))
        else:
            blocks: list[Block] = []
            for p in photos:
                if p["heading"]:
                    blocks.append(Block("heading", p["heading"], 1))
                blocks += text_to_blocks(p["text"])
                blocks += [Block("image", path=r) for r in images(p)]
            cards += [(s.heading, s.content_text, s.image_paths) for s in build_sections(blocks, title)]

        made = 0
        for heading, text, imgs in cards:
            if text.strip() or imgs:
                self.kb.add_section(doc_id, heading, text, imgs)
                made += 1
        self.app.notify_data_changed()
        self.app.set_status(f"Created {made} card(s) in “{title}” from {len(photos)} photo(s).")
        self.destroy()
        if self.on_done:
            self.on_done(doc_id)


class VoiceDialog(ctk.CTkToplevel):
    """Pick the read-aloud voice and speed. Changes apply and save immediately."""

    SAMPLE = "This is how your cards will sound when they are read aloud."

    def __init__(self, master, app: "App"):
        super().__init__(master)
        self.app = app
        self.title("Read-aloud voice")
        self.geometry("480x300")
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.grid_columnconfigure(0, weight=1)
        speaker = app.speaker
        if not speaker.available:
            ctk.CTkLabel(self, text=f"Read aloud is unavailable on this computer.\n\n{speaker.error}",
                         wraplength=420, justify="left").grid(padx=20, pady=30)
            return
        voices = speaker.voices() or ["Default"]
        ctk.CTkLabel(self, text="Voice").grid(row=0, column=0, sticky="w", padx=18, pady=(16, 2))
        self.voice = ctk.CTkOptionMenu(self, values=voices, command=lambda _v: self._apply())
        self.voice.set(app.kb.get_setting("tts_voice") if app.kb.get_setting("tts_voice") in voices else voices[0])
        self.voice.grid(row=1, column=0, sticky="ew", padx=18)
        self.rate_label = ctk.CTkLabel(self, text="")
        self.rate_label.grid(row=2, column=0, sticky="w", padx=18, pady=(14, 2))
        self.rate = ctk.CTkSlider(self, from_=-10, to=10, number_of_steps=20, command=lambda _v: self._apply())
        self.rate.set(int(app.kb.get_setting("tts_rate", "0") or 0))
        self.rate.grid(row=3, column=0, sticky="ew", padx=18)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=4, column=0, sticky="e", padx=18, pady=20)
        self.test_btn = ctk.CTkButton(bar, text="🔊 Test", width=90)
        self.test_btn.configure(command=lambda: app.toggle_speech("", self.SAMPLE, self.test_btn))
        self.test_btn.pack(side="left", padx=6)
        ctk.CTkButton(bar, text="Close", width=90, command=self.destroy).pack(side="left")
        self._apply(save=False)
        grab_later(self)

    def _apply(self, save: bool = True):
        rate = int(round(self.rate.get()))
        self.rate_label.configure(text=f"Speed: {'normal' if rate == 0 else f'{rate:+d}'}")
        if save:
            self.app.kb.set_setting("tts_voice", self.voice.get())
            self.app.kb.set_setting("tts_rate", str(rate))
        self.app.apply_voice_settings()

    def destroy(self):
        if hasattr(self, "test_btn"):
            self.app.stop_speech(self.test_btn)
        super().destroy()


def tag_chips(parent, tags: list[str], on_remove: Callable[[str], None], on_add: Callable[[], None]):
    """A row of clickable "#tag ✕" chips plus a "+ Tag" button."""
    row = ctk.CTkFrame(parent, fg_color="transparent")
    for name in tags:
        ctk.CTkButton(
            row, text=f"#{name}  ✕", height=24, width=10, corner_radius=12,
            fg_color=("gray80", "gray30"), text_color=("gray10", "gray90"), hover_color=(RED, RED),
            command=lambda n=name: on_remove(n),
        ).pack(side="left", padx=(0, 4))
    ctk.CTkButton(row, text="+ Tag", height=24, width=60, corner_radius=12, command=on_add).pack(side="left")
    return row


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self, kb: KnowledgeBase):
        super().__init__()
        self.kb = kb
        self.title(APP_NAME)
        try:  # CTk geometry is in scaled units; keep the window on small/high-DPI screens
            scale = ctk.ScalingTracker.get_window_scaling(self)
        except Exception:
            scale = 1.0
        self.geometry(f"1280x{max(640, min(820, int(self.winfo_screenheight() / scale) - 90))}")
        self.minsize(1000, 640)
        self.ui_queue: queue.Queue[Callable[[], None]] = queue.Queue()
        self._importing = False
        self.speaker = Speaker()
        self._speech_button: ctk.CTkButton | None = None
        self._speech_label = ""
        self._speech_poll: str | None = None
        self.apply_voice_settings()
        self.fonts = {
            "body": ctk.CTkFont(FONT_FAMILY, 15),
            "title": ctk.CTkFont(FONT_FAMILY, 20, "bold"),
            "h2": ctk.CTkFont(FONT_FAMILY, 16, "bold"),
            "small": ctk.CTkFont(FONT_FAMILY, 12),
            "small_bold": ctk.CTkFont(FONT_FAMILY, 12, "bold"),
            "big": ctk.CTkFont(FONT_FAMILY, 15, "bold"),
        }
        apply_treeview_style(self)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=0, column=1, sticky="nsew", padx=12, pady=12)
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.views: dict[str, ctk.CTkFrame] = {
            "read": ReadView(self.content, self),
            "review": ReviewView(self.content, self),
            "library": LibraryView(self.content, self),
            "tags": TagsView(self.content, self),
            "export": ExportView(self.content, self),
        }
        self.current_view = ""
        self.show_view("read")

        for key, action in (("<Right>", "know"), ("<Left>", "unknown"), ("<Down>", "delete"),
                            ("<Delete>", "delete")):
            tk.Misc.bind(self, key, lambda e, a=action: self._review_key(a), "+")
        tk.Misc.bind(self, "<Control-z>", lambda e: self._review_key("undo"), "+")
        tk.Misc.bind(self, "<space>", lambda e: self._review_key("reveal"), "+")
        tk.Misc.bind(self, "<Key-r>", lambda e: self._review_key("speak"), "+")

        self.after(80, self._drain_ui_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # -- layout --------------------------------------------------------------
    def _build_sidebar(self):
        bar = ctk.CTkFrame(self, width=210, corner_radius=0)
        bar.grid(row=0, column=0, sticky="nsw")
        bar.grid_rowconfigure(10, weight=1)
        ctk.CTkLabel(bar, text="🏛  Memory Palace", font=ctk.CTkFont(FONT_FAMILY, 18, "bold")).grid(
            row=0, column=0, padx=18, pady=(20, 18), sticky="w")
        self.nav_buttons = {}
        for i, (key, label) in enumerate(
            (("read", "📖  Read"), ("review", "🃏  Review"), ("library", "📚  Library"), ("tags", "🏷  Tags"),
             ("export", "📤  Export")),
            start=1,
        ):
            btn = ctk.CTkButton(bar, text=label, anchor="w", height=38, corner_radius=8,
                                fg_color="transparent", text_color=("gray10", "gray90"),
                                hover_color=("gray75", "gray28"), command=lambda k=key: self.show_view(k))
            btn.grid(row=i, column=0, padx=12, pady=3, sticky="ew")
            self.nav_buttons[key] = btn
        ctk.CTkButton(bar, text="＋  Import documents…", height=38, anchor="w", command=self.import_documents).grid(
            row=6, column=0, padx=12, pady=(18, 4), sticky="ew")
        ctk.CTkButton(bar, text="✎  New book…", height=38, anchor="w", command=self.new_book).grid(
            row=7, column=0, padx=12, pady=4, sticky="ew")
        ctk.CTkButton(bar, text="📷  Cards from photos…", height=38, anchor="w",
                      command=self.cards_from_photos).grid(row=8, column=0, padx=12, pady=4, sticky="ew")
        self.stats_label = ctk.CTkLabel(bar, text="", justify="left", anchor="w", font=self.fonts["small"])
        self.stats_label.grid(row=9, column=0, padx=18, pady=8, sticky="w")
        self.status_label = ctk.CTkLabel(bar, text="", justify="left", anchor="sw", wraplength=180,
                                         text_color=MUTED, font=self.fonts["small"])
        self.status_label.grid(row=10, column=0, padx=18, pady=8, sticky="sw")
        ctk.CTkButton(bar, text="🔊  Voice…", width=170, fg_color="gray45", hover_color="gray35",
                      command=lambda: VoiceDialog(self, self)).grid(row=11, column=0, padx=12, pady=(0, 8))
        mode = ctk.CTkOptionMenu(bar, values=["System", "Light", "Dark"], command=self._set_appearance, width=170)
        mode.set("System")
        mode.grid(row=12, column=0, padx=12, pady=(0, 16))

    def _set_appearance(self, mode: str):
        ctk.set_appearance_mode(mode)
        apply_treeview_style(self)

    def show_view(self, name: str):
        if name != self.current_view:
            self.stop_speech()
        if self.current_view:
            self.views[self.current_view].grid_remove()
        self.current_view = name
        view = self.views[name]
        view.grid(row=0, column=0, sticky="nsew")
        for key, btn in self.nav_buttons.items():
            btn.configure(fg_color=("gray75", "gray28") if key == name else "transparent")
        view.on_show()
        self.update_stats()

    def update_stats(self):
        s = self.kb.stats()
        self.stats_label.configure(
            text=f"Cards: {s['total']}\nDue now: {s['due']}  ·  New: {s['new']}\n"
                 f"Learning: {s['learning']}\nMastered: {s['mastered']}"
        )

    def set_status(self, text: str):
        self.status_label.configure(text=text)

    def notify_data_changed(self, source: str | None = None):
        """Mark every view stale. The visible one (unless it caused the change) refreshes now."""
        for name, view in self.views.items():
            view.dirty = True
        if self.current_view and self.current_view != source:
            self.views[self.current_view].on_show()
        self.update_stats()

    # -- threading ---------------------------------------------------------
    def post(self, fn: Callable[[], None]) -> None:
        """Thread-safe: run ``fn`` on the Tk main thread."""
        self.ui_queue.put(fn)

    def run_background(self, work: Callable[[], Any], on_success: Callable[[Any], None],
                       on_error: Callable[[BaseException, str], None] | None = None) -> None:
        def runner():
            try:
                result = work()
            except BaseException as exc:  # noqa: BLE001 - reported to the user
                tb = traceback.format_exc()
                self.post(lambda e=exc, t=tb: (on_error or self._show_error)(e, t))
            else:
                self.post(lambda r=result: on_success(r))
        threading.Thread(target=runner, daemon=True).start()

    def _show_error(self, exc: BaseException, tb: str):
        messagebox.showerror(APP_NAME, f"{type(exc).__name__}: {exc}")

    def _drain_ui_queue(self):
        try:
            while True:
                fn = self.ui_queue.get_nowait()
                try:
                    fn()
                except Exception as exc:  # keep the loop alive no matter what
                    traceback.print_exc()
                    messagebox.showerror(APP_NAME, f"Unexpected error: {exc}")
        except queue.Empty:
            pass
        self.after(80, self._drain_ui_queue)

    # -- keyboard ----------------------------------------------------------
    def _typing(self) -> bool:
        focus = self.focus_get()
        if isinstance(focus, (tk.Entry, tk.Text, ttk.Entry)):
            try:
                return str(focus.cget("state")) == "normal"
            except tk.TclError:
                pass
        return False

    def _review_key(self, action: str):
        if self.current_view == "read" and action == "undo" and not self._typing():
            self.views["read"].undo()  # type: ignore[attr-defined]
            return
        if self.current_view != "review" or self._typing():
            return
        review: ReviewView = self.views["review"]  # type: ignore[assignment]
        if action == "undo":
            review.undo()
        elif action == "reveal":
            review.reveal()
        elif action == "speak":
            review.toggle_speech()
        else:
            review.answer(action)

    # -- import --------------------------------------------------------------
    def import_documents(self, paths: list[str] | None = None):
        if self._importing:
            messagebox.showinfo(APP_NAME, "An import is already running.")
            return
        if paths is None:
            patterns = " ".join(f"*{e}" for e in SUPPORTED_EXTENSIONS)
            paths = list(filedialog.askopenfilenames(
                title="Import documents",
                filetypes=[("Supported documents", patterns), ("All files", "*.*")],
            ))
        files = [Path(p) for p in paths]
        if not files:
            return
        self._importing = True
        self._import_report = {"ok": [], "failed": [], "warnings": []}
        self._import_step(files, 0)

    def _import_step(self, files: list[Path], idx: int):
        if idx >= len(files):
            self._import_finished()
            return
        path = files[idx]
        label = f"({idx + 1}/{len(files)}) {path.name}"
        self.set_status(f"Reading {label}…")
        report = self._import_report

        def failed(exc: BaseException, tb: str):
            msg = str(exc) if isinstance(exc, (ParseError, OSError)) else f"{type(exc).__name__}: {exc}"
            if not isinstance(exc, (ParseError, OSError)):
                traceback.print_exception(exc)
            report["failed"].append((path.name, msg))
            self._import_step(files, idx + 1)

        def hashed(file_hash: str):
            existing = self.kb.find_document_by_hash(file_hash)
            if existing and not messagebox.askyesno(
                "Already imported",
                f"“{path.name}” is identical to “{existing['title']}”, which is already in your library.\n\n"
                "Import it again anyway?",
            ):
                report["failed"].append((path.name, "skipped (duplicate)"))
                self._import_step(files, idx + 1)
                return
            media_abs, media_rel = self.kb.new_media_dir()
            self.set_status(f"Parsing {label}…")

            def parsed_ok(parsed):
                try:
                    _, auto = self.kb.add_document(parsed, str(path), file_hash, media_rel)
                except Exception as exc:  # noqa: BLE001
                    shutil.rmtree(media_abs, ignore_errors=True)
                    failed(exc, "")
                    return
                report["ok"].append((parsed.title, len(parsed.sections), auto))
                report["warnings"] += [f"{path.name}: {w}" for w in parsed.warnings]
                self._import_step(files, idx + 1)

            def parse_failed(exc, tb):
                shutil.rmtree(media_abs, ignore_errors=True)
                failed(exc, tb)

            self.run_background(lambda: parse_file(path, media_abs, media_rel), parsed_ok, parse_failed)

        self.run_background(lambda: file_sha256(path), hashed, failed)

    def _import_finished(self):
        self._importing = False
        r = self._import_report
        lines = []
        if r["ok"]:
            lines.append("Imported:")
            for title, n, auto in r["ok"]:
                extra = f", {auto} auto-tags" if auto else ""
                lines.append(f"  ✓ {title}: {n} sections{extra}")
        if r["failed"]:
            lines.append("\nNot imported:")
            lines += [f"  ✗ {name}: {msg}" for name, msg in r["failed"]]
        if r["warnings"]:
            lines.append("\nNotes:")
            lines += [f"  • {w}" for w in r["warnings"][:15]]
            if len(r["warnings"]) > 15:
                lines.append(f"  … and {len(r['warnings']) - 15} more")
        self.set_status(f"Import finished: {len(r['ok'])} ok, {len(r['failed'])} failed.")
        self.notify_data_changed()
        (messagebox.showwarning if r["failed"] else messagebox.showinfo)("Import", "\n".join(lines))

    # -- shared tag flow -----------------------------------------------------
    def prompt_add_tag(self, section_id: int | None = None, document_id: int | None = None) -> bool:
        """Ask for a tag name and apply it. For a new tag, offer to auto-tag matches."""
        dialog = ctk.CTkInputDialog(title="Add tag", text="Tag name (a new tag will search your library):")
        name = dialog.get_input()
        if not name or not name.strip():
            return False
        existed = self.kb.get_tag(name) is not None
        try:
            if section_id is not None:
                tag_id = self.kb.add_tag_to_section(section_id, name)
            elif document_id is not None:
                tag_id = self.kb.tag_document(document_id, name)
            else:
                return False
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return False
        if not existed:
            self.offer_auto_tag(tag_id)
        self.notify_data_changed()
        return True

    def offer_auto_tag(self, tag_id: int):
        tag = self.kb.get_tag_by_id(tag_id)
        if not tag:
            return
        n = self.kb.count_auto_tag_candidates(tag_id)
        if n and messagebox.askyesno(
            "Auto-tag",
            f"“{tag['name']}” is a new tag.\n\n{n} other section(s) in your knowledge base mention it.\n"
            "Apply the tag to all of them now?",
        ):
            applied = self.kb.auto_tag(tag_id)
            self.set_status(f"Auto-tagged {applied} section(s) with #{tag['name']}.")

    # -- your own books and cards -----------------------------------------
    def new_book(self):
        NewBookDialog(self, self._create_book)

    def _create_book(self, title: str, author: str):
        try:
            doc_id = self.kb.create_document(title, author)
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.notify_data_changed()
        self.show_in_library(doc_id)
        self.set_status(f"Created “{title.strip()}”. Add your first card.")
        self.add_card(doc_id)

    def add_card(self, document_id: int):
        def saved(section_id: int):
            self.notify_data_changed()
            if self.current_view == "library":
                self.views["library"].select(document_id, section_id)
        CardDialog(self, self, document_id, on_saved=saved)

    def cards_from_photos(self, document_id: int | None = None):
        PhotoCardsDialog(self, self, document_id, on_done=self.show_in_library)

    def show_in_library(self, document_id: int, section_id: int | None = None):
        self.show_view("library")
        self.views["library"].select(document_id, section_id)

    # -- read aloud ----------------------------------------------------------
    def apply_voice_settings(self):
        try:
            rate = int(self.kb.get_setting("tts_rate", "0") or 0)
        except ValueError:
            rate = 0
        self.speaker.configure(self.kb.get_setting("tts_voice"), rate)

    def speak(self, heading: str | None, text: str | None, button: ctk.CTkButton | None = None) -> bool:
        """Read a card aloud. ``button`` shows "Stop" while this speech plays."""
        if not self.speaker.available:
            messagebox.showinfo(APP_NAME, f"Read aloud is unavailable.\n\n{self.speaker.error}")
            return False
        content = speakable(heading, text)
        if not content:
            self.set_status("Nothing to read on this card.")
            return False
        self._reset_speech_button()
        if not self.speaker.speak(content):
            messagebox.showerror(APP_NAME, f"Could not read aloud:\n{self.speaker.error}")
            return False
        self.attach_speech_button(button)
        if self._speech_poll is None:
            self._speech_poll = self.after(400, self._poll_speech)
        return True

    def toggle_speech(self, heading: str | None, text: str | None, button: ctk.CTkButton) -> None:
        if button is not None and button is self._speech_button:
            self.stop_speech()
        else:
            self.speak(heading, text, button)

    def stop_speech(self, button: ctk.CTkButton | None = None) -> None:
        """Stop reading. With ``button``, only if that button started it."""
        if button is not None and button is not self._speech_button:
            return
        self.speaker.stop()
        self._reset_speech_button()

    def attach_speech_button(self, button: ctk.CTkButton | None) -> None:
        """Let ``button`` show "Stop" for the speech that is playing now."""
        if button is None:
            return
        self._reset_speech_button()
        self._speech_button, self._speech_label = button, button.cget("text")
        button.configure(text="⏹ Stop")

    def _poll_speech(self):
        self._speech_poll = None
        if self.speaker.is_speaking():
            self._speech_poll = self.after(250, self._poll_speech)
        else:
            self._reset_speech_button()

    def _reset_speech_button(self):
        btn, self._speech_button = self._speech_button, None
        if btn is not None:
            try:
                btn.configure(text=self._speech_label)
            except Exception:  # the button was destroyed with its card
                pass

    def _on_close(self):
        self.speaker.stop()
        self.kb.close()
        self.destroy()


# ---------------------------------------------------------------------------
# Review view
# ---------------------------------------------------------------------------
class ReviewView(ctk.CTkFrame):
    SWIPE_X = 120     # px horizontal drag to count as a swipe
    SWIPE_DOWN = 160  # deleting needs a longer, deliberate drag

    def __init__(self, master, app: App):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.dirty = True
        self.queue: deque[int] = deque()
        self.current: dict | None = None
        self.card: ctk.CTkFrame | None = None
        self.editor: CardEditor | None = None
        self.revealed = True
        self._images: list = []
        self._drag: tuple[int, int] | None = None
        self._offset = (0, 0)
        self._busy = False
        self.undo_stack: list[tuple] = []
        self.session = {"known": 0, "unknown": 0, "deleted": 0}
        self.speak_btn: ctk.CTkButton | None = None
        self._spoken: tuple[int, bool] | None = None  # (section id, revealed) last auto-read
        self._doc_choices: dict[str, int | None] = {}
        self._tag_choices: dict[str, int | None] = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_toolbar()
        self.stage = ctk.CTkFrame(self, fg_color="transparent")
        self.stage.grid(row=1, column=0, sticky="nsew", pady=8)
        self._build_buttons()

    # -- layout --------------------------------------------------------------
    def _build_toolbar(self):
        bar = ctk.CTkFrame(self)
        bar.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(bar, text="Document").pack(side="left", padx=(12, 4), pady=8)
        self.doc_menu = ctk.CTkOptionMenu(bar, values=["All documents"], width=200, dynamic_resizing=False)
        self.doc_menu.pack(side="left")
        ctk.CTkLabel(bar, text="Tag").pack(side="left", padx=(12, 4))
        self.tag_menu = ctk.CTkOptionMenu(bar, values=["All tags"], width=140, dynamic_resizing=False)
        self.tag_menu.pack(side="left")
        ctk.CTkLabel(bar, text="New cards").pack(side="left", padx=(12, 4))
        self.new_limit = ctk.CTkEntry(bar, width=50)
        self.new_limit.insert(0, "20")
        self.new_limit.pack(side="left")
        self.recall_mode = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(bar, text="Recall mode", variable=self.recall_mode, width=20,
                        command=self._rerender).pack(side="left", padx=12)
        self.auto_read = ctk.BooleanVar(value=self.kb.get_setting("auto_read") == "1")
        ctk.CTkCheckBox(bar, text="Read aloud", variable=self.auto_read, width=20,
                        command=self._toggle_auto_read).pack(side="left", padx=(0, 12))
        ctk.CTkButton(bar, text="▶ Start session", width=120, command=self.start_session).pack(side="left", padx=4)

    def _build_buttons(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew")
        bar.grid_columnconfigure((0, 1, 2), weight=1)
        big = self.app.fonts["big"]
        self.btn_unknown = ctk.CTkButton(bar, text="✗  I don't know  (←)", height=48, font=big, fg_color=RED,
                                         hover_color="#b52f2b", command=lambda: self.answer("unknown"))
        self.btn_delete = ctk.CTkButton(bar, text="🗑  Delete  (↓)", height=48, font=big, fg_color="gray45",
                                        hover_color="gray35", command=lambda: self.answer("delete"))
        self.btn_know = ctk.CTkButton(bar, text="✓  I know  (→)", height=48, font=big, fg_color=GREEN,
                                      hover_color="#237a46", command=lambda: self.answer("know"))
        self.btn_unknown.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.btn_delete.grid(row=0, column=1, sticky="ew", padx=6)
        self.btn_know.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        hint = ctk.CTkFrame(bar, fg_color="transparent")
        hint.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ctk.CTkLabel(hint, text="Drag the card ← → ↓ or use the arrow keys.  Space reveals in recall mode.  "
                                "R reads aloud.",
                     font=self.app.fonts["small"], text_color=MUTED).pack(side="left")
        self.btn_undo = ctk.CTkButton(hint, text="↶ Undo (Ctrl+Z)", width=120, height=26, fg_color="gray45",
                                      command=self.undo)
        self.btn_undo.pack(side="right")
        self.progress = ctk.CTkLabel(hint, text="", font=self.app.fonts["small_bold"])
        self.progress.pack(side="right", padx=16)

    # -- filters / session ---------------------------------------------------
    def on_show(self):
        if self.dirty:
            self._refresh_filters()
            self.dirty = False
            if self.current and not self.kb.get_section(self.current["id"]):
                self.current = None
            if self.current or self.queue:
                self._show_next()
            else:
                self.start_session()

    def _refresh_filters(self):
        self._doc_choices = {"All documents": None}
        for d in self.kb.list_documents():
            label = d["title"] if d["title"] not in self._doc_choices else f"{d['title']} (#{d['id']})"
            self._doc_choices[label] = d["id"]
        self._tag_choices = {"All tags": None}
        self._tag_choices.update({f"#{t['name']}": t["id"] for t in self.kb.list_tags()})
        for menu, choices in ((self.doc_menu, self._doc_choices), (self.tag_menu, self._tag_choices)):
            current = menu.get()
            menu.configure(values=list(choices))
            if current not in choices:
                menu.set(next(iter(choices)))

    def set_filter(self, document_id: int | None = None, tag_id: int | None = None):
        """Called from other views ("Review this document")."""
        self._refresh_filters()
        for label, did in self._doc_choices.items():
            if did == document_id:
                self.doc_menu.set(label)
        for label, tid in self._tag_choices.items():
            if tid == tag_id:
                self.tag_menu.set(label)
        self.dirty = False
        self.start_session()

    def start_session(self):
        if not self._leave_card():
            return
        try:
            limit = max(0, int(self.new_limit.get()))
        except ValueError:
            limit = 20
            self.new_limit.delete(0, "end")
            self.new_limit.insert(0, "20")
        doc_id = self._doc_choices.get(self.doc_menu.get())
        tag_id = self._tag_choices.get(self.tag_menu.get())
        self.queue = deque(self.kb.get_review_queue(document_id=doc_id, tag_id=tag_id, new_limit=limit))
        self.session = {"known": 0, "unknown": 0, "deleted": 0}
        self.undo_stack.clear()
        self.current = None
        self._show_next()

    # -- rendering -----------------------------------------------------------
    def _show_next(self):
        while self.queue:
            section = self.kb.get_section(self.queue[0])
            if section and section["status"] == "active":
                self.current = section
                self.revealed = not self.recall_mode.get()
                self._render_card(section)
                self._update_progress()
                return
            self.queue.popleft()  # deleted elsewhere since the session started
        self.current = None
        self.app.stop_speech()
        self._spoken = None
        self._render_done()
        self._update_progress()

    def _rerender(self):
        if self.current:
            self.revealed = not self.recall_mode.get()
            self._render_card(self.current)

    def _update_progress(self):
        s = self.session
        self.progress.configure(
            text=f"Remaining {len(self.queue)}   ✓ {s['known']}   ✗ {s['unknown']}   🗑 {s['deleted']}"
        )
        state = "normal" if self.current else "disabled"
        for btn in (self.btn_know, self.btn_unknown, self.btn_delete):
            btn.configure(state=state)
        self.btn_undo.configure(state="normal" if self.undo_stack else "disabled")
        self.app.update_stats()

    def _clear_stage(self):
        if self.card is not None:
            self.card.destroy()
        self.editor = None
        self._images.clear()

    def _leave_card(self) -> bool:
        """Offer to save unsaved edits before the card goes away. False = stay."""
        return self.editor is None or self.editor.confirm_leave()

    def _render_done(self):
        self._clear_stage()
        card = ctk.CTkFrame(self.stage, corner_radius=14)
        card.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.card = card
        s = self.session
        reviewed = s["known"] + s["unknown"] + s["deleted"]
        if self.kb.stats()["total"] == 0:
            msg = ("Your palace is empty.\n\nImport a PDF, EPUB, Word, Markdown or text file,\n"
                   "write your own cards in a new book, or turn photos into cards.")
        elif reviewed:
            msg = f"🎉 Session complete!\n\n✓ {s['known']} known   ✗ {s['unknown']} relearned   🗑 {s['deleted']} deleted"
        else:
            msg = "Nothing is due right now. 🎉"
        nxt = self.kb.next_due_date()
        if nxt:
            msg += f"\n\nNext review due: {nxt}"
        ctk.CTkLabel(card, text=msg, font=self.app.fonts["h2"], justify="center").pack(expand=True, pady=(40, 10))
        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.pack(expand=True, pady=(0, 40))
        if self.kb.stats()["total"] == 0:
            ctk.CTkButton(btns, text="Import documents…", command=self.app.import_documents).pack(side="left", padx=4)
            ctk.CTkButton(btns, text="New book…", command=self.app.new_book).pack(side="left", padx=4)
            ctk.CTkButton(btns, text="Cards from photos…", command=self.app.cards_from_photos).pack(
                side="left", padx=4)
        else:
            ctk.CTkButton(btns, text="Check again / start new session", command=self.start_session).pack()

    def _render_card(self, s: dict):
        self._clear_stage()
        fonts = self.app.fonts
        card = ctk.CTkFrame(self.stage, corner_radius=14, border_width=3, border_color=("gray75", "gray30"))
        card.grid_columnconfigure(0, weight=1)
        self.card = card

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 0))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text=f"📖 {s['document_title']}", font=fonts["small_bold"], text_color=MUTED,
                     anchor="w").grid(row=0, column=0, sticky="w")
        self.swipe_hint = ctk.CTkLabel(header, text="", font=fonts["big"])
        self.swipe_hint.grid(row=0, column=1, padx=10)
        self.speak_btn = ctk.CTkButton(header, text="🔊 Read", width=72, height=26, command=self.toggle_speech)
        self.speak_btn.grid(row=0, column=2)

        # The card is edited in place; Save (or Ctrl+S) writes the changes.
        big = not s["content_text"]
        self.editor = CardEditor(card, self.app, s, on_saved=self._edited, heading_font=fonts["title"],
                                 img_size=(720, 440) if big else (360, 220), show_body=self.revealed,
                                 transparent_text=True)
        self.editor.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 0))
        card.grid_rowconfigure(1, weight=1)
        row = 2
        if not self.revealed:
            reveal = ctk.CTkButton(card, text="Show content  (Space)", height=44, font=fonts["big"],
                                   command=self.reveal)
            reveal.grid(row=row, column=0, pady=40)
            card.grid_rowconfigure(1, weight=0)
            card.grid_rowconfigure(row, weight=1)
            row += 1

        footer = ctk.CTkFrame(card, fg_color="transparent")
        footer.grid(row=row, column=0, sticky="ew", padx=18, pady=(4, 12))
        footer.grid_columnconfigure(0, weight=1)
        tag_chips(footer, s["tags"], self._remove_tag, self._add_tag).grid(row=0, column=0, sticky="w")
        srs = (f"{s['review_status'] or 'new'} · ease {s['ease_factor']:.2f} · "
               f"interval {s['interval']}d · reps {s['repetitions']}")
        ctk.CTkLabel(footer, text=srs, font=fonts["small"], text_color=MUTED).grid(row=0, column=1, sticky="e")

        card.place(relx=0, rely=0, relwidth=1, relheight=1)
        bind_tree(card, "<ButtonPress-1>", self._on_press)
        bind_tree(card, "<B1-Motion>", self._on_motion)
        bind_tree(card, "<ButtonRelease-1>", self._on_release)

        key = (s["id"], self.revealed)
        if self._spoken and self._spoken[0] == s["id"] and self.app.speaker.is_speaking():
            self.app.attach_speech_button(self.speak_btn)  # re-rendered while this card is being read
        if self.auto_read.get() and key != self._spoken:
            self._spoken = key
            self._speak_current()

    def reveal(self):
        if self.current and not self.revealed:
            heading = self.editor.get_heading() if self.editor else None
            self.revealed = True
            self._render_card(self.current)
            if self.editor and heading is not None and heading != self.editor.get_heading():
                self.editor.heading.delete(0, "end")  # keep a heading edited before revealing
                self.editor.heading.insert(0, heading)
                self.editor._update_state()

    # -- read aloud ----------------------------------------------------------
    def _speak_current(self):
        s = self.current
        if s:
            # Before the answer is revealed in recall mode, only the question is read.
            self._spoken = (s["id"], self.revealed)
            self.app.speak(s.get("heading"), s["content_text"] if self.revealed else "", self.speak_btn)

    def toggle_speech(self):
        if not self.current or not self.speak_btn:
            return
        if self.app._speech_button is self.speak_btn:
            self.app.stop_speech()
        else:
            self._speak_current()

    def _toggle_auto_read(self):
        self.kb.set_setting("auto_read", "1" if self.auto_read.get() else "0")
        if self.auto_read.get():
            self._speak_current()
        else:
            self.app.stop_speech()

    # -- swipe gestures ------------------------------------------------------
    def _on_press(self, event):
        # Scrollbars, buttons and the editable fields keep their own drag/click behaviour.
        if self._busy or not self.current or is_inside(
                event.widget, (ctk.CTkScrollbar, ctk.CTkButton, ctk.CTkEntry, ctk.CTkTextbox)):
            self._drag = None
            return
        self.focus_set()  # leave the text fields so the arrow keys grade the card again
        self._drag = (event.x_root, event.y_root)

    def _classify(self, dx: int, dy: int) -> str | None:
        if abs(dx) >= abs(dy):
            if dx >= self.SWIPE_X:
                return "know"
            if dx <= -self.SWIPE_X:
                return "unknown"
        elif dy >= self.SWIPE_DOWN:
            return "delete"
        return None

    def _on_motion(self, event):
        if not self._drag or not self.card:
            return
        dx, dy = event.x_root - self._drag[0], event.y_root - self._drag[1]
        if abs(dx) < 8 and abs(dy) < 8:
            return
        x, y = (dx, 0) if abs(dx) >= abs(dy) else (0, max(0, dy))
        self._offset = (x, y)
        self.card.place_configure(x=x, y=y)
        action = self._classify(x, y)
        color, hint = {
            "know": (GREEN, "✓ I know"), "unknown": (RED, "✗ Again"), "delete": (AMBER, "🗑 Delete"),
        }.get(action, (("gray75", "gray30"), ""))
        self.card.configure(border_color=color)
        self.swipe_hint.configure(text=hint, text_color=color if action else MUTED)

    def _on_release(self, event):
        if not self._drag or not self.card:
            return
        dx, dy = event.x_root - self._drag[0], event.y_root - self._drag[1]
        self._drag = None
        x, y = (dx, 0) if abs(dx) >= abs(dy) else (0, max(0, dy))
        action = self._classify(x, y)
        if action:
            self.answer(action, start=(x, y))
        else:
            self._snap_back()

    def _snap_back(self):
        if self.card:
            self._offset = (0, 0)
            self.card.place_configure(x=0, y=0)
            self.card.configure(border_color=("gray75", "gray30"))
            self.swipe_hint.configure(text="")

    # -- actions -------------------------------------------------------------
    def answer(self, action: str, start: tuple[int, int] = (0, 0)):
        if self._busy or not self.current or not self.card:
            return
        if action != "delete" and not self.revealed:
            self.reveal()  # grading an unseen card makes no sense; show it first
            return
        if action != "delete" and not self._leave_card():  # a deleted card's edits don't matter
            self._snap_back()
            return
        self._busy = True
        self.app.stop_speech()
        direction = {"know": (1, 0), "unknown": (-1, 0), "delete": (0, 1)}[action]
        self.card.configure(border_color={"know": GREEN, "unknown": RED, "delete": AMBER}[action])
        self._animate(direction, start, 0, lambda: self._commit(action))

    def _animate(self, direction, start, step, done, frames: int = 8):
        if step > frames or not self.card:
            done()
            return
        t = step / frames
        w, h = self.stage.winfo_width(), self.stage.winfo_height()
        x = start[0] + (direction[0] * w * 1.1 - start[0]) * t if direction[0] else 0
        y = start[1] + (direction[1] * h * 1.1 - start[1]) * t if direction[1] else 0
        self.card.place_configure(x=int(x), y=int(y))
        self.after(14, lambda: self._animate(direction, start, step + 1, done, frames))

    def _commit(self, action: str):
        try:
            sid = self.current["id"]
            snapshot = (list(self.queue), dict(self.session))
            prev = None
            if action == "delete":
                self.kb.soft_delete_section(sid)
                self.session["deleted"] += 1
                self.queue.popleft()
                self.app.notify_data_changed(source="review")
            else:
                prev = self.kb.record_answer(sid, known=(action == "know"))
                self.queue.popleft()
                if action == "unknown":
                    self.queue.append(sid)  # back of the current queue
                    self.session["unknown"] += 1
                else:
                    self.session["known"] += 1
            self.undo_stack.append((action, sid, prev, snapshot))
            del self.undo_stack[:-50]
        finally:
            self._busy = False
            self._offset = (0, 0)
            self._spoken = None  # a card that comes back round is read again
        self._show_next()

    def undo(self):
        if self._busy or not self.undo_stack or not self._leave_card():
            return
        action, sid, prev, (queue_snapshot, session) = self.undo_stack.pop()
        if action == "delete":
            self.kb.restore_section(sid)
            self.app.notify_data_changed(source="review")
        else:
            self.kb.restore_review(sid, prev)
        self.queue = deque(queue_snapshot)
        self.session = session
        self.app.set_status("Undid last action.")
        self._show_next()

    def _edited(self, _sid: int):
        """The card was saved (or split) in place: reload it from the database."""
        self.app.notify_data_changed(source="review")
        self._show_next()

    def _add_tag(self):
        if self.current and self._leave_card() and self.app.prompt_add_tag(section_id=self.current["id"]):
            self._show_next()

    def _remove_tag(self, name: str):
        if self.current and self._leave_card():
            self.kb.remove_tag_from_section(self.current["id"], name)
            self.app.notify_data_changed(source="review")
            self._show_next()


# ---------------------------------------------------------------------------
# Read view (the main screen)
# ---------------------------------------------------------------------------
class ReadView(ctk.CTkFrame):
    """Browse books by chapter and header, and read or prune their text.

    The tree is built from the section headings' breadcrumbs ("Chapter 2 ›
    Cells"). Selecting a book, chapter or header shows all the text under it,
    as full text or line by line: one sentence per row with a ✕ at the start
    that deletes it (Undo / Ctrl+Z brings it back). A single header in full
    text is editable in place with a Save button.
    """

    MODES = ("Full text", "Line by line")
    MAX_LINES = 3000  # a whole book can be huge; beyond this, pick a chapter

    def __init__(self, master, app: App):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.dirty = True
        self.nodes: dict[str, tuple[int, tuple[str, ...]]] = {}  # tree iid -> (document id, heading path)
        self._sections: dict[int, list[dict]] = {}               # document id -> active sections (cache)
        self.current: str | None = None
        self.editor: CardEditor | None = None
        self.speak_btn: ctk.CTkButton | None = None
        self.lines_text: tk.Text | None = None
        self.line_refs: list[tuple[int, int]] = []               # row -> (section id, line index)
        self.undo_stack: list[tuple[int, str, bool]] = []        # (section id, old text, card was removed)
        self.grid_columnconfigure(0, weight=2, uniform="read")
        self.grid_columnconfigure(1, weight=3, uniform="read")
        self.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(self)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ctk.CTkLabel(bar, text="Show").pack(side="left", padx=(12, 6), pady=8)
        self.mode = ctk.CTkSegmentedButton(bar, values=list(self.MODES), command=self._set_mode)
        saved_mode = self.kb.get_setting("read_mode")
        self._mode = saved_mode if saved_mode in self.MODES else self.MODES[1]
        self.mode.set(self._mode)
        self.mode.pack(side="left")
        self.btn_undo = ctk.CTkButton(bar, text="↶ Undo delete (Ctrl+Z)", width=170, fg_color="gray45",
                                      hover_color="gray35", command=self.undo, state="disabled")
        self.btn_undo.pack(side="left", padx=12)
        ctk.CTkLabel(bar, text="Click ✕ at the start of a line to delete it.", font=app.fonts["small"],
                     text_color=MUTED).pack(side="right", padx=12)

        frame, self.tree = make_tree(self, {"info": ("Cards", 60)}, "Chapters & headers")
        frame.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.tree.bind("<<TreeviewOpen>>", self._on_open)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.detail = ctk.CTkFrame(self)
        self.detail.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        self.detail.grid_columnconfigure(0, weight=1)
        self.detail.grid_rowconfigure(2, weight=1)

    # -- tree ----------------------------------------------------------------
    def on_show(self):
        if self.dirty:
            self.refresh()

    def _sections_of(self, doc_id: int) -> list[dict]:
        if doc_id not in self._sections:
            self._sections[doc_id] = self.kb.get_sections(doc_id)
        return self._sections[doc_id]

    @staticmethod
    def _path(s: dict) -> tuple[str, ...]:
        fallback = (s["content_text"] or "").strip().split("\n")[0][:60] or "(image)"
        return heading_path(s["heading"], fallback)

    def refresh(self):
        """Rebuild the tree from the database, keeping what was open and selected."""
        self.dirty = False
        self._sections.clear()
        opened = {key for iid, key in self.nodes.items() if self.tree.exists(iid) and self.tree.item(iid, "open")}
        selected = self.nodes.get(self.current) if self.current else None
        wanted = opened | ({selected} if selected else set())
        self.tree.delete(*self.tree.get_children())
        self.nodes.clear()
        for d in self.kb.list_documents():
            iid = f"d{d['id']}"
            self.nodes[iid] = (d["id"], ())
            self.tree.insert("", "end", iid=iid, text=d["title"], values=(d["active_count"] or 0,), tags=("doc",))
            if any(key[0] == d["id"] for key in wanted):
                self._load_children(iid)
            else:
                self.tree.insert(iid, "end", iid=f"{iid}.x", text="loading…")  # lazy placeholder
        for iid, key in self.nodes.items():
            if key in opened:
                self.tree.item(iid, open=True)
        target = next((i for i, key in self.nodes.items() if key == selected), None)
        if target is None and selected and f"d{selected[0]}" in self.nodes:
            target = f"d{selected[0]}"  # header gone (e.g. renamed): fall back to its book
        if target:
            self.current = target
            self.tree.selection_set(target)
            self.tree.see(target)
            self.show(target, keep_scroll=True)
        else:
            self.current = None
            self._placeholder()

    def _load_children(self, doc_iid: str):
        doc_id = self.nodes[doc_iid][0]
        self.tree.delete(*self.tree.get_children(doc_iid))
        for iid in [i for i in self.nodes if i.startswith(doc_iid + ".")]:
            del self.nodes[iid]
        by_path: dict[tuple[str, ...], str] = {(): doc_iid}
        counts: dict[tuple[str, ...], int] = {}
        for s in self._sections_of(doc_id):
            path = self._path(s)
            for depth in range(1, len(path) + 1):
                prefix = path[:depth]
                counts[prefix] = counts.get(prefix, 0) + 1
                if prefix not in by_path:
                    iid = f"{doc_iid}.{len(by_path)}"
                    by_path[prefix] = iid
                    self.nodes[iid] = (doc_id, prefix)
                    self.tree.insert(by_path[prefix[:-1]], "end", iid=iid, text=prefix[-1])
        for prefix, iid in by_path.items():
            if prefix:
                self.tree.set(iid, "info", counts[prefix])

    def _on_open(self, _event):
        iid = self.tree.focus()
        if self.tree.exists(f"{iid}.x"):
            self._load_children(iid)

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel or sel[0] not in self.nodes or sel[0] == self.current:
            return
        if self.editor is not None and not self.editor.confirm_leave():
            if self.current and self.tree.exists(self.current):
                self.tree.selection_set(self.current)  # stay on the card being edited
            return
        self.editor = None
        self.show(sel[0])

    # -- detail pane ---------------------------------------------------------
    def _clear_detail(self):
        if self.editor is not None:
            self.editor.confirm_leave(cancel=False)  # offer to save edits before they go
            self.editor = None
        if self.speak_btn is not None:
            self.app.stop_speech(self.speak_btn)  # only if this pane was reading
            self.speak_btn = None
        for w in self.detail.winfo_children():
            w.destroy()
        self.lines_text = None
        self.line_refs = []

    def _placeholder(self):
        self._clear_detail()
        if self.nodes:
            msg = "Select a book, chapter or header on the left."
        else:
            msg = ("Your palace is empty.\n\nImport a PDF, EPUB, Word, Markdown or text file,\n"
                   "write your own cards in a new book, or turn photos into cards.")
        ctk.CTkLabel(self.detail, text=msg, text_color=MUTED, justify="center").grid(row=0, column=0, pady=40)

    def _set_mode(self, mode: str):
        if mode == self._mode:
            return
        if self.editor is not None and not self.editor.confirm_leave():
            self.mode.set(self._mode)
            return
        self.editor = None
        self._mode = mode
        self.kb.set_setting("read_mode", mode)
        if self.current:
            self.show(self.current)

    def show(self, iid: str, keep_scroll: bool = False):
        doc_id, prefix = self.nodes[iid]
        scroll = self.lines_text.yview()[0] if keep_scroll and self.lines_text and iid == self.current else None
        self._clear_detail()
        self.current = iid
        self._update_undo()
        sections = [s for s in self._sections_of(doc_id) if self._path(s)[:len(prefix)] == prefix]
        book = self.tree.item(f"d{doc_id}", "text")
        title = HEADING_SEP.join(prefix) or book
        ctk.CTkLabel(self.detail, text=title, font=self.app.fonts["h2"], anchor="w", justify="left",
                     wraplength=560).grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 0))
        info = ctk.CTkFrame(self.detail, fg_color="transparent")
        info.grid(row=1, column=0, sticky="ew", padx=14, pady=(2, 0))
        meta = book if prefix else ""
        meta += f"{' · ' if meta else ''}{len(sections)} card{'s' * (len(sections) != 1)}"
        full_text = self._mode == self.MODES[0]
        if full_text and len(sections) > 1:
            meta += " · select a single header to edit it"
        ctk.CTkLabel(info, text=meta, text_color=MUTED, font=self.app.fonts["small"]).pack(side="left")
        if not sections:
            ctk.CTkLabel(self.detail, text="Nothing here.", text_color=MUTED).grid(row=2, column=0, pady=40)
            return
        self.speak_btn = ctk.CTkButton(info, text="🔊 Read", width=72, height=26)
        self.speak_btn.pack(side="right")

        if full_text and len(sections) == 1:
            editor = CardEditor(self.detail, self.app, sections[0], on_saved=lambda _sid: self._after_edit(True))
            editor.grid(row=2, column=0, sticky="nsew", padx=14, pady=(8, 12))
            self.editor = editor
            speak = lambda: (editor.get_heading(), editor.get_text())  # noqa: E731
        else:
            self._render_text(sections, prefix, lines=not full_text)
            if scroll is not None:
                self.lines_text.yview_moveto(scroll)
            text = "\n\n".join(s["content_text"] for s in sections if s["content_text"])
            speak = lambda: (title, text)  # noqa: E731
        self.speak_btn.configure(command=lambda b=self.speak_btn: self.app.toggle_speech(*speak(), b))

    def _render_text(self, sections: list[dict], prefix: tuple[str, ...], lines: bool):
        """All the text under the selected node, with sub-headers. ``lines``
        puts each sentence on its own row behind a clickable ✕."""
        box = ctk.CTkTextbox(self.detail, wrap="word", font=self.app.fonts["body"], border_spacing=8)
        box.grid(row=2, column=0, sticky="nsew", padx=14, pady=(8, 12))
        t: tk.Text = box._textbox  # tag fonts must be set on the Tk widget (CTk forbids them)
        self.lines_text = t
        try:
            scale = ctk.ScalingTracker.get_widget_scaling(self)
        except Exception:
            scale = 1.0
        px = lambda n: round(n * scale)  # noqa: E731
        t.tag_configure("head", font=(FONT_FAMILY, -px(16), "bold"), spacing1=px(14), spacing3=px(6))
        t.tag_configure("note", foreground="gray55", font=(FONT_FAMILY, -px(12), "italic"))
        t.tag_configure("line", lmargin2=px(30), spacing1=px(3), spacing3=px(3))  # wrapped text hangs past the ✕
        t.tag_configure("para", spacing3=px(10))
        t.tag_configure("x", foreground="gray55")
        t.tag_configure("hot", foreground=RED)
        t.tag_raise("hot")
        t.tag_bind("x", "<Enter>", lambda e: self._hover(e, True))
        t.tag_bind("x", "<Leave>", lambda e: self._hover(e, False))
        t.tag_bind("x", "<Button-1>", self._on_cross)

        last: tuple[str, ...] | None = None
        count = 0
        for s in sections:
            sub = self._path(s)[len(prefix):]
            if sub and sub != last:
                t.insert("end", HEADING_SEP.join(sub) + "\n", "head")
            last = sub
            if not lines:
                if s["content_text"]:
                    t.insert("end", s["content_text"] + "\n", "para")
            else:
                for i, line in enumerate(split_lines(s["content_text"] or "")):
                    if count >= self.MAX_LINES:
                        break
                    t.insert("end", "✕", ("x", f"L{count}", "line"))
                    t.insert("end", f"  {line}\n", "line")
                    self.line_refs.append((s["id"], i))
                    count += 1
            if s["image_paths"]:
                t.insert("end", f"🖼 {len(s['image_paths'])} image(s): select this header in Full text to see them\n",
                         "note")
            if count >= self.MAX_LINES:
                t.insert("end", f"\nShowing the first {self.MAX_LINES} lines. Select a chapter to see the rest.\n",
                         "note")
                break
        box.configure(state="disabled")

    def _line_at(self, event) -> int | None:
        t = self.lines_text
        if t is None:
            return None
        for tag in t.tag_names(t.index(f"@{event.x},{event.y}")):
            if tag.startswith("L") and tag[1:].isdigit():
                return int(tag[1:])
        return None

    def _hover(self, event, on: bool):
        t = self.lines_text
        if t is None:
            return
        t.tag_remove("hot", "1.0", "end")
        t.configure(cursor="hand2" if on else "xterm")
        n = self._line_at(event) if on else None
        if n is not None:
            t.tag_add("hot", *t.tag_ranges(f"L{n}")[:2])

    # -- edits ---------------------------------------------------------------
    def _on_cross(self, event):
        n = self._line_at(event)
        if n is None or n >= len(self.line_refs):
            return
        self.delete_line(n)

    def delete_line(self, n: int):
        """Delete the ``n``-th line shown; the card's last line removes the card."""
        sid, index = self.line_refs[n]
        s = self.kb.get_section(sid)
        if not s:
            return
        old = s["content_text"] or ""
        try:
            new = remove_line(old, index)
        except IndexError:  # text changed underneath; just redraw
            self._after_edit(False)
            return
        removed = not new and not s["image_paths"]
        if removed:  # Undo brings the card back too
            self.kb.soft_delete_section(sid)
            self.app.set_status("Deleted the card's last line, so the card was removed. Ctrl+Z undoes.")
        else:
            self.kb.update_section(sid, s["heading"] or "", new)
        self.undo_stack.append((sid, old, removed))
        del self.undo_stack[:-200]
        self._after_edit(removed)

    def undo(self):
        if not self.undo_stack:
            return
        sid, old, removed = self.undo_stack.pop()
        s = self.kb.get_section(sid)
        if not s:  # purged since
            self.app.set_status("That card no longer exists, so the line can't be restored.")
            self._update_undo()
            return
        self.kb.update_section(sid, s["heading"] or "", old)
        if removed:
            self.kb.restore_section(sid)
        self.app.set_status("Restored the deleted line.")
        self._after_edit(removed)

    def _update_undo(self):
        self.btn_undo.configure(state="normal" if self.undo_stack else "disabled")

    def _after_edit(self, structure_changed: bool):
        """Text changed here: tell the other views, then redraw this one."""
        self.editor = None  # saved already; nothing to ask about while the pane reloads
        self.app.notify_data_changed(source="read")
        if structure_changed:  # cards came or went, or a heading changed: rebuild the tree
            self.refresh()
            return
        self.dirty = False
        self._sections.clear()
        if self.current:
            self.show(self.current, keep_scroll=True)


# ---------------------------------------------------------------------------
# Library view
# ---------------------------------------------------------------------------
class LibraryView(ctk.CTkFrame):
    def __init__(self, master, app: App):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.dirty = True
        self._images: list = []
        self.speak_btn: ctk.CTkButton | None = None
        self.editor: CardEditor | None = None
        self.grid_columnconfigure(0, weight=2, uniform="lib")
        self.grid_columnconfigure(1, weight=3, uniform="lib")
        self.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(self)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.search = ctk.CTkEntry(bar, placeholder_text="Search headings and text…", width=240)
        self.search.pack(side="left", padx=(10, 6), pady=8)
        tk.Misc.bind(self.search, "<Return>", lambda e: self.refresh(), "+")
        ctk.CTkButton(bar, text="Search", width=80, command=self.refresh).pack(side="left")
        ctk.CTkButton(bar, text="Clear", width=60, fg_color="gray50", command=self._clear_search).pack(side="left", padx=6)
        self.show_deleted = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(bar, text="Show deleted", variable=self.show_deleted, command=self.refresh).pack(side="left", padx=12)
        ctk.CTkButton(bar, text="＋ Import…", width=100, command=app.import_documents).pack(side="right", padx=10)
        ctk.CTkButton(bar, text="📷 From photos…", width=120, command=app.cards_from_photos).pack(side="right")
        ctk.CTkButton(bar, text="✎ New book…", width=110, command=app.new_book).pack(side="right", padx=6)

        frame, self.tree = make_tree(self, {"info": ("Info", 160)}, "Documents & sections")
        frame.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.tree.bind("<<TreeviewOpen>>", self._on_open)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.detail = ctk.CTkFrame(self)
        self.detail.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        self.detail.grid_columnconfigure(0, weight=1)
        self._show_placeholder()

    # -- tree ----------------------------------------------------------------
    def on_show(self):
        if self.dirty:
            self.refresh()

    def _clear_search(self):
        self.search.delete(0, "end")
        self.refresh()

    def refresh(self):
        self.dirty = False
        selected = self.tree.selection()
        opened = {iid for iid in self.tree.get_children() if self.tree.item(iid, "open")}
        self.tree.delete(*self.tree.get_children())
        include_deleted = self.show_deleted.get()
        query = self.search.get().strip()
        docs = self.kb.list_documents()

        if query:
            hits = self.kb.search_sections(query, include_deleted)
            by_doc: dict[int, list[dict]] = {}
            for s in hits:
                by_doc.setdefault(s["document_id"], []).append(s)
            for d in docs:
                if d["id"] in by_doc:
                    self._insert_doc(d, open_=True)
                    for s in by_doc[d["id"]]:
                        self._insert_section(f"d{d['id']}", s)
            self.app.set_status(f"{len(hits)} matching section(s).")
        else:
            for d in docs:
                iid = self._insert_doc(d, open_=False)
                if iid in opened:
                    self._load_children(iid)
                    self.tree.item(iid, open=True)
                else:
                    self.tree.insert(iid, "end", iid=f"x{d['id']}", text="loading…")  # lazy placeholder

        if selected and self.tree.exists(selected[0]):
            self.tree.selection_set(selected[0])
            self.tree.see(selected[0])
        else:
            self._show_placeholder()

    def select(self, document_id: int, section_id: int | None = None):
        """Show a document (and optionally one of its sections) in the tree."""
        self.search.delete(0, "end")
        self.refresh()
        did = f"d{document_id}"
        if not self.tree.exists(did):
            return
        self._load_children(did)
        self.tree.item(did, open=True)
        sid = f"s{section_id}"
        iid = sid if section_id is not None and self.tree.exists(sid) else did
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)

    def _insert_doc(self, d: dict, open_: bool) -> str:
        n = d["active_count"] or 0
        fmt = "MY CARDS" if d["format"] == MANUAL_FORMAT else d["format"].upper()
        info = f"{fmt} · {n} card{'s' * (n != 1)}"
        if d["deleted_count"]:
            info += f" · {d['deleted_count']} del"
        return self.tree.insert("", "end", iid=f"d{d['id']}", text=d["title"], values=(info,),
                                open=open_, tags=("doc",))

    def _insert_section(self, parent: str, s: dict):
        label = s["heading"] or s["content_text"][:60]
        info = []
        if s["image_paths"]:
            info.append(f"{len(s['image_paths'])} img")
        if s["status"] == "deleted":
            info.append("deleted")
        self.tree.insert(parent, "end", iid=f"s{s['id']}", text=f"{s['sequence_order'] + 1}. {label}",
                         values=(" · ".join(info),), tags=("deleted",) if s["status"] == "deleted" else ())

    def _load_children(self, doc_iid: str):
        self.tree.delete(*self.tree.get_children(doc_iid))
        for s in self.kb.get_sections(int(doc_iid[1:]), include_deleted=self.show_deleted.get()):
            self._insert_section(doc_iid, s)

    def _on_open(self, _event):
        iid = self.tree.focus()
        if iid.startswith("d") and any(c.startswith("x") for c in self.tree.get_children(iid)):
            self._load_children(iid)

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        if iid.startswith("d"):
            self._show_document(int(iid[1:]))
        elif iid.startswith("s"):
            self._show_section(int(iid[1:]))

    # -- detail pane ---------------------------------------------------------
    def _clear_detail(self):
        if self.editor is not None:
            self.editor.confirm_leave(cancel=False)  # offer to save edits before they go
            self.editor = None
        if self.speak_btn is not None:
            self.app.stop_speech(self.speak_btn)  # only if this pane was reading
            self.speak_btn = None
        for w in self.detail.winfo_children():
            w.destroy()
        self._images.clear()

    def _show_placeholder(self):
        self._clear_detail()
        ctk.CTkLabel(self.detail, text="Select a document or section.", text_color=MUTED).grid(pady=40)

    def _buttons(self, specs: list[tuple[str, Callable, str | None]], row: int = 2) -> list[ctk.CTkButton]:
        bar = ctk.CTkFrame(self.detail, fg_color="transparent")
        bar.grid(row=row, column=0, sticky="w", padx=14, pady=(6, 0) if row == 2 else 6)
        buttons = []
        for text, cmd, color in specs:
            kwargs = {"fg_color": color} if color else {}
            btn = ctk.CTkButton(bar, text=text, width=10, height=30, command=cmd, **kwargs)
            btn.pack(side="left", padx=(0, 6))
            buttons.append(btn)
        return buttons

    def _show_document(self, doc_id: int):
        d = self.kb.get_document(doc_id)
        if not d:
            return self.refresh()
        self._clear_detail()
        ctk.CTkLabel(self.detail, text=d["title"], font=self.app.fonts["title"], anchor="w", justify="left",
                     wraplength=520).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 0))
        stats = self.kb.stats(document_id=doc_id)
        origin = ("your own cards, created" if d["format"] == MANUAL_FORMAT
                  else f"{d['format'].upper()}, imported")
        meta = (f"{d['author'] or 'Unknown author'} · {origin} {d['imported_at'][:10]}\n"
                f"{stats['total']} cards: {stats['new']} new, {stats['learning']} learning, "
                f"{stats['mastered']} mastered, {stats['due']} due\n{d['file_path'] or ''}")
        ctk.CTkLabel(self.detail, text=meta, text_color=MUTED, anchor="w", justify="left", wraplength=520).grid(
            row=1, column=0, sticky="w", padx=14)
        self._buttons([
            ("▶ Review", lambda: self._review_doc(doc_id), GREEN),
            ("＋ Add card", lambda: self.app.add_card(doc_id), None),
            ("📷 Add from photos", lambda: self.app.cards_from_photos(doc_id), None),
        ])
        self._buttons([
            ("Rename", lambda: self._rename(doc_id), None),
            ("Author", lambda: self._set_author(doc_id), None),
            ("+ Tag all", lambda: self.app.prompt_add_tag(document_id=doc_id), None),
            ("− Untag all", lambda: self._untag_doc(doc_id), None),
            ("Delete document", lambda: self._delete_doc(doc_id), RED),
        ], row=3)

    def _show_section(self, sid: int):
        s = self.kb.get_section(sid)
        if not s:
            return self.refresh()
        self._clear_detail()
        deleted = s["status"] == "deleted"
        srs = (f"{s['document_title']} · #{s['sequence_order'] + 1} · {s['review_status']} · "
               f"next {s['next_review_date']}" + (" · DELETED" if deleted else ""))
        ctk.CTkLabel(self.detail, text=srs, text_color=MUTED, anchor="w").grid(
            row=1, column=0, sticky="w", padx=14, pady=(14, 0))
        specs = [("🔊 Read", lambda: None, None), ("+ Tag", lambda: self._tag_section(sid), None),
                 ("＋ Card after", lambda: self._card_after(s), None)]
        if deleted:
            specs.append(("Restore", lambda: self._set_deleted(sid, False), GREEN))
        else:
            specs.append(("Delete", lambda: self._set_deleted(sid, True), "gray45"))
        specs.append(("Purge permanently", lambda: self._purge(sid), RED))
        self.speak_btn = self._buttons(specs)[0]

        tag_chips(self.detail, s["tags"], lambda n: self._untag_section(sid, n), lambda: self._tag_section(sid)).grid(
            row=3, column=0, sticky="w", padx=14, pady=(0, 6))
        # The card is edited right here; Save (or Ctrl+S) writes the changes.
        editor = CardEditor(self.detail, self.app, s, on_saved=lambda _sid: self._saved(), img_size=(320, 200))
        editor.grid(row=4, column=0, sticky="nsew", padx=14, pady=(0, 12))
        self.detail.grid_rowconfigure(4, weight=1)
        self.editor = editor
        self.speak_btn.configure(command=lambda b=self.speak_btn: self.app.toggle_speech(
            editor.get_heading(), editor.get_text(), b))

    # -- actions -------------------------------------------------------------
    def _changed(self):
        self.app.notify_data_changed(source="library")
        self.refresh()

    def _review_doc(self, doc_id: int):
        self.app.show_view("review")
        self.app.views["review"].set_filter(document_id=doc_id)

    def _rename(self, doc_id: int):
        name = ctk.CTkInputDialog(title="Rename", text="New title (category / book name):").get_input()
        if name and name.strip():
            self.kb.update_document(doc_id, title=name)
            self._changed()

    def _set_author(self, doc_id: int):
        author = ctk.CTkInputDialog(title="Author", text="Author:").get_input()
        if author is not None:
            self.kb.update_document(doc_id, author=author)
            self._changed()

    def _untag_doc(self, doc_id: int):
        name = ctk.CTkInputDialog(title="Remove tag", text="Remove which tag from every section?").get_input()
        if name and name.strip():
            n = self.kb.untag_document(doc_id, name)
            self.app.set_status(f"Removed #{name.strip()} from {n} section(s).")
            self._changed()

    def _delete_doc(self, doc_id: int):
        d = self.kb.get_document(doc_id)
        if d and messagebox.askyesno(
            "Delete document",
            f"Permanently delete “{d['title']}” with all its sections, review history and extracted images?\n\n"
            "This cannot be undone. (Your original file is not touched.)",
            icon="warning",
        ):
            self.kb.delete_document(doc_id)
            self._changed()

    def _saved(self):
        self.editor = None  # saved already; nothing to ask about while the pane reloads
        self._changed()

    def _card_after(self, s: dict):
        """Open a blank card that goes right after this one."""
        if self.editor is None or self.editor.confirm_leave():
            CardDialog(self, self.app, s["document_id"], on_saved=lambda _sid: self._changed(), after=s["id"])

    def _tag_section(self, sid: int):
        if self.app.prompt_add_tag(section_id=sid):
            self._show_section(sid)

    def _untag_section(self, sid: int, name: str):
        self.kb.remove_tag_from_section(sid, name)
        self.app.notify_data_changed(source="library")
        self._show_section(sid)

    def _set_deleted(self, sid: int, deleted: bool):
        (self.kb.soft_delete_section if deleted else self.kb.restore_section)(sid)
        self._changed()

    def _purge(self, sid: int):
        self.kb.purge_section(sid)
        self._changed()


# ---------------------------------------------------------------------------
# Tags view
# ---------------------------------------------------------------------------
class TagsView(ctk.CTkFrame):
    def __init__(self, master, app: App):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.dirty = True
        self.selected_id: int | None = None
        self.grid_columnconfigure(0, weight=1, uniform="tags")
        self.grid_columnconfigure(1, weight=1, uniform="tags")
        self.grid_rowconfigure(0, weight=1)

        frame, self.tree = make_tree(
            self, {"pattern": ("Pattern", 170), "type": ("Type", 70), "count": ("Sections", 70)}, "Tag", 160)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ctk.CTkFrame(self)
        form.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        form.grid_columnconfigure(0, weight=1)
        form.grid_rowconfigure(11, weight=1)
        ctk.CTkLabel(form, text="Tag", font=self.app.fonts["h2"]).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 4))
        self.name = ctk.CTkEntry(form, placeholder_text="Tag name, e.g. photosynthesis")
        self.name.grid(row=1, column=0, sticky="ew", padx=14)
        ctk.CTkLabel(form, text="Match pattern: a keyword or phrase (whole words, any case). Leave blank to "
                                "use the tag name. Tick the box for a regular expression.",
                     text_color=MUTED, wraplength=440, justify="left", font=self.app.fonts["small"]).grid(
            row=2, column=0, sticky="w", padx=14, pady=(8, 2))
        self.pattern = ctk.CTkEntry(form, placeholder_text="optional pattern")
        self.pattern.grid(row=3, column=0, sticky="ew", padx=14)
        self.is_regex = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(form, text="Regular expression", variable=self.is_regex).grid(
            row=4, column=0, sticky="w", padx=14, pady=8)

        row1 = ctk.CTkFrame(form, fg_color="transparent")
        row1.grid(row=5, column=0, sticky="w", padx=14)
        ctk.CTkButton(row1, text="Preview matches", width=10, command=self._preview).pack(side="left", padx=(0, 6))
        ctk.CTkButton(row1, text="Create tag + auto-apply", width=10, fg_color=GREEN, command=self._create).pack(side="left", padx=(0, 6))
        ctk.CTkButton(row1, text="New", width=10, fg_color="gray50", command=self._new).pack(side="left")

        ctk.CTkLabel(form, text="Selected tag", font=self.app.fonts["small_bold"]).grid(
            row=6, column=0, sticky="w", padx=14, pady=(14, 2))
        row2 = ctk.CTkFrame(form, fg_color="transparent")
        row2.grid(row=7, column=0, sticky="w", padx=14)
        self.sel_buttons = [
            ctk.CTkButton(row2, text="Save changes", width=10, command=self._save),
            ctk.CTkButton(row2, text="Auto-tag now", width=10, command=self._auto_one),
            ctk.CTkButton(row2, text="Remove auto-tags", width=10, fg_color="gray45", command=self._clear_auto),
            ctk.CTkButton(row2, text="Review", width=10, fg_color=GREEN, command=self._review),
            ctk.CTkButton(row2, text="Delete", width=10, fg_color=RED, command=self._delete),
        ]
        for b in self.sel_buttons:
            b.pack(side="left", padx=(0, 6))

        ctk.CTkButton(form, text="▶  Run auto-tagging (all tags, whole library)", height=40,
                      font=self.app.fonts["big"], command=self._auto_all).grid(
            row=8, column=0, sticky="ew", padx=14, pady=(18, 6))
        self.result = ctk.CTkLabel(form, text="", text_color=MUTED, anchor="w", justify="left", wraplength=440)
        self.result.grid(row=9, column=0, sticky="w", padx=14)
        self.match_label = ctk.CTkLabel(form, text="Matches", font=self.app.fonts["small_bold"])
        self.match_label.grid(row=10, column=0, sticky="w", padx=14)
        self.preview = ctk.CTkTextbox(form, wrap="word", font=self.app.fonts["small"])
        self.preview.grid(row=11, column=0, sticky="nsew", padx=14, pady=(2, 14))
        self.preview.configure(state="disabled")
        self._set_selected(None)

    def on_show(self):
        if self.dirty:
            self.refresh()

    def refresh(self):
        self.dirty = False
        self.tree.delete(*self.tree.get_children())
        for t in self.kb.list_tags():
            self.tree.insert("", "end", iid=str(t["id"]), text=f"#{t['name']}",
                             values=(t["pattern"] or "(name)", "regex" if t["is_regex"] else "keyword",
                                     t["section_count"]))
        if self.selected_id and self.tree.exists(str(self.selected_id)):
            self.tree.selection_set(str(self.selected_id))
        else:
            self._set_selected(None)

    def _set_selected(self, tag_id: int | None):
        self.selected_id = tag_id
        for b in self.sel_buttons:
            b.configure(state="normal" if tag_id else "disabled")

    def _on_select(self, _e):
        sel = self.tree.selection()
        if not sel:
            return
        tag = self.kb.get_tag_by_id(int(sel[0]))
        if not tag:
            return
        self._set_selected(tag["id"])
        self.name.delete(0, "end")
        self.name.insert(0, tag["name"])
        self.pattern.delete(0, "end")
        if tag["pattern"]:
            self.pattern.insert(0, tag["pattern"])
        self.is_regex.set(bool(tag["is_regex"]))
        self._preview()

    def _new(self):
        self.tree.selection_remove(*self.tree.selection())
        self._set_selected(None)
        self.name.delete(0, "end")
        self.pattern.delete(0, "end")
        self.is_regex.set(False)
        self._set_preview("")
        self.result.configure(text="")
        self.match_label.configure(text="Matches")

    def _set_preview(self, text: str):
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    def _preview(self):
        pattern = self.pattern.get().strip() or self.name.get().strip().lstrip("#")
        if not pattern:
            return
        try:
            count, rows = self.kb.preview_matches(pattern, self.is_regex.get(), limit=100)
        except ValueError as exc:
            self.result.configure(text=str(exc), text_color=RED)
            return
        self.match_label.configure(text=f"Matches: {count} active section(s)")
        lines = []
        for r in rows:
            snippet = " ".join(r["content_text"].split())[:140]
            lines.append(f"• {r['document_title']} — {r['heading'] or ''}\n   {snippet}…\n")
        if count > len(rows):
            lines.append(f"… and {count - len(rows)} more")
        self._set_preview("\n".join(lines))

    def _create(self):
        try:
            tag_id = self.kb.create_tag(self.name.get(), self.pattern.get(), self.is_regex.get())
        except ValueError as exc:
            self.result.configure(text=str(exc), text_color=RED)
            return
        applied = self.kb.auto_tag(tag_id)
        self.result.configure(text=f"Created tag and applied it to {applied} section(s).", text_color=GREEN)
        self.selected_id = tag_id
        self.app.notify_data_changed(source="tags")
        self.refresh()

    def _save(self):
        if not self.selected_id:
            return
        try:
            self.kb.update_tag(self.selected_id, self.name.get(), self.pattern.get(), self.is_regex.get())
        except ValueError as exc:
            self.result.configure(text=str(exc), text_color=RED)
            return
        self.result.configure(text="Saved. Run auto-tag to apply the new pattern.", text_color=GREEN)
        self.app.notify_data_changed(source="tags")
        self.refresh()

    def _auto_one(self):
        if self.selected_id:
            n = self.kb.auto_tag(self.selected_id)
            self.result.configure(text=f"Applied to {n} more section(s).", text_color=GREEN)
            self.app.notify_data_changed(source="tags")
            self.refresh()

    def _clear_auto(self):
        if self.selected_id and messagebox.askyesno(
                "Remove auto-tags", "Remove every automatically applied instance of this tag? Manual ones stay."):
            n = self.kb.clear_auto_tags(self.selected_id)
            self.result.configure(text=f"Removed from {n} section(s).", text_color=MUTED)
            self.app.notify_data_changed(source="tags")
            self.refresh()

    def _review(self):
        if self.selected_id:
            tag_id = self.selected_id
            self.app.show_view("review")
            self.app.views["review"].set_filter(tag_id=tag_id)

    def _delete(self):
        if self.selected_id and messagebox.askyesno("Delete tag", "Delete this tag from every section?"):
            self.kb.delete_tag(self.selected_id)
            self._new()
            self.app.notify_data_changed(source="tags")
            self.refresh()

    def _auto_all(self):
        results = self.kb.run_auto_tagging()
        if not results:
            self.result.configure(text="No tags yet. Create one first.", text_color=MUTED)
            return
        total = sum(results.values())
        detail = ", ".join(f"#{k}: +{v}" for k, v in results.items() if v) or "nothing new"
        self.result.configure(text=f"Auto-tagging done: {total} new tag(s). {detail}", text_color=GREEN)
        self.app.notify_data_changed(source="tags")
        self.refresh()


# ---------------------------------------------------------------------------
# Export view
# ---------------------------------------------------------------------------
CHECKED, UNCHECKED, PARTIAL = "☑", "☐", "◩"


class ExportView(ctk.CTkFrame):
    def __init__(self, master, app: App):
        super().__init__(master, fg_color="transparent")
        self.app, self.kb = app, app.kb
        self.dirty = True
        self.checked: set[int] = set()
        self.doc_sections: dict[int, list[int]] = {}
        self.labels: dict[str, str] = {}
        self._exporting = False
        self.grid_columnconfigure(0, weight=3, uniform="exp")
        self.grid_columnconfigure(1, weight=2, uniform="exp")
        self.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ctk.CTkLabel(bar, text="Pick books, chapters or sections", font=self.app.fonts["h2"]).pack(side="left")
        ctk.CTkButton(bar, text="Clear", width=60, fg_color="gray50", command=lambda: self._set_all(False)).pack(side="right")
        ctk.CTkButton(bar, text="Select all", width=80, command=lambda: self._set_all(True)).pack(side="right", padx=6)

        frame, self.tree = make_tree(self, {"info": ("Status", 110)}, "Click to check / uncheck")
        frame.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.tree.bind("<ButtonRelease-1>", self._on_click)
        self.tree.bind("<space>", self._on_space)

        opts = ctk.CTkFrame(self)
        opts.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(6, 0))
        opts.grid_columnconfigure(0, weight=1)
        r = 0

        def label(text):
            nonlocal r
            ctk.CTkLabel(opts, text=text, font=self.app.fonts["small_bold"]).grid(
                row=r, column=0, sticky="w", padx=14, pady=(12, 2))
            r += 1

        label("Scope")
        self.scope = ctk.StringVar(value="checked")
        for value, text in (("checked", "Checked items only"), ("all", "Entire knowledge base")):
            ctk.CTkRadioButton(opts, text=text, variable=self.scope, value=value).grid(row=r, column=0, sticky="w", padx=24, pady=2)
            r += 1

        label("Content filter")
        self.review_filter = ctk.CTkOptionMenu(
            opts, values=["All active sections", "Reviewed: known only", "Mastered only"])
        self.review_filter.grid(row=r, column=0, sticky="ew", padx=14)
        r += 1
        self.tag_menu = ctk.CTkOptionMenu(opts, values=["Any tag"])
        self.tag_menu.grid(row=r, column=0, sticky="ew", padx=14, pady=(6, 0))
        r += 1

        label("Format")
        self.fmt = ctk.StringVar(value="pdf")
        fmt_row = ctk.CTkFrame(opts, fg_color="transparent")
        fmt_row.grid(row=r, column=0, sticky="w", padx=24)
        r += 1
        for value, text in (("pdf", "PDF"), ("docx", "Word"), ("md", "Markdown"), ("txt", "Text")):
            ctk.CTkRadioButton(fmt_row, text=text, variable=self.fmt, value=value, width=10).pack(
                side="left", padx=(0, 16))

        label("Options")
        self.title_entry = ctk.CTkEntry(opts)
        self.title_entry.insert(0, "Memory Palace Export")
        self.title_entry.grid(row=r, column=0, sticky="ew", padx=14)
        r += 1
        self.include_images = ctk.BooleanVar(value=True)
        self.include_tags = ctk.BooleanVar(value=True)
        self.include_headings = ctk.BooleanVar(value=True)
        for var, text in ((self.include_images, "Embed images"), (self.include_tags, "Include tags"),
                          (self.include_headings, "Include section headings")):
            ctk.CTkCheckBox(opts, text=text, variable=var).grid(row=r, column=0, sticky="w", padx=24, pady=3)
            r += 1

        self.summary = ctk.CTkLabel(opts, text="", text_color=MUTED, justify="left", anchor="w")
        self.summary.grid(row=r, column=0, sticky="w", padx=14, pady=(14, 4))
        r += 1
        self.export_btn = ctk.CTkButton(opts, text="📤  Export…", height=44, font=self.app.fonts["big"],
                                        command=self._export)
        self.export_btn.grid(row=r, column=0, sticky="ew", padx=14, pady=6)
        r += 1
        self.bar = ctk.CTkProgressBar(opts)
        self.bar.set(0)
        self.bar.grid(row=r, column=0, sticky="ew", padx=14, pady=6)

    # -- tree with checkboxes ------------------------------------------------
    def on_show(self):
        if self.dirty:
            self.refresh()

    def refresh(self):
        self.dirty = False
        opened = {iid for iid in self.tree.get_children() if self.tree.item(iid, "open")}
        self.tree.delete(*self.tree.get_children())
        self.doc_sections.clear()
        self.labels.clear()
        valid: set[int] = set()
        review = self.kb.review_statuses()
        for d in self.kb.list_documents():
            did = f"d{d['id']}"
            self.labels[did] = d["title"]
            self.tree.insert("", "end", iid=did, text="", values=(f"{d['active_count'] or 0} card{'s' * ((d['active_count'] or 0) != 1)}",),
                             open=did in opened, tags=("doc",))
            ids = []
            for s in self.kb.get_sections(d["id"]):
                sid = f"s{s['id']}"
                self.labels[sid] = f"{s['sequence_order'] + 1}. {s['heading'] or s['content_text'][:60]}"
                self.tree.insert(did, "end", iid=sid, text="", values=(review.get(s["id"], ""),))
                ids.append(s["id"])
                valid.add(s["id"])
            self.doc_sections[d["id"]] = ids
        self.checked &= valid
        for did in self.doc_sections:
            self._redraw_doc(did)

        tags = {"Any tag": None}
        tags.update({f"#{t['name']}": t["id"] for t in self.kb.list_tags()})
        self._tag_choices = tags
        current = self.tag_menu.get()
        self.tag_menu.configure(values=list(tags))
        if current not in tags:
            self.tag_menu.set("Any tag")
        self._update_summary()

    def _redraw_doc(self, doc_id: int):
        ids = self.doc_sections.get(doc_id, [])
        n = sum(1 for i in ids if i in self.checked)
        mark = CHECKED if ids and n == len(ids) else PARTIAL if n else UNCHECKED
        did = f"d{doc_id}"
        self.tree.item(did, text=f"{mark}  {self.labels[did]}")
        for sid in ids:
            iid = f"s{sid}"
            self.tree.item(iid, text=f"{CHECKED if sid in self.checked else UNCHECKED}  {self.labels[iid]}")

    def _toggle(self, iid: str):
        if iid.startswith("d"):
            doc_id = int(iid[1:])
            ids = self.doc_sections.get(doc_id, [])
            if ids and all(i in self.checked for i in ids):
                self.checked.difference_update(ids)
            else:
                self.checked.update(ids)
        else:
            sid = int(iid[1:])
            self.checked.symmetric_difference_update({sid})
            doc_id = int(self.tree.parent(iid)[1:])
        self._redraw_doc(doc_id)
        self.scope.set("checked")
        self._update_summary()

    def _on_click(self, event):
        if "indicator" in self.tree.identify_element(event.x, event.y):
            return  # expand/collapse arrow
        iid = self.tree.identify_row(event.y)
        if iid and self.tree.identify_region(event.x, event.y) in ("tree", "cell"):
            self._toggle(iid)

    def _on_space(self, _event):
        iid = self.tree.focus()
        if iid:
            self._toggle(iid)
        return "break"

    def _set_all(self, value: bool):
        self.checked = {i for ids in self.doc_sections.values() for i in ids} if value else set()
        for did in self.doc_sections:
            self._redraw_doc(did)
        self._update_summary()

    def _update_summary(self):
        self.summary.configure(text=f"{len(self.checked)} section(s) checked")

    # -- export --------------------------------------------------------------
    def _export(self):
        if self._exporting:
            return
        scope_all = self.scope.get() == "all"
        if not scope_all and not self.checked:
            messagebox.showinfo("Export", "Check some books or sections first, or choose “Entire knowledge base”.")
            return
        review_filter = {"All active sections": "all", "Reviewed: known only": "known",
                         "Mastered only": "mastered"}[self.review_filter.get()]
        groups = self.kb.get_export_groups(
            section_ids=None if scope_all else set(self.checked),
            review_filter=review_filter,
            tag_id=self._tag_choices.get(self.tag_menu.get()),
        )
        n = sum(len(g.sections) for g in groups)
        if not n:
            messagebox.showinfo("Export", "No sections match these filters.")
            return
        fmt = self.fmt.get()
        ext = EXPORT_FORMATS[fmt]
        title = self.title_entry.get().strip() or "Memory Palace Export"
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in title).strip() or "export"
        out = filedialog.asksaveasfilename(
            title="Export to…", defaultextension=ext, initialfile=safe_name + ext,
            filetypes=[(fmt.upper(), f"*{ext}"), ("All files", "*.*")],
        )
        if not out:
            return
        options = ExportOptions(title, self.include_images.get(), self.include_tags.get(), self.include_headings.get())
        self._exporting = True
        self.export_btn.configure(state="disabled", text="Exporting…")
        self.bar.set(0)
        self.app.set_status(f"Exporting {n} section(s)…")

        def progress(done: int, total: int):
            if done == total or done % 20 == 0:
                self.app.post(lambda: self.bar.set(done / max(total, 1)))

        def finish():
            self._exporting = False
            self.export_btn.configure(state="normal", text="📤  Export…")

        def ok(path: Path):
            finish()
            self.bar.set(1)
            self.app.set_status(f"Exported {n} section(s) to {path.name}.")
            if messagebox.askyesno("Export complete", f"Wrote {n} section(s) to:\n{path}\n\nOpen it now?"):
                open_with_system(path)

        def failed(exc: BaseException, tb: str):
            finish()
            if not isinstance(exc, ExportError):
                traceback.print_exception(exc)
            messagebox.showerror("Export failed", str(exc))

        self.app.run_background(lambda: export(groups, out, fmt, options, progress), ok, failed)
