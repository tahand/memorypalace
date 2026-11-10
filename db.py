"""
db.py - SQLite persistence layer for Memory Palace.

Responsibilities
----------------
* Schema creation and versioning.
* CRUD for documents, sections and tags.
* Keyword / regex auto-tagging. Matching runs *inside* SQLite through a
  registered ``regexp()`` user function, so tagging a large knowledge base
  is a single ``INSERT ... SELECT`` rather than a Python loop.
* SuperMemo-2 (SM-2) spaced-repetition scheduling.
* Queries that feed the review queue and the export engine.

Image paths stored in the database are *relative to the data directory* and
always use forward slashes, so the whole ``data`` folder can be moved,
synced or backed up as a unit.
"""
from __future__ import annotations

import json
import re
import shutil
import sqlite3
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:  # pragma: no cover - only for type hints
    from parser import ParsedDocument

SCHEMA_VERSION = 1

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    title        TEXT    NOT NULL,              -- category / book name
    author       TEXT,
    format       TEXT    NOT NULL,              -- pdf, epub, docx, doc, txt, md
    file_path    TEXT,                          -- original location on disk
    file_hash    TEXT,                          -- sha256, used to detect re-imports
    media_dir    TEXT,                          -- folder of extracted images (relative)
    imported_at  TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_documents_hash ON documents(file_hash);

CREATE TABLE IF NOT EXISTS sections (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id     INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    sequence_order  INTEGER NOT NULL,
    heading         TEXT,                       -- breadcrumb, e.g. "Ch. 2 > Cells"
    content_text    TEXT    NOT NULL DEFAULT '',
    image_paths     TEXT    NOT NULL DEFAULT '[]',  -- JSON list of relative paths
    status          TEXT    NOT NULL DEFAULT 'active'
                    CHECK (status IN ('active', 'deleted'))
);
CREATE INDEX IF NOT EXISTS idx_sections_doc    ON sections(document_id, sequence_order);
CREATE INDEX IF NOT EXISTS idx_sections_status ON sections(status);

CREATE TABLE IF NOT EXISTS tags (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    name      TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    pattern   TEXT,                             -- NULL = match the tag name itself
    is_regex  INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS section_tags (
    section_id INTEGER NOT NULL REFERENCES sections(id) ON DELETE CASCADE,
    tag_id     INTEGER NOT NULL REFERENCES tags(id)     ON DELETE CASCADE,
    source     TEXT    NOT NULL DEFAULT 'manual' CHECK (source IN ('manual', 'auto')),
    PRIMARY KEY (section_id, tag_id)
);
CREATE INDEX IF NOT EXISTS idx_section_tags_tag ON section_tags(tag_id);

CREATE TABLE IF NOT EXISTS reviews (
    section_id        INTEGER PRIMARY KEY REFERENCES sections(id) ON DELETE CASCADE,
    ease_factor       REAL    NOT NULL DEFAULT 2.5,
    interval          INTEGER NOT NULL DEFAULT 0,   -- days
    repetitions       INTEGER NOT NULL DEFAULT 0,   -- consecutive successful recalls
    next_review_date  TEXT    NOT NULL,             -- ISO date (YYYY-MM-DD)
    last_reviewed_at  TEXT,                         -- ISO datetime
    status            TEXT    NOT NULL DEFAULT 'new'
                      CHECK (status IN ('new', 'learning', 'mastered'))
);
CREATE INDEX IF NOT EXISTS idx_reviews_due ON reviews(next_review_date);

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""

# ``documents.format`` for books created by hand inside the app.
MANUAL_FORMAT = "manual"


class DatabaseError(Exception):
    """Raised when the database cannot be opened or is corrupted."""


# ---------------------------------------------------------------------------
# SuperMemo-2
# ---------------------------------------------------------------------------
QUALITY_KNOWN = 4      # "I know": correct response after some hesitation
QUALITY_UNKNOWN = 1    # "I don't know": incorrect response
MIN_EASE = 1.3
DEFAULT_EASE = 2.5
MASTERED_INTERVAL = 21  # days; cards scheduled this far out count as mastered


@dataclass(frozen=True)
class ReviewState:
    ease_factor: float = DEFAULT_EASE
    interval: int = 0
    repetitions: int = 0


def sm2(state: ReviewState, quality: int) -> ReviewState:
    """Apply one SM-2 step.

    ``quality`` is the 0-5 SM-2 grade. Grades >= 3 count as a successful
    recall. A failed recall resets the repetition count and gives an
    interval of 0, so the card is due again today and goes to the back of
    the current queue.
    """
    if not 0 <= quality <= 5:
        raise ValueError("SM-2 quality must be between 0 and 5")

    ease = state.ease_factor + (0.1 - (5 - quality) * (0.08 + (5 - quality) * 0.02))
    ease = max(MIN_EASE, ease)

    if quality >= 3:
        if state.repetitions == 0:
            interval = 1
        elif state.repetitions == 1:
            interval = 6
        else:
            interval = max(1, round(state.interval * state.ease_factor))
        repetitions = state.repetitions + 1
    else:
        interval = 0
        repetitions = 0

    return ReviewState(round(ease, 4), interval, repetitions)


def review_status(state: ReviewState, reviewed: bool) -> str:
    if not reviewed:
        return "new"
    return "mastered" if state.interval >= MASTERED_INTERVAL else "learning"


# ---------------------------------------------------------------------------
# Regex helpers
# ---------------------------------------------------------------------------
@lru_cache(maxsize=512)
def _compile(pattern: str) -> re.Pattern[str]:
    return re.compile(pattern, re.IGNORECASE | re.MULTILINE)


def _sqlite_regexp(pattern: str | None, value: str | None) -> bool:
    """SQLite user function: regexp(pattern, value) -> 0/1."""
    if pattern is None or value is None:
        return False
    try:
        return _compile(pattern).search(value) is not None
    except re.error:
        return False


def keyword_to_regex(keyword: str) -> str:
    """Turn a keyword or phrase into a case-insensitive whole-word regex.

    ``"cell wall"`` matches "Cell wall" and a phrase wrapped across lines,
    but not "cell walls" or "intracell wall".
    """
    words = keyword.split()
    if not words:
        raise ValueError("Keyword is empty")
    body = r"\s+".join(re.escape(w) for w in words)
    return rf"(?<!\w){body}(?!\w)"


def validate_regex(pattern: str) -> None:
    try:
        re.compile(pattern)
    except re.error as exc:
        raise ValueError(f"Invalid regular expression: {exc}") from exc


def normalize_tag_name(name: str) -> str:
    name = (name or "").strip().lstrip("#").strip()
    name = re.sub(r"\s+", " ", name)
    if not name:
        raise ValueError("Tag name cannot be empty")
    if len(name) > 80:
        raise ValueError("Tag name is too long (max 80 characters)")
    return name


# ---------------------------------------------------------------------------
# Export data containers
# ---------------------------------------------------------------------------
@dataclass
class ExportSection:
    id: int
    heading: str
    content_text: str
    image_paths: list[Path]          # absolute paths, existing files only
    tags: list[str] = field(default_factory=list)


@dataclass
class ExportGroup:
    document: dict[str, Any]
    sections: list[ExportSection]


# ---------------------------------------------------------------------------
# Knowledge base
# ---------------------------------------------------------------------------
# Text a tag pattern is matched against: the heading breadcrumb plus the body.
_MATCH_TEXT = "COALESCE(s.heading, '') || char(10) || s.content_text"


class KnowledgeBase:
    """All database access goes through this class.

    The connection is not shared across threads. The UI does slow work such
    as parsing and exporting in worker threads and then calls back into the
    knowledge base on the Tk main thread.
    """

    def __init__(self, data_dir: str | Path = "data"):
        self.data_dir = Path(data_dir).resolve()
        self.media_root = self.data_dir / "media"
        try:
            self.media_root.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise DatabaseError(f"Cannot create data directory {self.data_dir}: {exc}") from exc

        self.db_path = self.data_dir / "memorypalace.db"
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.conn.row_factory = sqlite3.Row
            self.conn.create_function("regexp", 2, _sqlite_regexp, deterministic=True)
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.conn.execute("PRAGMA journal_mode = WAL")
            self._init_schema()
        except sqlite3.DatabaseError as exc:
            raise DatabaseError(
                f"Could not open the knowledge base at {self.db_path}.\n"
                f"The file may be corrupted. Move it aside to start fresh.\n\nDetails: {exc}"
            ) from exc

    # -- lifecycle ---------------------------------------------------------
    def _init_schema(self) -> None:
        with self.conn:
            self.conn.executescript(SCHEMA)
            version = self.conn.execute("PRAGMA user_version").fetchone()[0]
            if version < SCHEMA_VERSION:
                # Future migrations go here, keyed on ``version``.
                self.conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")

    def close(self) -> None:
        try:
            self.conn.close()
        except sqlite3.Error:
            pass

    # -- paths -------------------------------------------------------------
    def new_media_dir(self) -> tuple[Path, str]:
        """Return (absolute, relative) paths for a new document's image folder.

        The folder itself is created lazily by the parser on the first image.
        """
        rel = f"media/{uuid.uuid4().hex}"
        return self.data_dir / rel, rel

    def abs_path(self, rel: str) -> Path:
        return self.data_dir / rel

    def save_image(self, document_id: int, data: bytes, ext: str) -> str:
        """Store image bytes in a document's media folder; return the relative path."""
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError(f"No document with id {document_id}")
        rel_dir = doc.get("media_dir")
        if not rel_dir:
            _, rel_dir = self.new_media_dir()
            with self.conn:
                self.conn.execute("UPDATE documents SET media_dir = ? WHERE id = ?", (rel_dir, document_id))
        ext = ext.lower().lstrip(".") or "png"
        rel = f"{rel_dir}/card_{uuid.uuid4().hex[:12]}.{'jpg' if ext == 'jpeg' else ext}"
        path = self.data_dir / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return rel

    def _delete_unused_images(self, rel_paths: Iterable[str]) -> None:
        """Delete image files that no section refers to any more."""
        for rel in rel_paths:
            shared = self.conn.execute(
                "SELECT 1 FROM sections, json_each(sections.image_paths) WHERE json_each.value = ? LIMIT 1",
                (rel,),
            ).fetchone()
            if not shared:
                try:
                    (self.data_dir / rel).unlink(missing_ok=True)
                except OSError:
                    pass

    # -- settings ----------------------------------------------------------
    def get_setting(self, key: str, default: str | None = None) -> str | None:
        row = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row[0] if row and row[0] is not None else default

    def set_setting(self, key: str, value: str | None) -> None:
        with self.conn:
            self.conn.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?)"
                " ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, value),
            )

    # -- documents ---------------------------------------------------------
    def add_document(
        self,
        parsed: "ParsedDocument",
        file_path: str | None = None,
        file_hash: str | None = None,
        media_dir: str | None = None,
        auto_tag: bool = True,
    ) -> tuple[int, int]:
        """Insert a parsed document with its sections and review rows.

        Returns ``(document_id, auto_tag_applications)``. With ``auto_tag``
        set, every existing tag is matched against the new sections.
        """
        now = datetime.now().isoformat(timespec="seconds")
        today = date.today().isoformat()
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO documents (title, author, format, file_path, file_hash, media_dir, imported_at)"
                " VALUES (?, ?, ?, ?, ?, ?, ?)",
                (parsed.title, parsed.author, parsed.format, file_path, file_hash, media_dir, now),
            )
            doc_id = cur.lastrowid
            for order, section in enumerate(parsed.sections):
                cur = self.conn.execute(
                    "INSERT INTO sections (document_id, sequence_order, heading, content_text, image_paths)"
                    " VALUES (?, ?, ?, ?, ?)",
                    (doc_id, order, section.heading, section.content_text, json.dumps(section.image_paths)),
                )
                self.conn.execute(
                    "INSERT INTO reviews (section_id, next_review_date) VALUES (?, ?)",
                    (cur.lastrowid, today),
                )
        applied = 0
        if auto_tag:
            for tag in self.list_tags():
                applied += self.auto_tag(tag["id"], document_id=doc_id)
        return doc_id, applied

    def create_document(self, title: str, author: str | None = None) -> int:
        """Create an empty book to fill with hand-made cards."""
        title = (title or "").strip()
        if not title:
            raise ValueError("Title cannot be empty")
        _, media_rel = self.new_media_dir()
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO documents (title, author, format, media_dir, imported_at) VALUES (?, ?, ?, ?, ?)",
                (title, (author or "").strip() or None, MANUAL_FORMAT, media_rel,
                 datetime.now().isoformat(timespec="seconds")),
            )
        return cur.lastrowid

    def add_section(
        self,
        document_id: int,
        heading: str,
        content_text: str,
        image_paths: Iterable[str] = (),
        auto_tag: bool = True,
        after: int | None = None,
    ) -> int:
        """Add a card to a document, at the end or right after section ``after``.
        Returns the new section id."""
        content_text = (content_text or "").strip()
        image_paths = list(image_paths)
        if not content_text and not image_paths:
            raise ValueError("A card needs some text or at least one image")
        if not self.get_document(document_id):
            raise ValueError(f"No document with id {document_id}")
        prev = self.conn.execute(
            "SELECT sequence_order FROM sections WHERE id = ? AND document_id = ?", (after, document_id)
        ).fetchone() if after is not None else None
        with self.conn:
            if prev:
                order = prev["sequence_order"] + 1
                self.conn.execute(
                    "UPDATE sections SET sequence_order = sequence_order + 1"
                    " WHERE document_id = ? AND sequence_order >= ?",
                    (document_id, order),
                )
            else:
                order = self.conn.execute(
                    "SELECT COALESCE(MAX(sequence_order) + 1, 0) FROM sections WHERE document_id = ?",
                    (document_id,),
                ).fetchone()[0]
            cur = self.conn.execute(
                "INSERT INTO sections (document_id, sequence_order, heading, content_text, image_paths)"
                " VALUES (?, ?, ?, ?, ?)",
                (document_id, order, (heading or "").strip() or None, content_text, json.dumps(image_paths)),
            )
            section_id = cur.lastrowid
            self.conn.execute(
                "INSERT INTO reviews (section_id, next_review_date) VALUES (?, ?)",
                (section_id, date.today().isoformat()),
            )
        if auto_tag:
            for tag in self.list_tags():
                self._auto_tag_section(tag, section_id)
        return section_id

    def _auto_tag_section(self, tag: dict, section_id: int) -> None:
        with self.conn:
            self.conn.execute(
                "INSERT OR IGNORE INTO section_tags (section_id, tag_id, source)"
                f" SELECT s.id, ?, 'auto' FROM sections s WHERE s.id = ? AND regexp(?, {_MATCH_TEXT})",
                (tag["id"], section_id, self.tag_regex(tag)),
            )

    def find_document_by_hash(self, file_hash: str) -> dict | None:
        row = self.conn.execute(
            "SELECT * FROM documents WHERE file_hash = ? LIMIT 1", (file_hash,)
        ).fetchone()
        return dict(row) if row else None

    def list_documents(self) -> list[dict]:
        rows = self.conn.execute(
            """
            SELECT d.*,
                   SUM(CASE WHEN s.status = 'active'  THEN 1 ELSE 0 END) AS active_count,
                   SUM(CASE WHEN s.status = 'deleted' THEN 1 ELSE 0 END) AS deleted_count
            FROM documents d LEFT JOIN sections s ON s.document_id = d.id
            GROUP BY d.id
            ORDER BY d.title COLLATE NOCASE, d.id
            """
        ).fetchall()
        return [dict(r) for r in rows]

    def get_document(self, document_id: int) -> dict | None:
        row = self.conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        return dict(row) if row else None

    def update_document(self, document_id: int, title: str | None = None, author: str | None = None) -> None:
        with self.conn:
            if title is not None:
                title = title.strip()
                if not title:
                    raise ValueError("Title cannot be empty")
                self.conn.execute("UPDATE documents SET title = ? WHERE id = ?", (title, document_id))
            if author is not None:
                self.conn.execute(
                    "UPDATE documents SET author = ? WHERE id = ?", (author.strip() or None, document_id)
                )

    def delete_document(self, document_id: int) -> None:
        """Permanently remove a document, its sections, reviews and images."""
        doc = self.get_document(document_id)
        if not doc:
            return
        with self.conn:
            self.conn.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        if doc.get("media_dir"):
            media = (self.data_dir / doc["media_dir"]).resolve()
            # Only delete inside our own media root.
            if self.media_root in media.parents:
                shutil.rmtree(media, ignore_errors=True)

    # -- sections ----------------------------------------------------------
    @staticmethod
    def _section_dict(row: sqlite3.Row) -> dict:
        d = dict(row)
        try:
            d["image_paths"] = json.loads(d.get("image_paths") or "[]")
        except (TypeError, json.JSONDecodeError):
            d["image_paths"] = []
        return d

    def get_sections(self, document_id: int, include_deleted: bool = False) -> list[dict]:
        sql = "SELECT * FROM sections WHERE document_id = ?"
        if not include_deleted:
            sql += " AND status = 'active'"
        sql += " ORDER BY sequence_order"
        return [self._section_dict(r) for r in self.conn.execute(sql, (document_id,))]

    def get_section(self, section_id: int) -> dict | None:
        """One section with its document title, review state and tag names."""
        row = self.conn.execute(
            """
            SELECT s.*, d.title AS document_title, d.author AS document_author,
                   r.ease_factor, r.interval, r.repetitions, r.next_review_date,
                   r.last_reviewed_at, r.status AS review_status
            FROM sections s
            JOIN documents d ON d.id = s.document_id
            LEFT JOIN reviews r ON r.section_id = s.id
            WHERE s.id = ?
            """,
            (section_id,),
        ).fetchone()
        if not row:
            return None
        d = self._section_dict(row)
        d["tags"] = self.section_tags(section_id)
        return d

    def search_sections(self, query: str, include_deleted: bool = False, limit: int = 1000) -> list[dict]:
        """Case-insensitive substring search over headings and text."""
        like = f"%{query.strip()}%"
        sql = (
            "SELECT s.* FROM sections s WHERE (s.content_text LIKE ? OR s.heading LIKE ?)"
            + ("" if include_deleted else " AND s.status = 'active'")
            + " ORDER BY s.document_id, s.sequence_order LIMIT ?"
        )
        return [self._section_dict(r) for r in self.conn.execute(sql, (like, like, limit))]

    def update_section(
        self, section_id: int, heading: str, content_text: str, image_paths: list[str] | None = None
    ) -> None:
        """Edit a card. With ``image_paths``, also replace its images; files no
        other card uses are deleted."""
        with self.conn:
            self.conn.execute(
                "UPDATE sections SET heading = ?, content_text = ? WHERE id = ?",
                (heading.strip() or None, content_text.strip(), section_id),
            )
            if image_paths is None:
                return
            row = self.conn.execute("SELECT image_paths FROM sections WHERE id = ?", (section_id,)).fetchone()
            old = json.loads(row["image_paths"] or "[]") if row else []
            self.conn.execute(
                "UPDATE sections SET image_paths = ? WHERE id = ?", (json.dumps(list(image_paths)), section_id)
            )
        self._delete_unused_images(set(old) - set(image_paths))

    def split_section(self, section_id: int, before: str, after: str, heading: str | None = None) -> int:
        """Split a card in two: it keeps ``before`` and its images, and a new
        card with ``after`` is placed right after it with the same heading and
        manual tags. Returns the new section id."""
        before, after = (before or "").strip(), (after or "").strip()
        if not before or not after:
            raise ValueError("Both parts of a split card need some text")
        sec = self.get_section(section_id)
        if not sec:
            raise ValueError(f"No section with id {section_id}")
        heading = (sec["heading"] or "") if heading is None else heading
        self.update_section(section_id, heading, before)
        new_id = self.add_section(sec["document_id"], heading, after, after=section_id)
        with self.conn:
            self.conn.execute(
                "INSERT OR IGNORE INTO section_tags (section_id, tag_id, source)"
                " SELECT ?, tag_id, 'manual' FROM section_tags WHERE section_id = ? AND source = 'manual'",
                (new_id, section_id),
            )
        return new_id

    def soft_delete_section(self, section_id: int) -> None:
        with self.conn:
            self.conn.execute("UPDATE sections SET status = 'deleted' WHERE id = ?", (section_id,))

    def restore_section(self, section_id: int) -> None:
        with self.conn:
            self.conn.execute("UPDATE sections SET status = 'active' WHERE id = ?", (section_id,))

    def purge_section(self, section_id: int) -> None:
        """Permanently delete a section and any images no other section uses."""
        row = self.conn.execute("SELECT image_paths FROM sections WHERE id = ?", (section_id,)).fetchone()
        if not row:
            return
        images = json.loads(row["image_paths"] or "[]")
        with self.conn:
            self.conn.execute("DELETE FROM sections WHERE id = ?", (section_id,))
        self._delete_unused_images(images)

    # -- tags --------------------------------------------------------------
    def list_tags(self) -> list[dict]:
        rows = self.conn.execute(
            """
            SELECT t.*, COUNT(s.id) AS section_count
            FROM tags t
            LEFT JOIN section_tags st ON st.tag_id = t.id
            LEFT JOIN sections s ON s.id = st.section_id AND s.status = 'active'
            GROUP BY t.id ORDER BY t.name COLLATE NOCASE
            """
        ).fetchall()
        return [dict(r) for r in rows]

    def get_tag(self, name: str) -> dict | None:
        row = self.conn.execute("SELECT * FROM tags WHERE name = ?", (name.strip().lstrip("#"),)).fetchone()
        return dict(row) if row else None

    def get_tag_by_id(self, tag_id: int) -> dict | None:
        row = self.conn.execute("SELECT * FROM tags WHERE id = ?", (tag_id,)).fetchone()
        return dict(row) if row else None

    def create_tag(self, name: str, pattern: str | None = None, is_regex: bool = False) -> int:
        """Create a tag. Raises ``ValueError`` if it exists or the pattern is invalid."""
        name = normalize_tag_name(name)
        pattern = (pattern or "").strip() or None
        if is_regex:
            if not pattern:
                raise ValueError("A regex tag needs a pattern")
            validate_regex(pattern)
        if self.get_tag(name):
            raise ValueError(f"Tag '{name}' already exists")
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO tags (name, pattern, is_regex) VALUES (?, ?, ?)", (name, pattern, int(is_regex))
            )
        return cur.lastrowid

    def _ensure_tag(self, name: str) -> int:
        name = normalize_tag_name(name)
        tag = self.get_tag(name)
        return tag["id"] if tag else self.create_tag(name)

    def update_tag(self, tag_id: int, name: str, pattern: str | None, is_regex: bool) -> None:
        name = normalize_tag_name(name)
        pattern = (pattern or "").strip() or None
        if is_regex:
            if not pattern:
                raise ValueError("A regex tag needs a pattern")
            validate_regex(pattern)
        other = self.get_tag(name)
        if other and other["id"] != tag_id:
            raise ValueError(f"Tag '{name}' already exists")
        with self.conn:
            self.conn.execute(
                "UPDATE tags SET name = ?, pattern = ?, is_regex = ? WHERE id = ?",
                (name, pattern, int(is_regex), tag_id),
            )

    def delete_tag(self, tag_id: int) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM tags WHERE id = ?", (tag_id,))

    def section_tags(self, section_id: int) -> list[str]:
        rows = self.conn.execute(
            "SELECT t.name FROM tags t JOIN section_tags st ON st.tag_id = t.id"
            " WHERE st.section_id = ? ORDER BY t.name COLLATE NOCASE",
            (section_id,),
        )
        return [r[0] for r in rows]

    def add_tag_to_section(self, section_id: int, name: str) -> int:
        tag_id = self._ensure_tag(name)
        with self.conn:
            # A manual tag overrides an earlier automatic one.
            self.conn.execute(
                "INSERT INTO section_tags (section_id, tag_id, source) VALUES (?, ?, 'manual')"
                " ON CONFLICT(section_id, tag_id) DO UPDATE SET source = 'manual'",
                (section_id, tag_id),
            )
        return tag_id

    def remove_tag_from_section(self, section_id: int, name: str) -> None:
        tag = self.get_tag(name)
        if tag:
            with self.conn:
                self.conn.execute(
                    "DELETE FROM section_tags WHERE section_id = ? AND tag_id = ?", (section_id, tag["id"])
                )

    def tag_document(self, document_id: int, name: str) -> int:
        """Apply a tag to every active section of a document."""
        tag_id = self._ensure_tag(name)
        with self.conn:
            self.conn.execute(
                "INSERT OR IGNORE INTO section_tags (section_id, tag_id, source)"
                " SELECT id, ?, 'manual' FROM sections WHERE document_id = ? AND status = 'active'",
                (tag_id, document_id),
            )
        return tag_id

    def untag_document(self, document_id: int, name: str) -> int:
        tag = self.get_tag(name)
        if not tag:
            return 0
        with self.conn:
            cur = self.conn.execute(
                "DELETE FROM section_tags WHERE tag_id = ?"
                " AND section_id IN (SELECT id FROM sections WHERE document_id = ?)",
                (tag["id"], document_id),
            )
        return cur.rowcount

    # -- auto-tagging ------------------------------------------------------
    @staticmethod
    def tag_regex(tag: dict) -> str:
        """The regex used to auto-apply ``tag``."""
        if tag.get("is_regex"):
            return tag["pattern"]
        return keyword_to_regex(tag.get("pattern") or tag["name"])

    def preview_matches(self, pattern: str, is_regex: bool, limit: int = 50) -> tuple[int, list[dict]]:
        """Count active sections matching a pattern; return a sample of them."""
        regex = pattern if is_regex else keyword_to_regex(pattern)
        if is_regex:
            validate_regex(regex)
        where = f"s.status = 'active' AND regexp(?, {_MATCH_TEXT})"
        count = self.conn.execute(f"SELECT COUNT(*) FROM sections s WHERE {where}", (regex,)).fetchone()[0]
        rows = self.conn.execute(
            f"SELECT s.id, s.heading, s.content_text, d.title AS document_title"
            f" FROM sections s JOIN documents d ON d.id = s.document_id"
            f" WHERE {where} ORDER BY s.document_id, s.sequence_order LIMIT ?",
            (regex, limit),
        ).fetchall()
        return count, [dict(r) for r in rows]

    def count_auto_tag_candidates(self, tag_id: int) -> int:
        """Active sections that match a tag's pattern but don't have the tag yet."""
        tag = self.get_tag_by_id(tag_id)
        if not tag:
            return 0
        return self.conn.execute(
            f"""SELECT COUNT(*) FROM sections s
                WHERE s.status = 'active' AND regexp(?, {_MATCH_TEXT})
                  AND NOT EXISTS (SELECT 1 FROM section_tags st
                                  WHERE st.section_id = s.id AND st.tag_id = ?)""",
            (self.tag_regex(tag), tag_id),
        ).fetchone()[0]

    def auto_tag(self, tag_id: int, document_id: int | None = None) -> int:
        """Apply a tag to every active section matching its pattern.

        Returns the number of new tag applications.
        """
        tag = self.get_tag_by_id(tag_id)
        if not tag:
            return 0
        sql = (
            "INSERT OR IGNORE INTO section_tags (section_id, tag_id, source)"
            f" SELECT s.id, ?, 'auto' FROM sections s WHERE s.status = 'active' AND regexp(?, {_MATCH_TEXT})"
        )
        params: list[Any] = [tag_id, self.tag_regex(tag)]
        if document_id is not None:
            sql += " AND s.document_id = ?"
            params.append(document_id)
        with self.conn:
            cur = self.conn.execute(sql, params)
        return max(cur.rowcount, 0)

    def run_auto_tagging(self) -> dict[str, int]:
        """Run every tag over the whole knowledge base. Returns {tag name: new applications}."""
        return {tag["name"]: self.auto_tag(tag["id"]) for tag in self.list_tags()}

    def clear_auto_tags(self, tag_id: int) -> int:
        """Remove automatically applied instances of a tag (manual ones stay)."""
        with self.conn:
            cur = self.conn.execute("DELETE FROM section_tags WHERE tag_id = ? AND source = 'auto'", (tag_id,))
        return cur.rowcount

    # -- spaced repetition -------------------------------------------------
    @staticmethod
    def _filter_clause(document_id: int | None, tag_id: int | None) -> tuple[str, list[Any]]:
        clause, params = "", []
        if document_id is not None:
            clause += " AND s.document_id = ?"
            params.append(document_id)
        if tag_id is not None:
            clause += " AND EXISTS (SELECT 1 FROM section_tags st WHERE st.section_id = s.id AND st.tag_id = ?)"
            params.append(tag_id)
        return clause, params

    def get_review_queue(
        self,
        document_id: int | None = None,
        tag_id: int | None = None,
        new_limit: int | None = 20,
        today: date | None = None,
    ) -> list[int]:
        """Section ids to review now: due cards first, then up to ``new_limit`` new cards."""
        today_s = (today or date.today()).isoformat()
        clause, params = self._filter_clause(document_id, tag_id)
        base = "SELECT s.id FROM sections s JOIN reviews r ON r.section_id = s.id WHERE s.status = 'active'"
        due = self.conn.execute(
            f"{base} AND r.status != 'new' AND r.next_review_date <= ? {clause}"
            " ORDER BY r.next_review_date, r.ease_factor, s.document_id, s.sequence_order",
            [today_s, *params],
        ).fetchall()
        new = self.conn.execute(
            f"{base} AND r.status = 'new' {clause} ORDER BY s.document_id, s.sequence_order LIMIT ?",
            [*params, -1 if new_limit is None else new_limit],
        ).fetchall()
        return [r[0] for r in due] + [r[0] for r in new]

    def get_review(self, section_id: int) -> dict | None:
        row = self.conn.execute("SELECT * FROM reviews WHERE section_id = ?", (section_id,)).fetchone()
        return dict(row) if row else None

    def record_answer(self, section_id: int, known: bool, now: datetime | None = None) -> dict | None:
        """Grade a card with SM-2. Returns the *previous* review row so the UI can undo."""
        now = now or datetime.now()
        prev = self.get_review(section_id)
        state = (
            ReviewState(prev["ease_factor"], prev["interval"], prev["repetitions"]) if prev else ReviewState()
        )
        new = sm2(state, QUALITY_KNOWN if known else QUALITY_UNKNOWN)
        next_date = (now.date() + timedelta(days=new.interval)).isoformat()
        with self.conn:
            self.conn.execute(
                """
                INSERT INTO reviews (section_id, ease_factor, interval, repetitions,
                                     next_review_date, last_reviewed_at, status)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(section_id) DO UPDATE SET
                    ease_factor = excluded.ease_factor, interval = excluded.interval,
                    repetitions = excluded.repetitions, next_review_date = excluded.next_review_date,
                    last_reviewed_at = excluded.last_reviewed_at, status = excluded.status
                """,
                (
                    section_id, new.ease_factor, new.interval, new.repetitions, next_date,
                    now.isoformat(timespec="seconds"), review_status(new, reviewed=True),
                ),
            )
        return prev

    def restore_review(self, section_id: int, previous: dict | None) -> None:
        """Undo helper: put a review row back exactly as ``record_answer`` found it."""
        with self.conn:
            if previous is None:
                self.conn.execute("DELETE FROM reviews WHERE section_id = ?", (section_id,))
                return
            self.conn.execute(
                "UPDATE reviews SET ease_factor = ?, interval = ?, repetitions = ?, next_review_date = ?,"
                " last_reviewed_at = ?, status = ? WHERE section_id = ?",
                (
                    previous["ease_factor"], previous["interval"], previous["repetitions"],
                    previous["next_review_date"], previous["last_reviewed_at"], previous["status"], section_id,
                ),
            )

    def stats(self, document_id: int | None = None, tag_id: int | None = None) -> dict[str, int]:
        clause, params = self._filter_clause(document_id, tag_id)
        row = self.conn.execute(
            f"""
            SELECT COUNT(*) AS total,
                   SUM(r.status = 'new')       AS new,
                   SUM(r.status = 'learning')  AS learning,
                   SUM(r.status = 'mastered')  AS mastered,
                   SUM(r.status != 'new' AND r.next_review_date <= ?) AS due
            FROM sections s JOIN reviews r ON r.section_id = s.id
            WHERE s.status = 'active' {clause}
            """,
            [date.today().isoformat(), *params],
        ).fetchone()
        return {k: (row[k] or 0) for k in ("total", "new", "learning", "mastered", "due")}

    def review_statuses(self) -> dict[int, str]:
        """{section_id: 'new' | 'learning' | 'mastered'} for every section."""
        return {r[0]: r[1] for r in self.conn.execute("SELECT section_id, status FROM reviews")}

    def next_due_date(self) -> str | None:
        row = self.conn.execute(
            "SELECT MIN(r.next_review_date) FROM reviews r JOIN sections s ON s.id = r.section_id"
            " WHERE s.status = 'active' AND r.status != 'new' AND r.next_review_date > ?",
            (date.today().isoformat(),),
        ).fetchone()
        return row[0] if row else None

    # -- export ------------------------------------------------------------
    def get_export_groups(
        self,
        section_ids: Iterable[int] | None = None,
        document_ids: Iterable[int] | None = None,
        review_filter: str = "all",
        tag_id: int | None = None,
    ) -> list[ExportGroup]:
        """Active sections grouped by document, in reading order.

        ``review_filter``: ``"all"``; ``"known"`` (reviewed, and the last answer
        was "I know"); or ``"mastered"`` (interval >= 21 days).
        """
        if review_filter not in ("all", "known", "mastered"):
            raise ValueError(f"Unknown review filter: {review_filter}")
        sql = (
            "SELECT s.*, r.repetitions, r.status AS review_status, r.last_reviewed_at"
            " FROM sections s LEFT JOIN reviews r ON r.section_id = s.id"
            " WHERE s.status = 'active'"
        )
        params: list[Any] = []
        if review_filter == "known":
            sql += " AND r.last_reviewed_at IS NOT NULL AND r.repetitions > 0"
        elif review_filter == "mastered":
            sql += " AND r.status = 'mastered'"
        clause, extra = self._filter_clause(None, tag_id)
        sql += clause + " ORDER BY s.document_id, s.sequence_order"
        params += extra

        wanted_sections = set(section_ids) if section_ids is not None else None
        wanted_docs = set(document_ids) if document_ids is not None else None

        tag_map: dict[int, list[str]] = {}
        for sid, name in self.conn.execute(
            "SELECT st.section_id, t.name FROM section_tags st JOIN tags t ON t.id = st.tag_id"
            " ORDER BY t.name COLLATE NOCASE"
        ):
            tag_map.setdefault(sid, []).append(name)

        docs = {d["id"]: d for d in self.list_documents()}
        groups: dict[int, ExportGroup] = {}
        for row in self.conn.execute(sql, params):
            if wanted_sections is not None and row["id"] not in wanted_sections:
                continue
            if wanted_docs is not None and row["document_id"] not in wanted_docs:
                continue
            images = []
            for rel in json.loads(row["image_paths"] or "[]"):
                p = self.data_dir / rel
                if p.is_file():
                    images.append(p)
            group = groups.get(row["document_id"])
            if group is None:
                group = groups[row["document_id"]] = ExportGroup(docs[row["document_id"]], [])
            group.sections.append(
                ExportSection(row["id"], row["heading"] or "", row["content_text"], images, tag_map.get(row["id"], []))
            )
        # Keep library order (by title) for the compiled output.
        order = list(docs)
        return sorted(groups.values(), key=lambda g: order.index(g.document["id"]))
