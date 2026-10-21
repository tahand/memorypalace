"""
ocr.py - Turn photos into cards: text recognition (OCR) and photo clean-up.

OCR engines, picked in this order (both run entirely offline):
* Windows 10/11: the built-in ``Windows.Media.Ocr`` engine through the
  ``winrt-Windows.Media.Ocr`` packages. It uses the OCR language packs of
  the user's Windows display languages.
* Tesseract (any platform) if ``tesseract`` is on PATH or in its default
  Windows install folder.

``extract_text`` is blocking and safe to call from a worker thread. It
returns paragraphs separated by blank lines, so the result can go straight
into ``parser.split_text``.
"""
from __future__ import annotations

import asyncio
import io
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
from pathlib import Path

from parser import clean_text

PHOTO_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff")
CARD_IMAGE_MAX_SIDE = 2000  # px; photos stored on cards are downscaled to this


class OcrError(Exception):
    """Text could not be read from an image (no engine, unreadable file...)."""


# ---------------------------------------------------------------------------
# Photos
# ---------------------------------------------------------------------------
def load_photo(path: str | Path):
    """Open an image upright (phone photos store rotation in EXIF) as RGB."""
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise OcrError("The 'Pillow' package is needed to read photos.\nInstall it with:  pip install Pillow") from exc
    try:
        with Image.open(path) as im:
            im.load()
            im = ImageOps.exif_transpose(im)
            if im.mode in ("RGBA", "LA", "P"):
                background = Image.new("RGB", im.size, "white")
                rgba = im.convert("RGBA")
                background.paste(rgba, mask=rgba.getchannel("A"))
                return background
            return im.convert("RGB")
    except OcrError:
        raise
    except Exception as exc:
        raise OcrError(f"Cannot open {Path(path).name}: {exc}") from exc


def photo_for_card(path: str | Path, max_side: int = CARD_IMAGE_MAX_SIDE) -> tuple[bytes, str]:
    """The photo as (bytes, extension), upright and no bigger than ``max_side``.

    Screenshots and scans stay lossless PNG; camera photos become JPEG.
    """
    im = load_photo(path)
    im.thumbnail((max_side, max_side))
    out = io.BytesIO()
    if Path(path).suffix.lower() in (".png", ".bmp", ".gif", ".tif", ".tiff"):
        im.save(out, "PNG", optimize=True)
        return out.getvalue(), "png"
    im.save(out, "JPEG", quality=88)
    return out.getvalue(), "jpg"


# ---------------------------------------------------------------------------
# Text layout helpers
# ---------------------------------------------------------------------------
def join_lines(lines: list[str]) -> str:
    """Join visual lines into one paragraph, removing end-of-line hyphenation."""
    out = ""
    for line in (ln.strip() for ln in lines):
        if not line:
            continue
        if out.endswith("-") and line[:1].islower():
            out = out[:-1] + line
        else:
            out = f"{out} {line}" if out else line
    return clean_text(out)


def lines_to_paragraphs(lines: list[tuple[str, float, float]]) -> str:
    """``(text, top, height)`` per line -> paragraphs separated by blank lines.

    A vertical gap clearly larger than normal line spacing starts a new
    paragraph, so headings and paragraphs on a page stay apart.
    """
    if not lines:
        return ""
    lines = sorted(lines, key=lambda ln: ln[1])
    typical = statistics.median(h for _, _, h in lines) or 1
    paragraphs: list[list[str]] = [[lines[0][0]]]
    prev_bottom = lines[0][1] + lines[0][2]
    for text, top, height in lines[1:]:
        if top - prev_bottom > typical * 0.75:
            paragraphs.append([])
        paragraphs[-1].append(text)
        prev_bottom = top + height
    return "\n\n".join(p for p in (join_lines(par) for par in paragraphs) if p)


# ---------------------------------------------------------------------------
# Engines
# ---------------------------------------------------------------------------
def _windows_ocr_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        from winrt.windows.media.ocr import OcrEngine
        return OcrEngine.try_create_from_user_profile_languages() is not None
    except Exception:  # noqa: BLE001 - package missing or WinRT unavailable
        return False


def _windows_ocr(path: Path) -> str:
    from winrt.windows.graphics.imaging import BitmapPixelFormat, SoftwareBitmap
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter

    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise OcrError("Windows has no OCR language installed. Add one under Settings › Time & language › "
                       "Language & region (install the language with its optional features).")
    im = load_photo(path)
    limit = OcrEngine.max_image_dimension
    if max(im.size) > limit:
        im.thumbnail((limit, limit))
    r, g, b = im.split()
    from PIL import Image
    bgra = Image.merge("RGBA", (b, g, r, Image.new("L", im.size, 255))).tobytes()
    writer = DataWriter()
    writer.write_bytes(bgra)
    bitmap = SoftwareBitmap.create_copy_from_buffer(writer.detach_buffer(), BitmapPixelFormat.BGRA8,
                                                    im.width, im.height)

    async def recognize():
        return await engine.recognize_async(bitmap)

    result = asyncio.run(recognize())
    lines = []
    for line in result.lines:
        rects = [w.bounding_rect for w in line.words]
        if not rects:
            continue
        top = min(r.y for r in rects)
        bottom = max(r.y + r.height for r in rects)
        lines.append((line.text, top, bottom - top))
    return lines_to_paragraphs(lines)


def _find_tesseract() -> str | None:
    exe = shutil.which("tesseract")
    if exe:
        return exe
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("LOCALAPPDATA", "") + r"\Programs"):
            candidate = Path(base) / "Tesseract-OCR" / "tesseract.exe"
            if candidate.is_file():
                return str(candidate)
    return None


def _tesseract_ocr(path: Path, exe: str) -> str:
    im = load_photo(path)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "page.png"
        im.save(src, "PNG")
        flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        try:
            proc = subprocess.run([exe, str(src), "stdout"], capture_output=True, timeout=180,
                                  creationflags=flags)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise OcrError(f"Tesseract failed on {path.name}: {exc}") from exc
    if proc.returncode != 0:
        raise OcrError(f"Tesseract failed on {path.name}: {proc.stderr.decode(errors='replace').strip()}")
    text = proc.stdout.decode("utf-8", errors="replace")
    paragraphs = [join_lines(block.splitlines()) for block in text.split("\n\n")]
    return "\n\n".join(p for p in paragraphs if p)


def engine_name() -> str | None:
    """Name of the OCR engine that ``extract_text`` will use, or None."""
    if _windows_ocr_available():
        return "Windows OCR"
    if _find_tesseract():
        return "Tesseract"
    return None


def unavailable_message() -> str:
    if sys.platform == "win32":
        return ("No OCR engine is available.\n\nInstall the Windows OCR bindings with:\n"
                "  pip install -r requirements.txt\n\nor install Tesseract OCR "
                "(github.com/UB-Mannheim/tesseract/wiki).")
    return "No OCR engine is available. Install Tesseract OCR (e.g. 'brew install tesseract' or 'apt install tesseract-ocr')."


def extract_text(path: str | Path) -> str:
    """Recognize the text in a photo. Raises ``OcrError``."""
    path = Path(path)
    if not path.is_file():
        raise OcrError(f"File not found: {path}")
    if _windows_ocr_available():
        try:
            return _windows_ocr(path)
        except OcrError:
            raise
        except Exception as exc:  # noqa: BLE001 - fall through to Tesseract if present
            if not _find_tesseract():
                raise OcrError(f"Windows OCR failed on {path.name}: {type(exc).__name__}: {exc}") from exc
    exe = _find_tesseract()
    if exe:
        return _tesseract_ocr(path, exe)
    raise OcrError(unavailable_message())
