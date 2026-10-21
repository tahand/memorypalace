"""
speech.py - Offline text-to-speech for reading cards aloud.

Backends, picked in this order:
* Windows: the built-in SAPI voices through pywin32 (``SAPI.SpVoice``),
  driven from a background thread so the Tk main thread never blocks.
* macOS: the ``say`` command.
* Linux: ``espeak-ng`` or ``espeak``.

Everything runs locally; no text leaves the machine. If no backend is found,
``Speaker.available`` is False and ``speak`` does nothing.

Rate is an integer from -10 (slow) to 10 (fast); 0 is the voice's normal speed.
"""
from __future__ import annotations

import queue
import re
import shutil
import subprocess
import sys
import threading

# SAPI SpeechVoiceSpeakFlags
_SVSF_ASYNC = 1
_SVSF_PURGE = 2
_SVSF_NOT_XML = 16


def speakable(heading: str | None, text: str | None) -> str:
    """Card heading and text as one string, tidied for a speech engine."""
    parts = []
    if heading:
        # Breadcrumbs ("Ch. 2 › Cells") read better as separate phrases.
        parts.append(re.sub(r"\s*[›>]\s*", ". ", heading.strip()))
    if text:
        parts.append(text.strip())
    out = ".\n\n".join(p.rstrip(".") for p in parts if p)
    out = re.sub(r"[#*_`~|]+", " ", out)          # Markdown leftovers
    out = re.sub(r"https?://\S+", "link", out)
    out = re.sub(r"[ \t]+", " ", out)
    return re.sub(r" *\n *", "\n", out).strip()


class _SapiBackend:
    """SAPI on its own thread.

    Even "asynchronous" SAPI calls block for up to a second while they purge
    speech that is playing, so every call goes through a command queue and
    the UI thread never waits. Commands that pile up are coalesced: only the
    latest speak/stop runs.
    """

    def __init__(self):
        self._commands: queue.Queue[tuple[str, object]] = queue.Queue()
        self._ready = threading.Event()
        self._voices: list[str] = []
        self._error: BaseException | None = None
        self._requested = 0     # id of the latest speak request (UI side)
        self._speaking = False  # what the UI sees; set at once, cleared when audio ends
        self._lock = threading.Lock()
        threading.Thread(target=self._run, name="sapi", daemon=True).start()
        if not self._ready.wait(10):
            raise RuntimeError("The Windows speech engine did not start")
        if self._error:
            raise self._error

    # -- UI thread ---------------------------------------------------------
    def voices(self) -> list[str]:
        return list(self._voices)

    def configure(self, voice: str | None, rate: int, volume: int = 100) -> None:
        self._commands.put(("config", (voice, max(-10, min(10, rate)), max(0, min(100, volume)))))

    def speak(self, text: str) -> None:
        with self._lock:
            self._requested += 1
            self._speaking = True
            self._commands.put(("speak", (self._requested, text)))

    def stop(self) -> None:
        with self._lock:
            self._requested += 1
            self._speaking = False
            self._commands.put(("stop", None))

    def is_speaking(self) -> bool:
        return self._speaking

    # -- worker thread -----------------------------------------------------
    def _run(self):
        try:
            import pythoncom
            import win32com.client  # pywin32
            pythoncom.CoInitialize()
            voice = win32com.client.Dispatch("SAPI.SpVoice")
            tokens = {t.GetDescription(): t for t in voice.GetVoices()}
            self._voices = list(tokens)
        except BaseException as exc:  # noqa: BLE001 - reported by __init__
            self._error = exc
            self._ready.set()
            return
        self._ready.set()
        playing = 0  # request id of the speech SAPI is working on, 0 = none
        while True:
            try:
                batch = [self._commands.get(timeout=0.1)]
            except queue.Empty:
                pythoncom.PumpWaitingMessages()
                # Status.RunningState stays 0 for up to ~1.5 s while the audio
                # device starts, so ask whether the queue has finished instead.
                if playing and voice.WaitUntilDone(0):
                    with self._lock:
                        if self._requested == playing:
                            self._speaking = False
                    playing = 0
                continue
            while True:
                try:
                    batch.append(self._commands.get_nowait())
                except queue.Empty:
                    break
            last = next((c for c in reversed(batch) if c[0] in ("speak", "stop")), None)
            try:
                for kind, arg in batch:
                    if kind == "config":
                        name, rate, volume = arg
                        if name in tokens:
                            voice.Voice = tokens[name]
                        voice.Rate = rate
                        voice.Volume = volume
                if last and last[0] == "speak":
                    playing, text = last[1]
                    voice.Speak(text, _SVSF_ASYNC | _SVSF_PURGE | _SVSF_NOT_XML)
                elif last:
                    playing = 0
                    voice.Speak("", _SVSF_ASYNC | _SVSF_PURGE)
            except Exception:  # noqa: BLE001 - keep the speech thread alive
                with self._lock:
                    self._speaking = False
                playing = 0


class _CommandBackend:
    """Runs a speech command in a child process; stopping kills it."""

    def __init__(self, exe: str, kind: str):
        self.exe, self.kind = exe, kind
        self.voice: str | None = None
        self.rate = 0
        self.proc: subprocess.Popen | None = None

    def voices(self) -> list[str]:
        try:
            if self.kind == "say":
                out = subprocess.run([self.exe, "-v", "?"], capture_output=True, text=True, timeout=5).stdout
                return sorted({line.split("  ")[0].strip() for line in out.splitlines() if line.strip()})
            out = subprocess.run([self.exe, "--voices"], capture_output=True, text=True, timeout=5).stdout
            return [line.split()[1] for line in out.splitlines()[1:] if len(line.split()) > 1]  # language codes
        except (OSError, subprocess.SubprocessError, IndexError):
            return []

    def configure(self, voice: str | None, rate: int, volume: int = 100) -> None:
        self.voice, self.rate = voice or None, max(-10, min(10, rate))  # volume: system mixer

    def speak(self, text: str) -> None:
        self.stop()
        wpm = str(int(175 * 1.07 ** self.rate))
        args = [self.exe, "-r" if self.kind == "say" else "-s", wpm]
        if self.voice:
            args += ["-v", self.voice]
        try:
            self.proc = subprocess.Popen(args, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
                                         stderr=subprocess.DEVNULL, text=True)
            self.proc.stdin.write(text)  # type: ignore[union-attr]
            self.proc.stdin.close()      # type: ignore[union-attr]
        except OSError:
            self.proc = None

    def stop(self) -> None:
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
        self.proc = None

    def is_speaking(self) -> bool:
        return bool(self.proc and self.proc.poll() is None)


class Speaker:
    """Facade over the platform backend. Call it from one thread (the Tk main thread)."""

    def __init__(self):
        self._backend = None
        self.error = ""
        try:
            if sys.platform == "win32":
                self._backend = _SapiBackend()
            elif sys.platform == "darwin" and shutil.which("say"):
                self._backend = _CommandBackend(shutil.which("say"), "say")
            else:
                exe = shutil.which("espeak-ng") or shutil.which("espeak")
                if exe:
                    self._backend = _CommandBackend(exe, "espeak")
        except Exception as exc:  # noqa: BLE001 - missing pywin32, broken SAPI install...
            self.error = f"{type(exc).__name__}: {exc}"
        if self._backend is None and not self.error:
            self.error = ("No speech engine found. Install espeak-ng." if sys.platform.startswith("linux")
                          else "No speech engine found.")

    @property
    def available(self) -> bool:
        return self._backend is not None

    def voices(self) -> list[str]:
        return self._backend.voices() if self._backend else []

    def configure(self, voice: str | None = None, rate: int = 0, volume: int = 100) -> None:
        if self._backend:
            self._backend.configure(voice, rate, volume)

    def speak(self, text: str) -> bool:
        """Start reading ``text`` (interrupting anything already playing)."""
        if not self._backend or not text.strip():
            return False
        try:
            self._backend.speak(text)
            return True
        except Exception as exc:  # noqa: BLE001
            self.error = f"{type(exc).__name__}: {exc}"
            return False

    def stop(self) -> None:
        if self._backend:
            try:
                self._backend.stop()
            except Exception:  # noqa: BLE001
                pass

    def is_speaking(self) -> bool:
        try:
            return bool(self._backend and self._backend.is_speaking())
        except Exception:  # noqa: BLE001
            return False
