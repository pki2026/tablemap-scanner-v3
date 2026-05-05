"""
Tablemap Scanner V3 — Clipboard-only pipeline via Windows Snipping Tool „Text aus Bild kopieren“.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

import win32clipboard
import win32con


POLL_INTERVAL_MS = 500
TIMEOUT_S = 120
TOKEN_BBOX = {"x": 0, "y": 0, "w": 100, "h": 20}
TOKEN_CONFIDENCE = 95

ACTION_WORDS = frozenset({"fold", "call", "raise", "check", "bet", "all-in"})


def _first_label_boundary(line: str) -> int:
    idx_colon = line.find(":")
    idx_space = next((i for i, ch in enumerate(line) if ch.isspace()), -1)
    candidates = [i for i in (idx_colon, idx_space) if i >= 0]
    return min(candidates) if candidates else -1


def _norm_action_label(s: str) -> str:
    return s.strip().lower()


def _make_token(label: str, value: str) -> dict:
    return {
        "label": label,
        "value": value,
        "bbox": dict(TOKEN_BBOX),
        "confidence": TOKEN_CONFIDENCE,
    }


def _read_clipboard_text() -> str | None:
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return None
            raw = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None
    if raw is None:
        return None
    return str(raw)


def _normalize_clip(s: str | None) -> str:
    if s is None:
        return ""
    return s.replace("\r\n", "\n").replace("\r", "\n")


def parse_clipboard_to_tokens(text: str) -> list[dict]:
    """Snipping-Tool-Zeilen parsen: Strikt-Modus (Doppelpunkt / Label $x) plus breite Heuristik."""
    tokens: list[dict] = []
    dollar_re = re.compile(r"^(.+?)\s+(\$.+)$")

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        # Strikt: „Label: Wert“ (beide Teile nicht leer)
        if ":" in line:
            label_s, _, rest_s = line.partition(":")
            label_s, rest_s = label_s.strip(), rest_s.strip()
            if label_s and rest_s:
                tokens.append(_make_token(label_s, rest_s))
                continue

        # Strikt: „Label $Wert“
        m = dollar_re.match(line)
        if m:
            ls, vs = m.group(1).strip(), m.group(2).strip()
            if ls and vs:
                tokens.append(_make_token(ls, vs))
                continue

        # Heuristik: erstes Wort bis zum ersten ':' oder Whitespace, Rest = value
        boundary = _first_label_boundary(line)
        if boundary >= 0:
            label_h = line[:boundary].strip()
            value_h = line[boundary + 1 :].lstrip()
        else:
            label_h = line
            value_h = ""

        whole_norm = _norm_action_label(line)

        # Numerisch / Betrag: '$' im Wertteil (oder eingeklebtes „BB$2“ ohne Leerzeichen)
        if "$" in value_h:
            tokens.append(_make_token(label_h or "__money__", value_h))
            continue
        if "$" in label_h and not value_h:
            pre, _, post = label_h.partition("$")
            lb = pre.strip() or "__money__"
            vs = post.strip()
            tokens.append(_make_token(lb, f"${vs}" if vs and not vs.startswith("$") else vs))
            continue

        # Aktions-Buttons (auch ohne Dollar)
        if _norm_action_label(label_h) in ACTION_WORDS:
            tokens.append(_make_token(label_h, value_h))
            continue
        if boundary < 0 and whole_norm in ACTION_WORDS:
            tokens.append(_make_token(line.strip(), ""))
            continue

        # Übrige Zeilen als Freitext
        tokens.append(_make_token("__text__", line))

    return tokens


def kill_snipping_process(proc: subprocess.Popen | None) -> None:
    if proc is None:
        return
    pid = proc.pid
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Tablemap Scanner V3 (Snipping Tool → Clipboard)")
    parser.add_argument(
        "--screenshot",
        type=Path,
        required=True,
        help="Pfad zum Screenshot-Bild für Snipping Tool (/file)",
    )
    args = parser.parse_args()
    screenshot_path = args.screenshot.resolve()
    if not screenshot_path.is_file():
        print(f"Screenshot nicht gefunden: {screenshot_path}", file=sys.stderr)
        return 2

    out_dir = Path(__file__).resolve().parent / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    proc: subprocess.Popen | None = None
    exit_code = 0
    session_done = False

    try:
        proc = subprocess.Popen(
            ["snippingtool.exe", "/clip", "/file", str(screenshot_path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except FileNotFoundError:
        print(
            "snippingtool.exe nicht gefunden (PATH / Windows-Komponente).",
            file=sys.stderr,
        )
        return 3

    root = tk.Tk()
    root.title("Tablemap Scanner V3")
    root.geometry("520x120")
    root.attributes("-topmost", True)

    status = tk.Label(
        root,
        text=(
            "Bitte Spielfeld im Snipping Tool markieren und "
            "'Text aus Bild kopieren' klicken."
        ),
        wraplength=500,
        justify=tk.CENTER,
        padx=12,
        pady=12,
    )
    status.pack(fill=tk.BOTH, expand=True)
    root.update_idletasks()
    root.update()

    baseline_raw = _read_clipboard_text()
    baseline = _normalize_clip(baseline_raw)
    start = time.monotonic()

    def on_poll() -> None:
        nonlocal session_done, exit_code
        if session_done:
            return

        elapsed = time.monotonic() - start
        if elapsed >= TIMEOUT_S:
            messagebox.showwarning(
                "Timeout",
                f"Keine neue Zwischenablage innerhalb von {TIMEOUT_S} Sekunden.",
                parent=root,
            )
            session_done = True
            exit_code = 1
            kill_snipping_process(proc)
            root.destroy()
            return

        current_raw = _read_clipboard_text()
        current = _normalize_clip(current_raw)
        if current and current != baseline:
            tokens = parse_clipboard_to_tokens(current)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = out_dir / f"pokerth_tablemap_{stamp}.json"
            payload = {
                "generated_at": datetime.now().isoformat(timespec="seconds"),
                "screenshot": str(screenshot_path),
                "token_count": len(tokens),
                "tokens": tokens,
            }
            out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

            status.config(
                text=f"Daten aus Snipping Tool übernommen: {len(tokens)} Token(s)."
            )
            session_done = True
            kill_snipping_process(proc)

            def close_ui() -> None:
                root.destroy()

            root.after(2000, close_ui)
            return

        root.after(POLL_INTERVAL_MS, on_poll)

    root.after(POLL_INTERVAL_MS, on_poll)

    try:
        root.mainloop()
    finally:
        kill_snipping_process(proc)

    return exit_code if session_done else 1


if __name__ == "__main__":
    raise SystemExit(main())
