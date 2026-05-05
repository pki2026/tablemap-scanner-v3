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


def parse_tokens(text: str) -> list[dict]:
    tokens: list[dict] = []
    dollar_re = re.compile(r"^(.+?)\s+(\$.+)$")

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        if ":" in line:
            label, _, rest = line.partition(":")
            label = label.strip()
            value = rest.strip()
            if label and value:
                tokens.append(
                    {
                        "label": label,
                        "value": value,
                        "bbox": dict(TOKEN_BBOX),
                        "confidence": TOKEN_CONFIDENCE,
                    }
                )
            continue

        m = dollar_re.match(line)
        if m:
            label, value = m.group(1).strip(), m.group(2).strip()
            if label and value:
                tokens.append(
                    {
                        "label": label,
                        "value": value,
                        "bbox": dict(TOKEN_BBOX),
                        "confidence": TOKEN_CONFIDENCE,
                    }
                )

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
            tokens = parse_tokens(current)
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
