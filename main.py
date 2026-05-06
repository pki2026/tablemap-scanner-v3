"""
Tablemap Scanner V3 — Clipboard-only pipeline via Windows Snipping Tool „Text aus Bild kopieren“.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import tkinter as tk

import win32clipboard
import win32con


POLL_INTERVAL_MS = 500
TOKEN_BBOX = {"x": 0, "y": 0, "w": 100, "h": 20}
TOKEN_CONFIDENCE = 95

ACTION_WORDS = frozenset({"fold", "call", "raise", "check", "bet", "all-in"})

# Namenskatalog aus Tablemap-Scanner V2 / PokerTH-Heuristik (OpenHoldem-orientiert).
REGION_CATALOG: tuple[str, ...] = (
    "game_id",
    "hand_id",
    "c0pot0",
    "c0pot_total",
    "c0pot_bets",
    "c0smallblind",
    "c0bigblind",
    *tuple(f"p{i}name" for i in range(10)),
    *tuple(f"p{i}balance" for i in range(10)),
    *tuple(f"p{i}dealer" for i in range(10)),
    *tuple(f"p{i}bet" for i in range(10)),
    *tuple(f"i{i}label" for i in range(10)),
)
REGION_CATALOG_SET: frozenset[str] = frozenset(REGION_CATALOG)


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


_SOFT_MERGE_LINE = re.compile(r"^[A-Za-z][A-Za-z\s\-']{0,46}$")


def _leading_ws_units(s: str) -> int:
    """Vergleichbare Einrückung (Leerzeichen + Tab als 4)."""
    u = 0
    for ch in s:
        if ch == " ":
            u += 1
        elif ch == "\t":
            u += 4
        else:
            break
    return u


def _merge_lines_for_multiword_labels(lines: list[str]) -> list[str]:
    """Führt Zeilen zusammen, die logisch eine Einheit bilden (SMALL+BLIND, stärkere Einrückung)."""
    out: list[str] = []
    i = 0
    while i < len(lines):
        raw = lines[i]
        stripped = raw.strip()
        low = stripped.lower()
        if i + 1 < len(lines):
            raw_n = lines[i + 1]
            stripped_n = raw_n.strip()
            low_n = stripped_n.lower()
            if stripped and stripped_n and "$" not in stripped and "$" not in stripped_n:
                if ":" not in stripped and ":" not in stripped_n:
                    if low == "small" and low_n == "blind":
                        out.append(f"{stripped} {stripped_n}")
                        i += 2
                        continue
                    if low == "big" and low_n == "blind":
                        out.append(f"{stripped} {stripped_n}")
                        i += 2
                        continue
                    wi = _leading_ws_units(raw)
                    wj = _leading_ws_units(raw_n)
                    if (
                        wj > wi
                        and len(stripped) <= 48
                        and len(stripped_n) <= 48
                        and _SOFT_MERGE_LINE.fullmatch(stripped)
                        and _SOFT_MERGE_LINE.fullmatch(stripped_n)
                    ):
                        out.append(f"{stripped} {stripped_n}")
                        i += 2
                        continue
        out.append(raw.rstrip("\r"))
        i += 1
    return out


def parse_clipboard_to_tokens(text: str) -> list[dict]:
    """Snipping-Tool-Zeilen parsen: Strikt-Modus (Doppelpunkt / Label $x) plus breite Heuristik."""
    tokens: list[dict] = []
    dollar_re = re.compile(r"^(.+?)\s+(\$.+)$")

    merged_lines = _merge_lines_for_multiword_labels(text.split("\n"))

    for raw_line in merged_lines:
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


def token_full_text(t: dict) -> str:
    lab = (t.get("label") or "").strip()
    val = (t.get("value") or "").strip()
    if lab == "__text__":
        return val
    if lab and val:
        return f"{lab} {val}".strip()
    return lab or val


def _currency_display(t: dict) -> str:
    """Reinen Geldbetrag aus Parser-Tokens (__money__ + Betrag)."""
    lab = (t.get("label") or "").strip()
    val = (t.get("value") or "").strip()
    if lab == "__money__":
        return val
    return token_full_text(t).strip()


def _money_value_follows(t: dict) -> bool:
    val = (t.get("value") or "").strip()
    if val.startswith("$"):
        return True
    return token_full_text(t).strip().startswith("$")


def _extract_player_seat(full: str) -> int | None:
    low = full.lower().strip()
    if low.startswith("human player"):
        return 0
    m = re.search(r"player\s+(\d+)", low, re.IGNORECASE)
    if not m:
        return None
    n = int(m.group(1))
    if 1 <= n <= 10:
        return n - 1
    return None


def _is_player_header_token(t: dict) -> bool:
    low = token_full_text(t).strip().lower()
    return low.startswith("player") or low.startswith("human player")


def _small_blind_header(ft_low: str) -> bool:
    return ft_low.startswith("small blind") or ft_low == "small blind"


def _big_blind_header(ft_low: str) -> bool:
    return ft_low.startswith("big blind") or ft_low == "big blind"


def _total_header_pair(ft_low: str) -> bool:
    return ft_low.rstrip(":").strip() == "total"


def _bets_header_pair(ft_low: str) -> bool:
    return ft_low.rstrip(":").strip() == "bets"


def _total_bets_single_token(t: dict) -> tuple[str, str] | None:
    lab = (t.get("label") or "").strip().lower().rstrip(":")
    val = (t.get("value") or "").strip()
    if lab == "total" and "$" in val:
        return ("c0pot_total", val.strip())
    if lab == "bets" and "$" in val:
        return ("c0pot_bets", val.strip())
    ft = token_full_text(t).strip()
    low = ft.lower()
    if low.startswith("total") and "$" in ft:
        return ("c0pot_total", ft.strip())
    if low.startswith("bets") and "$" in ft:
        return ("c0pot_bets", ft.strip())
    return None


def _game_hand_match(t: dict) -> tuple[str, str] | None:
    lab = (t.get("label") or "").strip().lower().rstrip(":")
    val = (t.get("value") or "").strip()
    if lab == "game" and re.fullmatch(r"\d+", val):
        return ("game_id", val)
    if lab == "hand" and re.fullmatch(r"\d+", val):
        return ("hand_id", val)
    ft = token_full_text(t).strip()
    m = re.match(r"(?i)game\s*:\s*(\d+)\s*$", ft)
    if m:
        return ("game_id", m.group(1))
    m = re.match(r"(?i)hand\s*:\s*(\d+)\s*$", ft)
    if m:
        return ("hand_id", m.group(1))
    return None


def _is_dealer_token(t: dict) -> bool:
    ft = token_full_text(t).strip().lower().rstrip(":")
    return ft == "dealer"


def _button_display_label(t: dict) -> str | None:
    lab = (t.get("label") or "").strip()
    if lab and lab != "__text__":
        low_lab = _norm_action_label(lab)
        if low_lab in ACTION_WORDS:
            return lab.strip()
    ft = token_full_text(t).strip()
    low_ft = _norm_action_label(ft)
    if low_ft in ACTION_WORDS:
        return ft.strip()
    return None


def _blind_region_from_single_line(ft: str) -> tuple[str, str] | None:
    low = ft.lower()
    if ("small blind" in low or low.startswith("small blind")) and "$" in ft:
        return ("c0smallblind", ft.strip())
    if ("big blind" in low or low.startswith("big blind")) and "$" in ft:
        return ("c0bigblind", ft.strip())
    return None


def _find_next_unclaimed_money(tokens: list[dict], start: int, consumed: set[int], n: int) -> int | None:
    """Nächster Token-Index mit Geldbetrag, der noch keiner Region zugeordnet ist."""
    for j in range(start, n):
        if j in consumed:
            continue
        if _money_value_follows(tokens[j]):
            return j
    return None


def _try_resolve_blind(tokens: list[dict], i: int, consumed: set[int], n: int) -> tuple[str, str, list[int], int | None] | None:
    """SMALL BLIND / BIG BLIND auch über zwei Zeilen; `$`-Betrag per Vorwärtssuche (nicht nur direkt folgend)."""
    if i in consumed:
        return None
    t0 = tokens[i]
    ft0 = token_full_text(t0).strip()
    ft0_low = ft0.lower()
    kind: str | None = None
    end = i
    hdr: str | None = None

    if i + 1 < n and (i + 1) not in consumed:
        ft1_low = token_full_text(tokens[i + 1]).strip().lower()
        if ft0_low == "small" and ft1_low == "blind":
            kind = "small"
            end = i + 2
            hdr = "SMALL BLIND"
        elif ft0_low == "big" and ft1_low == "blind":
            kind = "big"
            end = i + 2
            hdr = "BIG BLIND"

    if kind is None and "$" not in ft0:
        if _small_blind_header(ft0_low):
            kind = "small"
            end = i + 1
            hdr = ft0.strip()
        elif _big_blind_header(ft0_low):
            kind = "big"
            end = i + 1
            hdr = ft0.strip()

    if kind is None:
        return None

    rn = "c0smallblind" if kind == "small" else "c0bigblind"
    money_j = _find_next_unclaimed_money(tokens, end, consumed, n)
    header_idxs = list(range(i, end))
    if money_j is not None:
        val = f"{hdr} {_currency_display(tokens[money_j])}".strip()
    else:
        val = hdr
    return (rn, val, header_idxs, money_j)


def group_tokens_into_regions(tokens: list[dict]) -> tuple[list[dict], dict[str, int]]:
    """Gruppiert Folgen von Roh-Tokens zu Poker-Regionen (Lesereihenfolge = Snipping Tool)."""
    regions: list[dict] = []
    assigned: set[str] = set()
    consumed: set[int] = set()
    dealer_pending = False
    button_idx = 0

    def emit(region_name: str, value: str) -> bool:
        if region_name in assigned:
            return False
        assigned.add(region_name)
        regions.append(
            {
                "region_name": region_name,
                "value": value,
                "catalog_match": region_name in REGION_CATALOG_SET,
            }
        )
        return True

    n = len(tokens)
    i = 0
    while i < n:
        if i in consumed:
            i += 1
            continue

        t = tokens[i]
        ft = token_full_text(t)
        ft_low = ft.strip().lower()

        single_blind = _blind_region_from_single_line(ft)
        if single_blind:
            rn, val = single_blind
            emit(rn, val)
            consumed.add(i)
            i += 1
            continue

        tb = _total_bets_single_token(t)
        if tb:
            rn, val = tb
            emit(rn, val)
            consumed.add(i)
            i += 1
            continue

        gh = _game_hand_match(t)
        if gh:
            rn, val = gh
            emit(rn, val)
            consumed.add(i)
            i += 1
            continue

        if _is_dealer_token(t):
            dealer_pending = True
            print("[tablemap] DEALER erkannt; Zuordnung beim nächsten Spieler-Paar (pXdealer).", flush=True)
            consumed.add(i)
            i += 1
            continue

        btn_lab = _button_display_label(t)
        if btn_lab is not None:
            emit(f"i{button_idx}label", btn_lab)
            button_idx += 1
            consumed.add(i)
            i += 1
            continue

        if _is_player_header_token(t) and i + 1 < n and _money_value_follows(tokens[i + 1]):
            seat = _extract_player_seat(ft)
            if seat is not None:
                stack_txt = _currency_display(tokens[i + 1])
                emit(f"p{seat}name", ft.strip())
                emit(f"p{seat}balance", stack_txt)
                if dealer_pending:
                    ok_dealer = emit(f"p{seat}dealer", "1")
                    print(
                        f"[tablemap] Dealer → Sitz {seat} (p{seat}dealer), gesetzt={ok_dealer}",
                        flush=True,
                    )
                    dealer_pending = False
                consumed.update((i, i + 1))
                i += 2
                continue

        br = _try_resolve_blind(tokens, i, consumed, n)
        if br:
            rn, val, header_idxs, money_j = br
            ok = emit(rn, val)
            consumed.update(header_idxs)
            if ok and money_j is not None:
                consumed.add(money_j)
            i += 1
            continue

        if i + 1 < n and _money_value_follows(tokens[i + 1]) and "$" not in ft:
            nxt_val = _currency_display(tokens[i + 1])
            if _total_header_pair(ft_low):
                emit("c0pot_total", nxt_val)
                consumed.update((i, i + 1))
                i += 2
                continue
            if _bets_header_pair(ft_low):
                emit("c0pot_bets", nxt_val)
                consumed.update((i, i + 1))
                i += 2
                continue

        i += 1

    remaining_text = sum(1 for x in tokens if (x.get("label") or "") == "__text__")
    stats = {"region_count": len(regions), "remaining_text_lines": remaining_text}
    return regions, stats


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
    root.geometry("520x260")
    root.minsize(480, 220)
    root.attributes("-topmost", True)

    wait_frame = tk.Frame(root, padx=16, pady=16)
    wait_frame.pack(fill=tk.BOTH, expand=True)

    wait_inner = tk.Frame(wait_frame)
    wait_inner.pack(fill=tk.BOTH, expand=True)

    clipboard_text: str | None = None
    polling_active = True

    def _shutdown() -> None:
        nonlocal session_done, exit_code
        if not session_done:
            session_done = True
            exit_code = 0
        kill_snipping_process(proc)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _shutdown)

    def show_results(
        tokens: list[dict],
        regions: list[dict],
        stats: dict[str, int],
        out_path: Path,
        clipboard_raw: str,
    ) -> None:
        wait_frame.pack_forget()
        root.geometry("800x780")
        root.minsize(620, 620)
        body = tk.Frame(root)
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        hdr = tk.Label(
            body,
            text=(
                f"Tokens: {len(tokens)}   "
                f"Regionen: {stats['region_count']}   "
                f"Freie Textzeilen (__text__): {stats['remaining_text_lines']}\n"
                f"Export: {out_path}"
            ),
            justify=tk.LEFT,
            anchor="w",
        )
        hdr.pack(fill=tk.X)

        tk.Label(body, text="Regionen", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 2))
        scroll_wrap = tk.Frame(body)
        scroll_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        sb = tk.Scrollbar(scroll_wrap)
        lb = tk.Listbox(scroll_wrap, height=11, width=96, yscrollcommand=sb.set, font=("Segoe UI", 10))
        sb.config(command=lb.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for r in regions:
            mark = "*" if r["catalog_match"] else "?"
            lb.insert(tk.END, f"{mark} {r['region_name']}  →  {r['value']}")

        text_lines = [
            token_full_text(t)
            for t in tokens
            if (t.get("label") or "").strip() == "__text__"
        ]
        tk.Label(body, text=f"__text__-Tokens ({len(text_lines)})", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(10, 2)
        )
        tf = tk.Frame(body)
        tf.pack(fill=tk.BOTH, expand=False, pady=(0, 4))
        tsb = tk.Scrollbar(tf)
        ttxt = tk.Text(
            tf,
            height=6,
            width=96,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            yscrollcommand=tsb.set,
            state=tk.DISABLED,
        )
        tsb.config(command=ttxt.yview)
        tsb.pack(side=tk.RIGHT, fill=tk.Y)
        ttxt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttxt.config(state=tk.NORMAL)
        ttxt.insert("1.0", "\n".join(text_lines) if text_lines else "(keine __text__-Tokens)")
        ttxt.config(state=tk.DISABLED)

        tk.Label(body, text="Rohdaten aus Snipping Tool", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(8, 2)
        )
        rf = tk.Frame(body)
        rf.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        rsb = tk.Scrollbar(rf)
        rtx = tk.Text(
            rf,
            height=14,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=rsb.set,
            state=tk.DISABLED,
        )
        rsb.config(command=rtx.yview)
        rsb.pack(side=tk.RIGHT, fill=tk.Y)
        rtx.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rtx.config(state=tk.NORMAL)
        rtx.insert("1.0", clipboard_raw if clipboard_raw.strip() else "(leer)")
        rtx.config(state=tk.DISABLED)

        tk.Button(body, text="Schließen", command=_shutdown).pack(pady=(10, 4))

    def _clear_wait_inner() -> None:
        for child in wait_inner.winfo_children():
            child.destroy()

    def process_clipboard_and_show() -> None:
        nonlocal session_done, clipboard_text
        if session_done or not clipboard_text:
            return
        current = clipboard_text
        tokens = parse_clipboard_to_tokens(current)
        regions, stats = group_tokens_into_regions(tokens)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"pokerth_tablemap_{stamp}.json"
        payload = {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "screenshot": str(screenshot_path),
            "token_count": len(tokens),
            "tokens": tokens,
            "regions": regions,
            "region_catalog": list(REGION_CATALOG),
            "region_count": stats["region_count"],
            "remaining_text_lines": stats["remaining_text_lines"],
            "clipboard_raw": current,
        }
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        session_done = True
        kill_snipping_process(proc)
        show_results(tokens, regions, stats, out_path, current)

    def build_state1() -> None:
        _clear_wait_inner()
        tk.Label(wait_inner, text="Tablemap Scanner V3", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.CENTER, pady=(0, 6)
        )
        tk.Label(
            wait_inner,
            text="Warte auf Text aus dem Snipping Tool ...",
            wraplength=460,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text=(
                "Bitte im Snipping Tool den Bereich markieren und "
                "„Text aus Bild kopieren“ klicken."
            ),
            wraplength=460,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 14))
        tk.Button(
            wait_inner,
            text="Scan abbrechen / Beenden",
            command=_shutdown,
        ).pack(anchor=tk.CENTER)

    def build_state2() -> None:
        _clear_wait_inner()
        tk.Label(wait_inner, text="Tablemap Scanner V3", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.CENTER, pady=(0, 6)
        )
        tk.Label(
            wait_inner,
            text="Text aus dem Snipping Tool wurde erkannt.",
            wraplength=460,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text="Bitte bestätigen, um die Daten zu verarbeiten.",
            wraplength=460,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 14))
        btn_row = tk.Frame(wait_inner)
        btn_row.pack(anchor=tk.CENTER, pady=(4, 0))
        tk.Button(btn_row, text="Daten verarbeiten", command=process_clipboard_and_show).pack(
            side=tk.LEFT, padx=8
        )
        tk.Button(btn_row, text="Abbrechen / Beenden", command=_shutdown).pack(side=tk.LEFT, padx=8)

    build_state1()

    root.update_idletasks()
    root.update()

    baseline_raw = _read_clipboard_text()
    baseline = _normalize_clip(baseline_raw)

    def on_poll() -> None:
        nonlocal clipboard_text, polling_active
        if session_done or not polling_active:
            return

        current_raw = _read_clipboard_text()
        current = _normalize_clip(current_raw)
        if current and current != baseline:
            clipboard_text = current
            polling_active = False
            kill_snipping_process(proc)
            build_state2()
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
