"""
Tablemap Scanner V3 — Clipboard-only pipeline via Windows Snipping Tool „Text aus Bild kopieren“.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import traceback
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

import win32clipboard
import win32con


POLL_INTERVAL_MS = 500
MIN_POKERTH_CLIPBOARD_CHARS = 50
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
_REGION_CATALOG_INDEX = {r: i for i, r in enumerate(REGION_CATALOG)}


_C0_BUCKET_ORDER = {
    "c0pot_total": 0,
    "c0pot_bets": 1,
    "c0smallblind": 2,
    "c0bigblind": 3,
    "c0pot0": 4,
}


_PLAYER_SUFFIX_ORDER = {
    "name": 0,
    "balance": 1,
    "bet": 2,
    "dealer": 3,
    "action": 4,
    "cards": 5,
}


_HERO_BUTTON_SIZING_ORDER = {
    "betsize_hero": 0,
    "raisesize_hero": 1,
    "callsize_hero": 2,
}


_OWN_RESULTS_LINE_PATTERN = re.compile(
    r"aktuell.{0,300}?zuerst\s*:\s*Scan\b.{0,300}?zuletzt\s*:\s*Scan\b.{0,300}?Treffer",
    re.IGNORECASE | re.DOTALL,
)


_OWN_MARKER_STRONG = ("→ aktuell:",)


_OWN_MARKER_WEAK_CI = frozenset(
    {
        "-> aktuell:",
        "zuerst: scan",
        "zuletzt: scan",
        "treffer:",
        "gesamt-stats-historie",
        "detail-historie pro region",
        "aggregierte regionen",
        "__text__-tokens letzter scan",
        "scan-historie (kurz)",
        "rohdaten letzter scan",
        "tablemap scanner v3 - ergebnisse",
    }
)


def region_sort_key(region_name: str) -> tuple:
    """Priorität: Meta → Pot/Blinds → Board → Hero-Hole → Spieler nach Sitz → Buttons → Sizes → Catalog → Sonstiges."""
    name = region_name

    if name == "game_id":
        return (100, 0, 0, "", name)
    if name == "hand_id":
        return (100, 1, 0, "", name)
    if name == "street":
        return (100, 2, 0, "", name)

    if name.startswith("c0"):
        if name in _C0_BUCKET_ORDER:
            return (200, _C0_BUCKET_ORDER[name], 0, "", name)
        return (260, hash(name) % (10**9), 0, name.lower(), name)

    bm = re.fullmatch(r"board_card_(\d+)", name)
    if bm:
        return (300, int(bm.group(1)), 0, "", name)

    fm = re.fullmatch(r"flop_card_(\d+)", name)
    if fm:
        return (310, int(fm.group(1)), 0, "", name)

    if name == "turn_card":
        return (311, 0, 0, "", name)
    if name == "river_card":
        return (312, 0, 0, "", name)

    hm = re.fullmatch(r"hero_card_(\d+)", name)
    if hm:
        return (400, int(hm.group(1)), 0, "", name)

    pm = re.fullmatch(r"p(\d+)(.+)", name)
    if pm:
        seat = int(pm.group(1))
        suffix = pm.group(2)
        suf_key = _PLAYER_SUFFIX_ORDER.get(suffix, 500)
        return (500, seat, suf_key, suffix, name)

    im = re.fullmatch(r"i(\d+)(\w*)", name)
    if im:
        sid = int(im.group(1))
        suf = im.group(2) or ""
        irank = 0 if suf == "label" else 100 + ord(suf[0]) if suf else 200
        return (600, sid, irank, suf, name)

    if name in _HERO_BUTTON_SIZING_ORDER:
        return (650, _HERO_BUTTON_SIZING_ORDER[name], 0, "", name)

    if name in REGION_CATALOG_SET:
        return (700, _REGION_CATALOG_INDEX[name], 0, "", name)

    return (9000, 0, 0, "", name.lower())


def sorted_region_names(names: Iterable[str]) -> list[str]:
    return sorted(frozenset(names), key=region_sort_key)


def looks_like_own_results_text(clipboard_text: str) -> bool:
    """Erkennt typischen Text aus dem Tk-Ergebnisfenster, um Fehl-Scans zu vermeiden."""
    if not clipboard_text.strip():
        return False

    low = clipboard_text.lower()
    weak_hits = sum(1 for m in _OWN_MARKER_WEAK_CI if m in low)
    arrow_aktuell = any(s in clipboard_text for s in _OWN_MARKER_STRONG)
    ascii_aktuell = "-> aktuell:" in low

    if _OWN_RESULTS_LINE_PATTERN.search(clipboard_text):
        return True

    if weak_hits >= 3:
        return True

    for line in clipboard_text.replace("\r\n", "\n").split("\n"):
        ls = line.strip()
        if "→ aktuell:" in ls or "-> aktuell:" in ls.lower():
            zfirst = bool(re.search(r"zuerst:\s*Scan\s+", ls, re.IGNORECASE))
            zlast = bool(re.search(r"zuletzt:\s*Scan\s+", ls, re.IGNORECASE))
            treff = "treffer" in ls.lower()
            if zfirst and zlast and treff:
                return True

    if (arrow_aktuell or ascii_aktuell) and weak_hits >= 1:
        return True

    if weak_hits >= 2 and ("scan" in low and "|" in clipboard_text):
        return True

    return False


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


def _grouping_plaintext(t: dict) -> str:
    """Volltext für Heuristiken (Geldtokens nur als reines $-Literal, ohne __money__-Präfix)."""
    if (t.get("label") or "").strip() == "__money__":
        return _currency_display(t).strip()
    return token_full_text(t).strip()


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


_PLAYER_NAME_ALLOWED = frozenset(
    {"human player"} | {f"player {i}" for i in range(1, 10)},
)

_LOG_LINE_COMMENT_RE = re.compile(
    r"(?ims)^[^\n]*(---\s*flop\s*---|---\s*turn\s*---|---\s*river\s*---|##\s*game\s*:)"
)


def player_region_prefix(name: str) -> str | None:
    low = name.strip().lower()
    if low == "human player":
        return "p0"
    if m := re.fullmatch(r"player\s+([1-9])", low):
        return f"p{m.group(1)}"
    return None


def is_player_name_text(text: str) -> bool:
    return text.strip().lower() in _PLAYER_NAME_ALLOWED


def money_to_number(text: str) -> float:
    s = text.strip()
    if not s.startswith("$"):
        return -1.0
    body = s[1:].strip().replace(" ", "")
    if not body or not re.fullmatch(r"\d[\d]*(?:[.,]\d[\d]*)*", body):
        return -1.0
    dot_cnt = body.count(".")
    comma_cnt = body.count(",")

    if dot_cnt >= 1 and comma_cnt == 0:
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", body):
            return float(body.replace(".", ""))
        try:
            return float(body.replace(",", "."))
        except ValueError:
            return -1.0
    if comma_cnt >= 1 and dot_cnt == 0:
        if re.fullmatch(r"\d{1,3}(?:,\d{3})+", body):
            return float(body.replace(",", ""))
        if re.fullmatch(r"\d+,\d{1,2}$", body):
            return float(body.replace(",", "."))
        return -1.0
    if dot_cnt >= 1 and comma_cnt >= 1 and re.fullmatch(r"\d{1,3}(?:\.\d{3})+,\d{1,2}$", body):
        return float(body.replace(".", "").replace(",", "."))

    collapsed = body.replace(",", "")
    try:
        return float(collapsed)
    except ValueError:
        return -1.0


def is_money_only(text: str) -> bool:
    return bool(re.fullmatch(r"\$\d+(?:[.,]\d+)*", text.strip()))


def is_plausible_blind_amount(text: str) -> bool:
    if not is_money_only(text):
        return False
    n = money_to_number(text)
    return 0 < n < 1000


def _starts_log_section(ft: str) -> bool:
    low = ft.strip().lower()
    if low == "log":
        return True
    if low.startswith("## game"):
        return True
    return bool(re.match(r"^---\s*flop\s*---|^---\s*turn\s*---|^---\s*river\s*---", low))


def is_log_line(text: str) -> bool:
    raw = text.strip()
    if not raw:
        return False
    low = raw.lower()

    if _starts_log_section(raw):
        return True

    if _LOG_LINE_COMMENT_RE.match(raw.strip()):
        return True

    if re.search(r"(?i)\bhas\s+\[", raw):
        return True
    if re.search(r"(?i)\bsits\s+out\b", raw):
        return True

    if re.match(
        r"(?is)^human\s+player\s+(calls|checks|bets|folds|raises|posts|shows|ties|stands|mucks)\b.*",
        raw,
    ):
        return True
    if re.match(r"(?is)^human\s+player\s+wins\b", raw):
        return True

    if re.match(
        r"(?is)^player\s+\d+\s+(calls|checks|bets|folds|raises|posts|shows|ties|stands|mucks)\b.*",
        raw,
    ):
        return True
    if re.match(r"(?is)^player\s+\d+\s+wins\b", raw):
        return True
    return False


def is_pure_button_label_text(t: dict, ft: str) -> bool:
    ft_s = ft.strip()
    low = ft_s.lower().rstrip(".!?")
    if not low or "$" in ft_s:
        return False
    val = (t.get("value") or "").strip()
    if val.startswith("$"):
        return False

    words = ft_s.strip().split()
    if len(words) == 1:
        wl = words[0].lower().rstrip(".!?")
        return wl.replace("-", "") in {w.replace("-", "") for w in ACTION_WORDS}
    if (
        len(words) == 2
        and words[0].lower() == "all"
        and words[1].lower().rstrip(".!?") == "in"
    ):
        return True

    lab = (t.get("label") or "").strip()
    if lab and lab != "__text__" and not val:
        wl = lab.lower().rstrip(".!?")
        if wl.replace("-", "") in {w.replace("-", "") for w in ACTION_WORDS}:
            return True
    return False


def _indices_clear_for_player_map(ft_list: list[str], in_log_zone: list[bool], ia: int, ib: int) -> bool:
    lo, hi = (ia, ib) if ia <= ib else (ib, ia)
    for k in range(lo, hi + 1):
        if k < 0 or k >= len(ft_list):
            return False
        if in_log_zone[k]:
            return False
        if is_log_line(ft_list[k]):
            return False
    return True


def _extract_player_seat(full: str) -> int | None:
    prefix = player_region_prefix(full.strip())
    if prefix is None:
        return None
    return int(prefix.removeprefix("p"))


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
    low = ft.strip().lower()
    if _small_blind_header(low):
        rn, hdr = "c0smallblind", "SMALL BLIND"
    elif _big_blind_header(low):
        rn, hdr = "c0bigblind", "BIG BLIND"
    else:
        return None

    if "$" not in ft:
        return None

    amt = ""
    for m in re.finditer(r"\$\d+(?:[.,]\d+)*", ft):
        cand = m.group(0)
        if is_plausible_blind_amount(cand):
            amt = cand
        else:
            print(f"[V3] skipped implausible blind amount: {cand}", flush=True)

    return (rn, hdr) if not amt else (rn, f"{hdr} {amt}".strip())


def _find_next_plausible_blind_money(
    tokens: list[dict],
    ft_list: list[str],
    start: int,
    consumed: set[int],
    n: int,
    in_log_zone: list[bool],
) -> int | None:
    for j in range(start, n):
        if j in consumed:
            continue
        if in_log_zone[j]:
            continue
        if not _money_value_follows(tokens[j]):
            continue
        cur = ft_list[j].strip()
        if not is_money_only(cur):
            cur = _currency_display(tokens[j]).strip()
            if not is_money_only(cur):
                continue
        if is_plausible_blind_amount(cur):
            return j
        print(f"[V3] skipped implausible blind amount: {cur}", flush=True)
    return None


def _try_resolve_blind(
    tokens: list[dict],
    ft_list: list[str],
    i: int,
    consumed: set[int],
    n: int,
    in_log_zone: list[bool],
) -> tuple[str, str, list[int], int | None] | None:
    """SMALL BLIND / BIG BLIND auch über zwei Zeilen; `$`-Betrag per Vorwärtssuche (nicht nur direkt folgend)."""
    if i in consumed:
        return None
    if in_log_zone[i]:
        return None
    t0 = tokens[i]
    ft0 = token_full_text(t0).strip()
    ft0_low = ft0.lower()
    kind: str | None = None
    end = i

    if i + 1 < n and (i + 1) not in consumed:
        ft1_low = token_full_text(tokens[i + 1]).strip().lower()
        if ft0_low == "small" and ft1_low == "blind":
            kind = "small"
            end = i + 2
        elif ft0_low == "big" and ft1_low == "blind":
            kind = "big"
            end = i + 2

    if kind is None and "$" not in ft0:
        if _small_blind_header(ft0_low):
            kind = "small"
            end = i + 1
        elif _big_blind_header(ft0_low):
            kind = "big"
            end = i + 1

    if kind is None:
        return None

    hdr = "SMALL BLIND" if kind == "small" else "BIG BLIND"
    rn = "c0smallblind" if kind == "small" else "c0bigblind"
    money_j = _find_next_plausible_blind_money(tokens, ft_list, end, consumed, n, in_log_zone)
    header_idxs = list(range(i, end))
    if money_j is not None:
        val = f"{hdr} {ft_list[money_j].strip()}".strip()
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
    ft_list = [_grouping_plaintext(tokens[ii]) for ii in range(n)]

    in_log_now = False
    in_log_zone: list[bool] = []
    for ii in range(n):
        in_log_zone.append(in_log_now)
        if _starts_log_section(ft_list[ii]):
            in_log_now = True

    print("[V3] player/log guard active", flush=True)
    dbg_skip_pl = 0
    dbg_skip_btn = 0

    i = 0
    while i < n:
        if i in consumed:
            i += 1
            continue

        t = tokens[i]
        ft = ft_list[i]
        ft_low = ft.strip().lower()

        if (
            dbg_skip_pl < 30
            and is_log_line(ft)
            and re.match(r"(?is)^(player\s+[1-9]\b|human\s+player)\s+\S", ft)
        ):
            snip = ft[:120] + ("..." if len(ft) > 120 else "")
            print(f"[V3] skipped log line for player mapping: {snip}", flush=True)
            dbg_skip_pl += 1

        paired = False
        if not in_log_zone[i]:
            if is_player_name_text(ft) and not is_log_line(ft):
                prefix = player_region_prefix(ft.strip())
                if prefix:
                    pname_r, pb_r = f"{prefix}name", f"{prefix}balance"
                    if pname_r not in assigned and pb_r not in assigned:
                        for delta in (1, 2, 3):
                            j = i + delta
                            if j >= n or j in consumed:
                                continue
                            mx = ft_list[j].strip()
                            if not is_money_only(mx) or is_log_line(mx):
                                continue
                            if not _indices_clear_for_player_map(ft_list, in_log_zone, i, j):
                                continue
                            seat = _extract_player_seat(ft.strip())
                            if seat is None:
                                continue
                            if emit(pname_r, ft.strip()) and emit(pb_r, mx):
                                print(
                                    f"[V3] player mapped: {prefix}name={ft.strip()}, {prefix}balance={mx}",
                                    flush=True,
                                )
                                if dealer_pending:
                                    ok_dealer = emit(f"p{seat}dealer", "1")
                                    print(
                                        f"[tablemap] Dealer → Sitz {seat} (p{seat}dealer), gesetzt={ok_dealer}",
                                        flush=True,
                                    )
                                    dealer_pending = False
                                consumed.update((i, j))
                                i = j + 1
                                paired = True
                                break

            if not paired and is_money_only(ft) and not is_log_line(ft):
                for delta in (1, 2, 3):
                    j = i + delta
                    if j >= n or j in consumed:
                        continue
                    nx = ft_list[j].strip()
                    if not is_player_name_text(nx) or is_log_line(nx):
                        continue
                    if not _indices_clear_for_player_map(ft_list, in_log_zone, i, j):
                        continue
                    prefix = player_region_prefix(nx)
                    if not prefix:
                        continue
                    pname_r, pb_r = f"{prefix}name", f"{prefix}balance"
                    if pname_r in assigned or pb_r in assigned:
                        continue
                    seat = _extract_player_seat(nx)
                    if seat is None:
                        continue
                    if emit(pname_r, nx) and emit(pb_r, ft):
                        print(
                            f"[V3] player mapped: {prefix}name={nx}, {prefix}balance={ft}",
                            flush=True,
                        )
                        if dealer_pending:
                            ok_dealer = emit(f"p{seat}dealer", "1")
                            print(
                                f"[tablemap] Dealer → Sitz {seat} (p{seat}dealer), gesetzt={ok_dealer}",
                                flush=True,
                            )
                            dealer_pending = False
                        consumed.update((i, j))
                        i = j + 1
                        paired = True
                        break

        if paired:
            continue

        if not in_log_zone[i]:
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
            if not in_log_zone[i]:
                dealer_pending = True
                print("[tablemap] DEALER erkannt; Zuordnung beim nächsten Spieler-Paar (pXdealer).", flush=True)
            consumed.add(i)
            i += 1
            continue

        btn_lab = _button_display_label(t)
        if btn_lab is not None:
            skip_btn = in_log_zone[i] or is_log_line(ft) or not is_pure_button_label_text(t, ft)
            if skip_btn:
                if dbg_skip_btn < 28 and (in_log_zone[i] or is_log_line(ft)):
                    snip = ft[:120] + ("..." if len(ft) > 120 else "")
                    print(f"[V3] skipped log line for button mapping: {snip}", flush=True)
                    dbg_skip_btn += 1
                i += 1
                continue
            emit(f"i{button_idx}label", btn_lab)
            button_idx += 1
            consumed.add(i)
            i += 1
            continue

        br = _try_resolve_blind(tokens, ft_list, i, consumed, n, in_log_zone)
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


def compute_region_summary(
    region_history: dict[str, list[dict]],
    aggregate_by_region: dict[str, dict],
) -> dict[str, dict]:
    """first_seen_scan / last_seen_scan / hit_count / current_value aus Historie."""
    summary: dict[str, dict] = {}
    for name, hist in region_history.items():
        if not hist:
            continue
        scans_idxs = [int(e["scan_index"]) for e in hist]
        cur = aggregate_by_region.get(name, {})
        summary[name] = {
            "current_value": cur.get("value", hist[-1]["value"]),
            "first_seen_scan": min(scans_idxs),
            "last_seen_scan": max(scans_idxs),
            "hit_count": len(hist),
        }
    return summary


def is_valid_pokerth_scan(clipboard_text: str, regions: list[dict]) -> bool:
    """Kein Merge, wenn keine Regionen oder Text offensichtlich zu kurz."""
    if len(regions) == 0:
        return False
    if len(clipboard_text.strip()) < MIN_POKERTH_CLIPBOARD_CHARS:
        return False
    return True


def readonly_copyable_text_finalize(w: tk.Text, app: tk.Misc) -> None:
    """Text ist schreibgeschützt; Mausauswahl, Ctrl+A/C, Rechtsklick-Menü, Kopieren per Hilfsfunktion."""

    def clipboard_copy_selection() -> None:
        try:
            if not w.tag_ranges(tk.SEL):
                return
            txt = w.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            return
        try:
            app.clipboard_clear()
            app.clipboard_append(txt)
            app.update_idletasks()
            app.update()
        except tk.TclError:
            pass

    def on_select_all(_event: tk.Event | None = None) -> str:
        w.focus_set()
        try:
            w.tag_remove(tk.SEL, "1.0", tk.END)
            w.tag_add(tk.SEL, "1.0", tk.END + "-1c")
        except tk.TclError:
            pass
        return "break"

    def on_copy_hotkey(_event: tk.Event | None = None) -> str:
        clipboard_copy_selection()
        return "break"

    def block_editing_keys(ev: tk.Event) -> str | None:
        st = ev.state or 0
        ctrl = bool(st & 0x0004)

        ks = ev.keysym

        if ctrl:
            if ks.lower() in {"v", "x"}:
                return "break"
            return None

        navigation = frozenset(
            {
                "Left",
                "Right",
                "Up",
                "Down",
                "Home",
                "End",
                "Next",
                "Prior",
                "KP_Left",
                "KP_Right",
                "KP_Up",
                "KP_Down",
                "KP_Page_Up",
                "KP_Page_Down",
                "Shift_L",
                "Shift_R",
                "Control_L",
                "Control_R",
                "Alt_L",
                "Alt_R",
                "Caps_Lock",
                "Num_Lock",
            }
        )
        if ks in navigation:
            return None

        block_named = frozenset(
            {
                "Return",
                "KP_Enter",
                "BackSpace",
                "Delete",
                "Tab",
                "ISO_Left_Tab",
                "Insert",
                "space",
                "Escape",
            }
        )
        if ks in block_named:
            return "break"

        if ev.char:
            try:
                if ord(ev.char) >= 32 or ev.char in ("\t", "\n"):
                    return "break"
            except (TypeError, ValueError):
                return "break"

        return None

    w.configure(state="normal", undo=False, exportselection=True)

    w.bind("<Control-a>", on_select_all)
    w.bind("<Control-A>", on_select_all)
    w.bind("<Control-c>", on_copy_hotkey)
    w.bind("<Control-C>", on_copy_hotkey)
    w.bind("<Control-Insert>", on_copy_hotkey)

    popup = tk.Menu(w, tearoff=0)
    popup.add_command(label="Alles auswählen", command=lambda: on_select_all(None))
    popup.add_command(label="Kopieren", command=clipboard_copy_selection)

    def show_context_menu(ev: tk.Event) -> None:
        try:
            popup.tk_popup(ev.x_root, ev.y_root)
        finally:
            try:
                popup.grab_release()
            except tk.TclError:
                pass

    w.bind("<Button-3>", show_context_menu)

    w.bind("<Key>", block_editing_keys, add="+")


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
    root.geometry("600x340")
    root.minsize(520, 300)
    root.attributes("-topmost", True)

    wait_frame = tk.Frame(root, padx=16, pady=16)
    wait_frame.pack(fill=tk.BOTH, expand=True)

    wait_inner = tk.Frame(wait_frame)
    wait_inner.pack(fill=tk.BOTH, expand=True)

    clipboard_text: str | None = None
    polling_active = True
    _processing_results = False

    scans: list[dict] = []
    aggregate_by_region: dict[str, dict] = {}
    region_history: dict[str, list[dict]] = {}
    tokens_latest: list[dict] = []
    stats_latest: dict[str, int] = {"region_count": 0, "remaining_text_lines": 0}
    results_toplevel: tk.Toplevel | None = None
    results_body_frame: tk.Frame | None = None

    def _start_snipping() -> bool:
        nonlocal proc
        try:
            proc = subprocess.Popen(
                ["snippingtool.exe", "/clip", "/file", str(screenshot_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except FileNotFoundError:
            proc = None
            print(
                "snippingtool.exe nicht gefunden (PATH / Windows-Komponente).",
                file=sys.stderr,
            )
            return False

    def save_and_exit() -> None:
        nonlocal session_done, exit_code, proc
        n_scans = len(scans)
        if n_scans == 0:
            try:
                messagebox.showinfo(
                    "Tablemap Scanner V3",
                    "Noch keine Daten verarbeitet.",
                    parent=root,
                )
            except tk.TclError:
                pass
            if not session_done:
                session_done = True
                exit_code = 0
            kill_snipping_process(proc)
            try:
                if results_toplevel is not None and results_toplevel.winfo_exists():
                    results_toplevel.destroy()
            except tk.TclError:
                pass
            root.destroy()
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"pokerth_tablemap_aggregate_{stamp}.json"
        order = sorted_region_names(aggregate_by_region.keys())
        regions_flat = {name: aggregate_by_region[name]["value"] for name in order}
        regions_detail = [
            {
                "region_name": name,
                "value": aggregate_by_region[name]["value"],
                "catalog_match": aggregate_by_region[name]["catalog_match"],
            }
            for name in order
        ]
        region_summary = compute_region_summary(region_history, aggregate_by_region)
        rh_order = sorted_region_names(region_history.keys())
        region_history_ordered = {name: region_history[name] for name in rh_order}
        rs_order = sorted_region_names(region_summary.keys())
        region_summary_ordered = {name: region_summary[name] for name in rs_order}
        payload = {
            "schema": "pokerth_tablemap_v3_aggregate",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "screenshot": str(screenshot_path),
            "scan_count": n_scans,
            "regions": regions_flat,
            "regions_detail": regions_detail,
            "tokens_latest": tokens_latest,
            "region_catalog": list(REGION_CATALOG),
            "region_count_aggregate": len(aggregate_by_region),
            "remaining_text_lines_latest": stats_latest.get("remaining_text_lines", 0),
            "scans": list(scans),
            "region_history": region_history_ordered,
            "region_summary": region_summary_ordered,
        }
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[V3] final export written:", out_path, flush=True)

        if not session_done:
            session_done = True
            exit_code = 0
        kill_snipping_process(proc)
        try:
            if results_toplevel is not None and results_toplevel.winfo_exists():
                results_toplevel.destroy()
        except tk.TclError:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", save_and_exit)

    def refresh_aggregate_results_window() -> None:
        nonlocal results_toplevel, results_body_frame
        print("[V3] show_results: building GUI", flush=True)

        def _close_results_only() -> None:
            try:
                if results_toplevel is not None and results_toplevel.winfo_exists():
                    results_toplevel.destroy()
            except tk.TclError:
                pass

        n_scans = len(scans)
        n_regions = len(aggregate_by_region)
        export_note = "Noch nicht gespeichert — „Speichern und Beenden“ im Steuerfenster."

        if results_toplevel is None or not results_toplevel.winfo_exists():
            results_toplevel = tk.Toplevel(root)
            results_window = results_toplevel
            results_window.title("Tablemap Scanner V3 - Ergebnisse")
            results_window.geometry("860x860")
            results_window.minsize(620, 680)
            results_window.transient(root)

            results_window.protocol("WM_DELETE_WINDOW", _close_results_only)
            results_body_frame = tk.Frame(results_window)
            results_body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        else:
            results_window = results_toplevel
            assert results_body_frame is not None
            for child in results_body_frame.winfo_children():
                child.destroy()

        assert results_body_frame is not None
        body = results_body_frame

        region_summary = compute_region_summary(region_history, aggregate_by_region)
        rng = sorted_region_names(region_summary.keys())
        summary_lines: list[str] = []
        for name in rng:
            s = region_summary[name]
            summary_lines.append(
                f"{name} → aktuell: {s['current_value']} | zuerst: Scan {s['first_seen_scan']} | "
                f"zuletzt: Scan {s['last_seen_scan']} | Treffer: {s['hit_count']}"
            )
        rag = sorted_region_names(aggregate_by_region.keys())
        aggregated_lines = [
            f"{'*' if aggregate_by_region[name]['catalog_match'] else '?'} {name}  →  {aggregate_by_region[name]['value']}"
            for name in rag
        ]
        detail_blocks: list[str] = []
        for name in sorted_region_names(region_history.keys()):
            lines = [f"{name}:"]
            for e in region_history[name]:
                lines.append(
                    f"  Scan {e['scan_index']} @ {e['created_at']} → {e['value']}"
                )
            detail_blocks.append("\n".join(lines))
        text_lines = [
            token_full_text(t)
            for t in tokens_latest
            if (t.get("label") or "").strip() == "__text__"
        ]
        hist_lines: list[str] = []
        for s in scans:
            hist_lines.append(
                f"--- Scan {s['scan_index']} @ {s['created_at']} — "
                f"Tokens={s['token_count']}, Regionen={s['region_count']} — "
                f"{len(s['clipboard_raw'])} Zeichen Rohdaten"
            )
        last_raw = scans[-1]["clipboard_raw"] if scans else ""

        hdr_text = (
            f"Sammel-Scans: {n_scans}   "
            f"Aktuelle Regionen: {n_regions}   "
            f"Letzte Token-Anzahl: {len(tokens_latest)}   "
            f"Freie Textzeilen (__text__, letzter Scan): {stats_latest.get('remaining_text_lines', 0)}\n"
            f"{export_note}"
        )
        tk.Label(body, text="Zusammenfassung / Status", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(0, 2)
        )
        hdr_f = tk.Frame(body)
        hdr_f.pack(fill=tk.X, pady=(0, 8))
        hdr_sb = tk.Scrollbar(hdr_f)
        hdr_txt = tk.Text(
            hdr_f,
            height=4,
            width=96,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            yscrollcommand=hdr_sb.set,
        )
        hdr_sb.config(command=hdr_txt.yview)
        hdr_sb.pack(side=tk.RIGHT, fill=tk.Y)
        hdr_txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
        hdr_txt.insert("1.0", hdr_text)
        readonly_copyable_text_finalize(hdr_txt, root)

        tk.Label(body, text="Aggregierte Regionen", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(10, 2)
        )
        scroll_wrap = tk.Frame(body)
        scroll_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        sb = tk.Scrollbar(scroll_wrap)
        reg_txt = tk.Text(
            scroll_wrap,
            height=11,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=sb.set,
        )
        sb.config(command=reg_txt.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        reg_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        reg_txt.insert(
            "1.0",
            "\n".join(aggregated_lines) if aggregated_lines else "(noch keine Regionen)",
        )
        readonly_copyable_text_finalize(reg_txt, root)

        tk.Label(body, text="Gesamt-Stats-Historie", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(10, 2)
        )
        sumf = tk.Frame(body)
        sumf.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        sum_sb = tk.Scrollbar(sumf)
        sum_txt = tk.Text(
            sumf,
            height=7,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=sum_sb.set,
        )
        sum_sb.config(command=sum_txt.yview)
        sum_sb.pack(side=tk.RIGHT, fill=tk.Y)
        sum_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sum_txt.insert(
            "1.0",
            "\n".join(summary_lines) if summary_lines else "(noch keine Regionen-Historie)",
        )
        readonly_copyable_text_finalize(sum_txt, root)

        tk.Label(body, text="Detail-Historie pro Region", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(8, 2)
        )
        dtf = tk.Frame(body)
        dtf.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        dt_sb = tk.Scrollbar(dtf)
        dt_txt = tk.Text(
            dtf,
            height=9,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=dt_sb.set,
        )
        dt_sb.config(command=dt_txt.yview)
        dt_sb.pack(side=tk.RIGHT, fill=tk.Y)
        dt_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dt_txt.insert(
            "1.0",
            "\n\n".join(detail_blocks) if detail_blocks else "(noch keine Einträge)",
        )
        readonly_copyable_text_finalize(dt_txt, root)

        tk.Label(body, text=f"__text__-Tokens letzter Scan ({len(text_lines)})", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(10, 2)
        )
        tf = tk.Frame(body)
        tf.pack(fill=tk.BOTH, expand=False, pady=(0, 4))
        tsb = tk.Scrollbar(tf)
        ttxt = tk.Text(
            tf,
            height=5,
            width=96,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            yscrollcommand=tsb.set,
        )
        tsb.config(command=ttxt.yview)
        tsb.pack(side=tk.RIGHT, fill=tk.Y)
        ttxt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttxt.insert("1.0", "\n".join(text_lines) if text_lines else "(keine __text__-Tokens)")
        readonly_copyable_text_finalize(ttxt, root)

        tk.Label(body, text="Scan-Historie (Kurz)", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(8, 2)
        )
        hf = tk.Frame(body)
        hf.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        hsb = tk.Scrollbar(hf)
        htxt = tk.Text(
            hf,
            height=8,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=hsb.set,
        )
        hsb.config(command=htxt.yview)
        hsb.pack(side=tk.RIGHT, fill=tk.Y)
        htxt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        htxt.insert(
            "1.0",
            "\n".join(hist_lines) if hist_lines else "(noch keine Scans)",
        )
        readonly_copyable_text_finalize(htxt, root)

        tk.Label(body, text="Rohdaten letzter Scan", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", pady=(8, 2)
        )
        rf = tk.Frame(body)
        rf.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        rsb = tk.Scrollbar(rf)
        rtx = tk.Text(
            rf,
            height=8,
            width=96,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=rsb.set,
        )
        rsb.config(command=rtx.yview)
        rsb.pack(side=tk.RIGHT, fill=tk.Y)
        rtx.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rtx.insert("1.0", last_raw if last_raw.strip() else "(leer)")
        readonly_copyable_text_finalize(rtx, root)

        def copy_everything_to_clipboard() -> None:
            agg_block = (
                "\n".join(
                    f"{'*' if aggregate_by_region[n]['catalog_match'] else '?'} {n}  →  {aggregate_by_region[n]['value']}"
                    for n in sorted_region_names(aggregate_by_region.keys())
                )
                or "(noch keine Regionen)"
            )
            parts = [
                "--- Zusammenfassung ---",
                hdr_text,
                "",
                "--- Aggregierte Regionen ---",
                agg_block,
                "",
                "--- Gesamt-Stats-Historie ---",
                "\n".join(summary_lines) if summary_lines else "(leer)",
                "",
                "--- Detail-Historie pro Region ---",
                "\n\n".join(detail_blocks) if detail_blocks else "(leer)",
                "",
                "--- __text__-Tokens letzter Scan ---",
                "\n".join(text_lines) if text_lines else "(leer)",
                "",
                "--- Scan-Historie ---",
                "\n".join(hist_lines) if hist_lines else "(leer)",
                "",
                "--- Rohdaten letzter Scan ---",
                last_raw if last_raw.strip() else "(leer)",
            ]
            blob = "\n".join(parts)
            root.clipboard_clear()
            root.clipboard_append(blob)
            root.update_idletasks()
            root.update()

        btn_bar = tk.Frame(body)
        btn_bar.pack(fill=tk.X, pady=(10, 4))
        tk.Button(btn_bar, text="Alles kopieren", command=copy_everything_to_clipboard).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        tk.Button(btn_bar, text="Fenster schließen", command=_close_results_only).pack(side=tk.LEFT)
        results_window.update_idletasks()
        results_window.update()
        results_window.deiconify()
        results_window.lift()
        try:
            results_window.focus_force()
        except tk.TclError as exc:
            print("[V3][WARN] focus_force failed:", repr(exc), flush=True)
        if aggregate_by_region:
            print("[V3] region sort applied", flush=True)
        print("[V3] show_results returned / GUI built", flush=True)

    def _clear_wait_inner() -> None:
        for child in wait_inner.winfo_children():
            child.destroy()

    def process_clipboard_and_show() -> None:
        nonlocal clipboard_text, _processing_results, baseline, proc, polling_active
        nonlocal scans, aggregate_by_region, region_history, tokens_latest, stats_latest
        if _processing_results or not clipboard_text:
            print(
                "[V3] process_clipboard_and_show skipped:",
                f"_processing_results={_processing_results}",
                f"session_done={session_done}",
                f"has_clipboard={bool(clipboard_text)}",
                flush=True,
            )
            return
        _processing_results = True
        current = clipboard_text
        try:
            print("[V3] process_clipboard_and_show started", flush=True)
            print("[V3] clipboard length:", len(current), flush=True)
            tokens = parse_clipboard_to_tokens(current)
            regions, stats = group_tokens_into_regions(tokens)
            nt_ok, nr_ok = len(tokens), len(regions)
            print(f"[V3] candidate tokens:{nt_ok} regions:{nr_ok}", flush=True)

            def resume_after_rejected_scan(dialog_title: str, dialog_body: str) -> None:
                nonlocal baseline, clipboard_text, polling_active
                try:
                    messagebox.showinfo(dialog_title, dialog_body, parent=root)
                except tk.TclError:
                    pass
                kill_snipping_process(proc)
                if not _start_snipping():
                    try:
                        messagebox.showwarning(
                            "Tablemap Scanner V3",
                            "snippingtool.exe konnte nicht neu gestartet werden. "
                            "Erneuter Clipboard-Empfang ist ggf. eingeschränkt.",
                            parent=root,
                        )
                    except tk.TclError:
                        pass
                baseline = _normalize_clip(current)
                clipboard_text = None
                polling_active = True
                print("[V3] waiting for next scan", flush=True)
                build_state1()
                root.after(POLL_INTERVAL_MS, on_poll)

            if looks_like_own_results_text(current):
                print("[V3] invalid scan ignored: own results text detected", flush=True)
                resume_after_rejected_scan(
                    "Tablemap Scanner V3",
                    "Der kopierte Text stammt offenbar aus dem Tablemap-Scanner-Ergebnisfenster "
                    "und wird nicht als PokerTH-Scan übernommen.\n\n"
                    "Bitte erneut Text direkt aus dem Snipping Tool kopieren.",
                )
                return

            if not is_valid_pokerth_scan(current, regions):
                print(
                    f"[V3] invalid scan ignored: clipboard length={len(current)}, tokens={nt_ok}, regions={nr_ok}",
                    flush=True,
                )
                resume_after_rejected_scan(
                    "Tablemap Scanner V3",
                    "Kein verwertbarer PokerTH-Text erkannt.\n\n"
                    "Bitte erneut mit dem Snipping Tool kopieren.",
                )
                return

            scan_no = len(scans) + 1
            print("[V3] scan detected", flush=True)
            print(f"[V3] processing scan #{scan_no}", flush=True)
            print(f"[V3] scan #{scan_no} tokens:", len(tokens), flush=True)
            print(f"[V3] scan #{scan_no} regions:", len(regions), flush=True)
            ts = datetime.now().isoformat(timespec="seconds")
            for r in regions:
                aggregate_by_region[r["region_name"]] = {
                    "value": r["value"],
                    "catalog_match": r["catalog_match"],
                }
                region_history.setdefault(r["region_name"], []).append(
                    {
                        "scan_index": scan_no,
                        "created_at": ts,
                        "value": r["value"],
                    }
                )
            print("[V3] aggregate regions:", len(aggregate_by_region), flush=True)
            print(f"[V3] region history updated: {len(regions)}", flush=True)
            print(f"[V3] region summary entries: {len(region_history)}", flush=True)
            tokens_latest = tokens
            stats_latest = dict(stats)
            scans.append(
                {
                    "scan_index": scan_no,
                    "created_at": ts,
                    "clipboard_raw": current,
                    "token_count": len(tokens),
                    "region_count": stats["region_count"],
                }
            )
            print("[V3] json written: (finale Datei erst bei „Speichern und Beenden“)", flush=True)
            kill_snipping_process(proc)
            if not _start_snipping():
                try:
                    messagebox.showwarning(
                        "Tablemap Scanner V3",
                        "snippingtool.exe konnte nicht neu gestartet werden. "
                        "Erneuter Clipboard-Empfang ist ggf. eingeschränkt.",
                        parent=root,
                    )
                except tk.TclError:
                    pass
            print("[V3] calling show_results", flush=True)
            refresh_aggregate_results_window()
            print("[V3] show_results finished (after build)", flush=True)
            baseline = _normalize_clip(current)
            clipboard_text = None
            polling_active = True
            print("[V3] waiting for next scan", flush=True)
            build_state1()
            root.after(POLL_INTERVAL_MS, on_poll)
        except Exception as exc:
            print("[V3][ERROR]", repr(exc), flush=True)
            traceback.print_exc()
            try:
                messagebox.showerror(
                    "Tablemap Scanner V3 — Fehler",
                    "Beim Verarbeiten ist ein Fehler aufgetreten:\n\n"
                    f"{exc}\n\nDetails siehe Konsole (Terminal).",
                    parent=root,
                )
            except tk.TclError as tke:
                print("[V3][ERROR] could not show messagebox:", repr(tke), flush=True)
        finally:
            _processing_results = False

    def discard_pending_scan() -> None:
        nonlocal baseline, clipboard_text, polling_active, proc
        if session_done or not clipboard_text:
            return
        print("[V3] discard pending scan", flush=True)
        baseline = _normalize_clip(clipboard_text)
        clipboard_text = None
        polling_active = True
        kill_snipping_process(proc)
        if not _start_snipping():
            try:
                messagebox.showwarning(
                    "Tablemap Scanner V3",
                    "snippingtool.exe konnte nicht neu gestartet werden.",
                    parent=root,
                )
            except tk.TclError:
                pass
        build_state1()
        root.after(POLL_INTERVAL_MS, on_poll)

    def build_state1() -> None:
        _clear_wait_inner()
        n_done = len(scans)
        tk.Label(wait_inner, text="Tablemap Scanner V3", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.CENTER, pady=(0, 6)
        )
        if n_done == 0:
            tk.Label(
                wait_inner,
                text="Warte auf Text aus dem Snipping Tool ...",
                wraplength=520,
                justify=tk.CENTER,
            ).pack(anchor=tk.CENTER, pady=(0, 4))
        else:
            tk.Label(
                wait_inner,
                text="Warte auf weiteren Text aus dem Snipping Tool ...",
                wraplength=520,
                justify=tk.CENTER,
            ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text=(
                "Bitte im Snipping Tool den Bereich markieren und "
                "„Text aus Bild kopieren“ klicken."
            ),
            wraplength=520,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text=f"Bisherige Scans: {n_done}",
            wraplength=520,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 14))
        tk.Button(
            wait_inner,
            text="Speichern und Beenden",
            command=save_and_exit,
        ).pack(anchor=tk.CENTER)

    def build_state2() -> None:
        _clear_wait_inner()
        n_done = len(scans)
        tk.Label(wait_inner, text="Tablemap Scanner V3", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.CENTER, pady=(0, 6)
        )
        tk.Label(
            wait_inner,
            text="Text aus dem Snipping Tool wurde erkannt.",
            wraplength=520,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text="Bitte bestätigen, um die Daten in den Ergebnisbestand zu übernehmen.",
            wraplength=520,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 4))
        tk.Label(
            wait_inner,
            text=f"Bisherige Scans: {n_done}",
            wraplength=520,
            justify=tk.CENTER,
        ).pack(anchor=tk.CENTER, pady=(0, 8))
        btn_row1 = tk.Frame(wait_inner)
        btn_row1.pack(anchor=tk.CENTER, pady=(4, 2))
        tk.Button(btn_row1, text="Daten verarbeiten", command=process_clipboard_and_show).pack(
            side=tk.LEFT, padx=6
        )
        tk.Button(btn_row1, text="Verwerfen und weiter warten", command=discard_pending_scan).pack(
            side=tk.LEFT, padx=6
        )
        tk.Button(wait_inner, text="Speichern und Beenden", command=save_and_exit).pack(
            anchor=tk.CENTER, pady=(6, 0)
        )

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
            print("[V3] scan detected (new clipboard)", flush=True)
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
