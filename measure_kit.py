"""
Snipping-Tool-Fenster erfassen (BBox), Anker per Template-Matching, heuristische Textboxen.
"""

from __future__ import annotations

from dataclasses import dataclass
from collections import defaultdict, deque
from datetime import datetime
from pathlib import Path
import math
import re

import cv2
import numpy as np
import win32gui
from PIL import ImageGrab

# Titel-Fragmente (inkl. häufige Lokalisierungen)
_SNIPPING_TITLE_KEYWORDS: tuple[str, ...] = (
    "snipping tool",
    "snipping",
    "screen snipping",
    "snip & sketch",
    "bildschirm",
    "ausschnitt",
)


class MeasureError(Exception):
    """Vermessung nicht möglich (Fenster fehlt, Anker passt nicht, …)."""


def _iou(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> float:
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ix1, iy1 = max(ax1, bx1), max(ay1, by1)
    ix2, iy2 = min(ax2, bx2), min(ay2, by2)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = float((ix2 - ix1) * (iy2 - iy1))
    aa = float(max(0, ax2 - ax1) * max(0, ay2 - ay1))
    bb = float(max(0, bx2 - bx1) * max(0, by2 - by1))
    u = aa + bb - inter
    return inter / u if u > 0 else 0.0


def _merge_boxes(
    boxes: list[tuple[int, int, int, int]], *, iou_thresh: float = 0.25
) -> list[tuple[int, int, int, int]]:
    if len(boxes) <= 1:
        return list(boxes)
    by_area = sorted(
        boxes,
        key=lambda b: max(0, b[2] - b[0]) * max(0, b[3] - b[1]),
        reverse=True,
    )
    kept: list[tuple[int, int, int, int]] = []
    for b in by_area:
        if any(_iou(b, k) >= iou_thresh for k in kept):
            continue
        kept.append(b)
    return sorted(kept, key=lambda t: (t[1] // 16, t[0]))


def _detect_snipping_text_boxes(bgr: np.ndarray) -> list[tuple[int, int, int, int]]:
    h0, w0 = bgr.shape[:2]
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blur, 25, 80)
    k = max(3, min(w0, h0) // 180)
    if k % 2 == 0:
        k += 1
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k, k))
    edges = cv2.dilate(edges, kernel, iterations=2)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    raw: list[tuple[int, int, int, int]] = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w < 8 or h < 8:
            continue
        if w * h < 120:
            continue
        if w > int(w0 * 0.94) and h > int(h0 * 0.94):
            continue
        raw.append((x, y, x + w, y + h))
    merged = _merge_boxes(raw)
    merged.sort(key=lambda t: (t[1] // 16, t[0]))
    return merged


@dataclass(frozen=True)
class _HwndPick:
    hwnd: int
    title: str
    left: int
    top: int
    right: int
    bottom: int


def _find_largest_snipping_window() -> _HwndPick | None:
    best: _HwndPick | None = None
    best_area = 0

    def _cb(hwnd: int, _: object) -> None:
        nonlocal best, best_area
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = (win32gui.GetWindowText(hwnd) or "").strip()
        if not title:
            return
        tl = title.lower()
        if not any(k in tl for k in _SNIPPING_TITLE_KEYWORDS):
            return
        try:
            L, T, R, B = win32gui.GetWindowRect(hwnd)
        except win32gui.error:
            return
        w, h = max(0, R - L), max(0, B - T)
        area = w * h
        if area < 8000:
            return
        if area > best_area:
            best_area = area
            best = _HwndPick(hwnd=hwnd, title=title, left=L, top=T, right=R, bottom=B)

    try:
        win32gui.EnumWindows(_cb, None)
    except win32gui.error:
        return None
    return best


def capture_and_measure_snipping_tool(repo_root: Path, anchor_patch_path: Path) -> dict:
    """
    Snipping-Fenster per win32-Titel finden, nur dieses Rechteck mit ImageGrab erfassen,
    Anker laden, Boxen heuristisch ableiten.
    """
    pick = _find_largest_snipping_window()
    if pick is None:
        raise MeasureError(
            "Snipping-Tool-Fenster nicht gefunden.\n\n"
            "Fenster mit markiertem Text sichtbar lassen (Titel enthält z. B. „Snipping Tool“), "
            "dann erneut „Markierung abgeschlossen – Boxen vermessen“ klicken."
        )

    bbox = (pick.left, pick.top, pick.right, pick.bottom)
    pil = ImageGrab.grab(bbox=bbox)
    iw, ih = pil.size
    captures_dir = repo_root / "captures"
    captures_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cap_path = captures_dir / f"marked_snip_{stamp}.png"
    pil.save(str(cap_path), format="PNG", optimize=True)

    bgr = cv2.cvtColor(np.asarray(pil, dtype=np.uint8), cv2.COLOR_RGB2BGR)
    if not anchor_patch_path.is_file():
        raise MeasureError(f"Anker-Patch fehlt:\n{anchor_patch_path}")

    tmpl = cv2.imread(str(anchor_patch_path), cv2.IMREAD_COLOR)
    if tmpl is None:
        raise MeasureError(f"Anker-Patch konnte nicht gelesen werden:\n{anchor_patch_path}")

    th, tw = tmpl.shape[:2]
    img_h, img_w = bgr.shape[:2]
    if th > img_h or tw > img_w:
        raise MeasureError("Anker-Vorlage ist größer als das Snipping-Fenster-Bild.")

    res = cv2.matchTemplate(bgr, tmpl, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    score = float(max_val)
    anchor_thresh = 0.42
    if score < anchor_thresh:
        raise MeasureError(
            f"Anker im Snipping-Screenshot nicht zuverlässig gefunden (Treffer {score:.2f}, "
            f"Schwelle {anchor_thresh:.2f}).\n\n"
            "Kalibrierung am Referenzbild prüfen oder ein größeres/markierteres Fenster nutzen."
        )

    ax, ay = int(max_loc[0]), int(max_loc[1])
    boxes = _detect_snipping_text_boxes(bgr)

    snipping_text_boxes: list[dict] = []
    for i, (x1, y1, x2, y2) in enumerate(boxes):
        snipping_text_boxes.append(
            {
                "box_index": i,
                "abs": {"x1": x1, "y1": y1, "x2": x2, "y2": y2},
                "rel": {
                    "x1": x1 - ax,
                    "y1": y1 - ay,
                    "x2": x2 - ax,
                    "y2": y2 - ay,
                },
            }
        )

    ts = datetime.now().isoformat(timespec="seconds")
    return {
        "marked_capture": {
            "path": str(cap_path.resolve()),
            "width": iw,
            "height": ih,
            "captured_at": ts,
            "window_title": pick.title,
            "window_rect": {
                "left": pick.left,
                "top": pick.top,
                "right": pick.right,
                "bottom": pick.bottom,
            },
        },
        "anchor": {
            "x": ax,
            "y": ay,
            "found": True,
            "match_score": score,
            "template_width": tw,
            "template_height": th,
        },
        "snipping_text_boxes": snipping_text_boxes,
    }


def filter_boxes_for_mapping(
    snipping_text_boxes: list[dict],
    img_w: int,
    img_h: int,
) -> tuple[list[dict], list[dict]]:
    """
    Zu große/kleine Kandidaten aus der MVP-Zuordnung nehmen; nutzbare Liste sortiert
    oben→unten, links→rechts.
    """
    usable: list[dict] = []
    ignored: list[dict] = []
    img_area = max(1, int(img_w) * int(img_h))
    for b in snipping_text_boxes:
        abs_ = b["abs"]
        w = int(abs_["x2"]) - int(abs_["x1"])
        h = int(abs_["y2"]) - int(abs_["y1"])
        area = w * h
        row = dict(b)
        if area >= 0.82 * img_area:
            row["ignored_reason"] = "too_large"
            ignored.append(row)
            continue
        if w >= 0.90 * img_w and h >= 0.50 * img_h:
            row["ignored_reason"] = "too_large"
            ignored.append(row)
            continue
        if area < 130:
            row["ignored_reason"] = "too_small"
            ignored.append(row)
            continue
        max_map_area_ratio = 0.32
        max_map_w_ratio = 0.55
        max_map_h_ratio = 0.50
        if (
            area > max_map_area_ratio * img_area
            or w > max_map_w_ratio * img_w
            or h > max_map_h_ratio * img_h
        ):
            row = dict(b)
            row["ignored_reason"] = "too_large_for_region_mapping"
            ignored.append(row)
            continue
        usable.append(b)
    usable.sort(key=lambda bb: (bb["abs"]["y1"], bb["abs"]["x1"]))
    return usable, ignored


# --- Layout-Zonen (normierte Snipping-Bildkoordinaten, Box-Mittelpunkt) ---
_POT_REGION_NAMES = frozenset(
    {
        "c0pot_total",
        "c0pot_bets",
        "c0pot0",
        "c0smallblind",
        "c0bigblind",
    }
)


def expected_layout_zone(region_name: str) -> str | None:
    """Erwartete UI-Zone für eine Region (Player-Sitz, Pot, Game/Hand, Buttons)."""
    m = re.match(r"^p([0-9])(name|balance|dealer|bet)$", region_name)
    if m:
        return f"player_{m.group(1)}"
    if region_name in _POT_REGION_NAMES:
        return "pot"
    if region_name in ("game_id", "hand_id", "street"):
        return "game_hand"
    if re.match(r"^i\d+label$", region_name):
        return "button"
    return None


def _box_norm_center(abs_box: dict, img_w: int, img_h: int) -> tuple[float, float]:
    x1, y1, x2, y2 = (
        float(abs_box["x1"]),
        float(abs_box["y1"]),
        float(abs_box["x2"]),
        float(abs_box["y2"]),
    )
    iw, ih = max(1, int(img_w)), max(1, int(img_h))
    return ((x1 + x2) * 0.5 / iw, (y1 + y2) * 0.5 / ih)


def classify_box_layout_zone(abs_box: dict, img_w: int, img_h: int) -> str:
    """
    Heuristik: Game/Hand oben, p5/p6 oben links/mitte, Hero (p0) unten mitte,
    Buttons als unterste Leiste, Pot zentral, sonst Sitz-Umlauf (leicht versetzt).
    """
    cx, cy = _box_norm_center(abs_box, img_w, img_h)
    dx, dy = cx - 0.5, cy - 0.5
    dist = math.hypot(dx, dy)

    if cy < 0.35 and cx < 0.66:
        return "game_hand"

    if cy < 0.44:
        if cx < 0.34:
            return "player_5"
        if cx < 0.53 and cy < 0.43:
            return "player_6"

    if cy > 0.60 and abs(cx - 0.5) < 0.38 and 0.10 < dist < 0.52:
        if cy > 0.64 or abs(cx - 0.5) < 0.26:
            return "player_0"

    if cy > 0.82 and 0.14 < cx < 0.86:
        return "button"
    if cy > 0.70 and 0.20 < cx < 0.80 and dist > 0.38:
        return "button"

    if dist < 0.185:
        return "pot"
    if abs(dx) + abs(dy) < 0.038:
        return "pot"

    ang = math.atan2(dy, dx)
    seat_phase = 0.14
    t = (math.pi / 2 - ang + seat_phase) / (2 * math.pi / 10)
    seat = int(round(t)) % 10
    return f"player_{seat}"


class _UnionFind:
    def __init__(self, n: int) -> None:
        self._p = list(range(n))
        self._r = [0] * n

    def find(self, x: int) -> int:
        while self._p[x] != x:
            self._p[x] = self._p[self._p[x]]
            x = self._p[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return
        if self._r[ra] < self._r[rb]:
            ra, rb = rb, ra
        self._p[rb] = ra
        if self._r[ra] == self._r[rb]:
            self._r[ra] += 1


def _merge_token_clusters_from_regions(regions: list[dict], n_tokens: int) -> _UnionFind:
    """Verknüpft Tokens, die dieselbe sichtbare Box teilen können (name+balance, Multi-Index-Regionen)."""
    uf = _UnionFind(n_tokens)
    for r in regions:
        sti = r.get("source_token_indices") or []
        sti_i = [int(x) for x in sti if x is not None and 0 <= int(x) < n_tokens]
        for k in range(1, len(sti_i)):
            uf.union(sti_i[0], sti_i[k])

    by_seat: dict[str, dict[str, list[int]]] = {}
    for r in regions:
        rn = r["region_name"]
        m = re.match(r"^p([0-9])(name|balance)$", rn)
        if not m:
            continue
        seat, kind = m.group(1), m.group(2)
        sti = r.get("source_token_indices") or []
        if not sti:
            continue
        t0 = int(sti[0])
        if not (0 <= t0 < n_tokens):
            continue
        by_seat.setdefault(seat, {})[kind] = sti

    for _seat, parts in by_seat.items():
        name_ti = parts.get("name")
        bal_ti = parts.get("balance")
        if not name_ti or not bal_ti:
            continue
        a, b = int(name_ti[0]), int(bal_ti[0])
        if 0 <= a < n_tokens and 0 <= b < n_tokens:
            uf.union(a, b)
    return uf


def _root_to_target_zones(uf: _UnionFind, regions: list[dict], n_tokens: int) -> tuple[dict[int, str | None], set[int]]:
    roots: dict[int, str | None] = {}
    conflict_root: set[int] = set()
    for r in regions:
        z = expected_layout_zone(r["region_name"])
        if z is None:
            continue
        for tx in r.get("source_token_indices") or []:
            ti = int(tx)
            if not (0 <= ti < n_tokens):
                continue
            root = int(uf.find(ti))
            if root in conflict_root:
                continue
            prev = roots.get(root)
            if prev is None:
                roots[root] = z
            elif prev != z:
                conflict_root.add(root)
                roots[root] = None
    return roots, conflict_root


def map_tokens_to_boxes_many_to_one(
    tokens: list[dict],
    regions: list[dict],
    usable_boxes: list[dict],
    *,
    capture_width: int,
    capture_height: int,
) -> dict:
    """
    Zonenbasiert: Cluster (z. B. p4name+p4balance) → Box in passender Player-4-Zone usw.;
    Pot/Game-Hand/Button ebenso. Innerhalb einer Zone: oben→unten, links→rechts.
    """
    n_t = len(tokens)
    token_box_map: dict[str, int] = {}
    unmatched_tokens: list[int] = []
    unmatched_boxes: list[int] = []
    box_layout_zones: dict[str, str] = {}

    if n_t == 0 or not usable_boxes:
        if n_t:
            unmatched_tokens = [int(t.get("token_index", i)) for i, t in enumerate(tokens)]
        unmatched_boxes = [int(b["box_index"]) for b in usable_boxes]
        return {
            "token_box_map": token_box_map,
            "mapping_confidence": "none",
            "unmatched_tokens": unmatched_tokens,
            "unmatched_boxes": unmatched_boxes,
            "forced_cluster_merges": 0,
            "box_layout_zones": box_layout_zones,
        }

    iw, ih = int(capture_width), int(capture_height)
    uf = _merge_token_clusters_from_regions(regions, n_t)
    root_zone_ok, conflict_root = _root_to_target_zones(uf, regions, n_t)

    boxes_by_idx = {int(b["box_index"]): b for b in usable_boxes}

    zone_queues: dict[str, deque[int]] = defaultdict(deque)
    for b in usable_boxes:
        bi = int(b["box_index"])
        zn = classify_box_layout_zone(b["abs"], iw, ih)
        box_layout_zones[str(bi)] = zn
        zone_queues[zn].append(bi)

    for zn, q in list(zone_queues.items()):
        sorted_bis = sorted(
            q,
            key=lambda bi: (
                boxes_by_idx[bi]["abs"]["y1"],
                boxes_by_idx[bi]["abs"]["x1"],
            ),
        )
        zone_queues[zn] = deque(sorted_bis)

    root_to_tokens: defaultdict[int, list[int]] = defaultdict(list)
    for i in range(n_t):
        root_to_tokens[uf.find(i)].append(i)
    for t_list in root_to_tokens.values():
        t_list.sort()

    zone_miss = False
    for root in sorted(root_to_tokens.keys(), key=lambda rr: min(root_to_tokens[rr])):
        toks = root_to_tokens[root]
        if root in conflict_root or root_zone_ok.get(root) is None:
            continue
        z_need = root_zone_ok[root]
        q = zone_queues.get(z_need)
        if not q:
            zone_miss = True
            unmatched_tokens.extend(toks)
            continue
        bi = q.popleft()
        for tidx in toks:
            token_box_map[str(int(tidx))] = bi

    assigned_box_indices = set(token_box_map.values())
    for b in usable_boxes:
        bi = int(b["box_index"])
        if bi not in assigned_box_indices:
            unmatched_boxes.append(bi)

    for i in range(n_t):
        if str(i) not in token_box_map:
            unmatched_tokens.append(i)

    unmatched_tokens = sorted(set(unmatched_tokens))
    conf = "medium" if (zone_miss or conflict_root) else "high"

    return {
        "token_box_map": token_box_map,
        "mapping_confidence": conf,
        "unmatched_tokens": unmatched_tokens,
        "unmatched_boxes": sorted(set(unmatched_boxes)),
        "forced_cluster_merges": 0,
        "box_layout_zones": box_layout_zones,
    }


def build_region_boxes(
    regions: list[dict],
    mapping: dict,
    boxes_by_index: dict[int, dict],
) -> list[dict]:
    """Region → source_token_index → box; matched bei high/medium (Token↔Box-Map vorhanden)."""
    conf = mapping.get("mapping_confidence", "none")
    tbm: dict[str, int] = mapping.get("token_box_map") or {}
    allow_match = conf in ("high", "medium")
    out: list[dict] = []
    for r in regions:
        rn = r["region_name"]
        val = r["value"]
        sti = list(r.get("source_token_indices") or [])
        entry: dict = {
            "region_name": rn,
            "value": val,
            "source_token_indices": sti,
            "layout_zone": expected_layout_zone(rn),
            "box_index": None,
            "abs": None,
            "rel": None,
            "geometry_status": "unmatched",
            "geometry_source": "snipping_box",
            "box_layout_zone": None,
            "unmatched_reason": None,
        }
        if not allow_match:
            out.append(entry)
            continue
        if not sti:
            entry["unmatched_reason"] = "no_source_tokens"
            out.append(entry)
            continue
        box_ids: list[int] = []
        for tx in sti:
            sid = str(int(tx))
            if sid not in tbm:
                box_ids = []
                break
            box_ids.append(int(tbm[sid]))
        if not box_ids or len(set(box_ids)) != 1:
            out.append(entry)
            continue
        bi = box_ids[0]
        boxd = boxes_by_index.get(bi)
        if not boxd:
            out.append(entry)
            continue
        blz = (mapping.get("box_layout_zones") or {}).get(str(bi))
        exp_z = entry.get("layout_zone")
        if exp_z is not None and blz is not None and blz != exp_z:
            entry["unmatched_reason"] = "layout_zone_mismatch"
            out.append(entry)
            continue
        entry["box_index"] = bi
        entry["abs"] = dict(boxd["abs"])
        entry["rel"] = dict(boxd["rel"])
        entry["box_layout_zone"] = blz
        entry["geometry_status"] = "matched_box"
        entry["geometry_source"] = "snipping_box"
        out.append(entry)
    return out


def save_mapping_debug_image(
    capture_path: Path,
    snipping_text_boxes: list[dict],
    region_boxes: list[dict],
    repo_root: Path,
    *,
    box_layout_zones: dict[str, str] | None = None,
    ignored_snipping_boxes: list[dict] | None = None,
) -> Path | None:
    """
    Debug: box_index, box_layout_zone, gematchte region_name;
    ignorierte Boxen mit ignored_reason (andere Farbe).
    """
    img = cv2.imread(str(capture_path))
    if img is None:
        return None
    zl = box_layout_zones or {}

    box_to_regions: dict[int, list[str]] = {}
    for rb in region_boxes:
        if rb.get("geometry_status") != "matched_box":
            continue
        bi = rb.get("box_index")
        if bi is None:
            continue
        box_to_regions.setdefault(int(bi), []).append(str(rb["region_name"]))

    ignored_idx = {int(x["box_index"]) for x in (ignored_snipping_boxes or [])}

    for b in snipping_text_boxes:
        bi = int(b["box_index"])
        if bi in ignored_idx:
            continue
        a = b["abs"]
        x1, y1, x2, y2 = int(a["x1"]), int(a["y1"]), int(a["x2"]), int(a["y2"])
        cv2.rectangle(img, (x1, y1), (x2, y2), (0, 200, 0), 2)
        zone = zl.get(str(bi), "?")
        line1 = f"{bi} z={zone}"
        names = box_to_regions.get(bi)
        line2 = ",".join(names[:4]) if names else ""
        cv2.putText(
            img,
            line1,
            (x1, max(14, y1 - 6)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.38,
            (0, 220, 255),
            1,
            cv2.LINE_AA,
        )
        if line2:
            cv2.putText(
                img,
                line2[:60],
                (x1, max(28, y1 + 10)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.36,
                (180, 255, 180),
                1,
                cv2.LINE_AA,
            )

    for ign in ignored_snipping_boxes or []:
        bi = int(ign["box_index"])
        a = ign["abs"]
        x1, y1, x2, y2 = int(a["x1"]), int(a["y1"]), int(a["x2"]), int(a["y2"])
        cv2.rectangle(img, (x1, y1), (x2, y2), (0, 100, 255), 2)
        reason = str(ign.get("ignored_reason", ""))
        cv2.putText(
            img,
            f"{bi} IGN {reason}"[:72],
            (x1, max(14, y1 - 3)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.36,
            (0, 220, 255),
            1,
            cv2.LINE_AA,
        )
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = repo_root / "captures" / f"snipping_marked_boxes_debug_{stamp}.png"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(out_path), img)
    return out_path
