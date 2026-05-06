"""
Snipping-Tool-Fenster erfassen (BBox), Anker per Template-Matching, heuristische Textboxen.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

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
        usable.append(b)
    usable.sort(key=lambda bb: (bb["abs"]["y1"], bb["abs"]["x1"]))
    return usable, ignored


def _classify_mapping_confidence(n_tok: int, n_box: int) -> str:
    if n_tok == 0 or n_box == 0:
        return "none"
    mx = max(n_tok, n_box)
    gap = mx - min(n_tok, n_box)
    if gap > max(5, int(0.22 * mx)):
        return "low"
    if gap > max(2, int(0.10 * mx)):
        return "medium"
    return "high"


def map_ordered_tokens_to_boxes(
    tokens: list[dict],
    usable_boxes: list[dict],
) -> dict:
    """1:1 in Clipboard-/Token-Reihenfolge zu sortierten Boxen; bei low confidence kein Map."""
    ordered = sorted(tokens, key=lambda t: int(t.get("token_index", -1)))
    n_t, n_b = len(ordered), len(usable_boxes)
    conf = _classify_mapping_confidence(n_t, n_b)
    unmatched_tokens: list[int] = []
    unmatched_boxes: list[int] = []
    token_box_map: dict[str, int] = {}

    if conf in ("none", "low"):
        unmatched_tokens = [int(t["token_index"]) for t in ordered]
        unmatched_boxes = [int(b["box_index"]) for b in usable_boxes]
        return {
            "token_box_map": token_box_map,
            "mapping_confidence": conf,
            "unmatched_tokens": unmatched_tokens,
            "unmatched_boxes": unmatched_boxes,
        }

    k = min(n_t, n_b)
    for j in range(k):
        token_box_map[str(int(ordered[j]["token_index"]))] = int(usable_boxes[j]["box_index"])
    if n_t > k:
        unmatched_tokens = [int(ordered[j]["token_index"]) for j in range(k, n_t)]
    if n_b > k:
        unmatched_boxes = [int(usable_boxes[j]["box_index"]) for j in range(k, n_b)]

    return {
        "token_box_map": token_box_map,
        "mapping_confidence": conf,
        "unmatched_tokens": unmatched_tokens,
        "unmatched_boxes": unmatched_boxes,
    }


def build_region_boxes(
    regions: list[dict],
    mapping: dict,
    boxes_by_index: dict[int, dict],
) -> list[dict]:
    """Region → erste source_token_index → box; nur bei mapping_confidence „high“ als matched."""
    conf = mapping.get("mapping_confidence", "none")
    tbm: dict[str, int] = mapping.get("token_box_map") or {}
    allow_match = conf == "high"
    out: list[dict] = []
    for r in regions:
        rn = r["region_name"]
        val = r["value"]
        sti = list(r.get("source_token_indices") or [])
        entry: dict = {
            "region_name": rn,
            "value": val,
            "source_token_indices": sti,
            "box_index": None,
            "abs": None,
            "rel": None,
            "geometry_status": "unmatched",
        }
        if not allow_match or not sti:
            out.append(entry)
            continue
        t0 = sti[0]
        sid = str(int(t0))
        if sid not in tbm:
            out.append(entry)
            continue
        bi = int(tbm[sid])
        boxd = boxes_by_index.get(bi)
        if not boxd:
            out.append(entry)
            continue
        entry["box_index"] = bi
        entry["abs"] = dict(boxd["abs"])
        entry["rel"] = dict(boxd["rel"])
        entry["geometry_status"] = "matched"
        out.append(entry)
    return out


def save_mapping_debug_image(
    capture_path: Path,
    snipping_text_boxes: list[dict],
    region_boxes: list[dict],
    repo_root: Path,
) -> Path | None:
    """Optional: alle Boxen mit Index; bei gematchten Regionen region_name anhängen."""
    img = cv2.imread(str(capture_path))
    if img is None:
        return None
    box_to_regions: dict[int, list[str]] = {}
    for rb in region_boxes:
        bi = rb.get("box_index")
        if bi is None:
            continue
        box_to_regions.setdefault(int(bi), []).append(str(rb["region_name"]))

    for b in snipping_text_boxes:
        bi = int(b["box_index"])
        a = b["abs"]
        x1, y1, x2, y2 = int(a["x1"]), int(a["y1"]), int(a["x2"]), int(a["y2"])
        cv2.rectangle(img, (x1, y1), (x2, y2), (0, 200, 0), 2)
        label = str(bi)
        names = box_to_regions.get(bi)
        if names:
            label += " " + ",".join(names[:3])
        cv2.putText(
            img,
            label,
            (x1, max(14, y1 - 3)),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.42,
            (0, 220, 255),
            1,
            cv2.LINE_AA,
        )
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = repo_root / "captures" / f"snipping_marked_boxes_debug_{stamp}.png"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(out_path), img)
    return out_path
