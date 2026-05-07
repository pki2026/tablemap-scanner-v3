"""
Textpixel-Ursprung pro Region: feste Suchzonen relativ zum Anker, keine Snipping-Box-Abhängigkeit.

Erkannte Werte kommen weiter aus Parser/Heuristik (main.group_tokens_into_regions).
Dieses Modul vermisst nur sichtbare Vordergrund-Pixel im Capture.
Tablemap-Koordinate text_origin_* = linker oberer Punkt der Text-Bounding-Box (x1/y1).
"""

from __future__ import annotations

import math
import re
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

import cv2
import numpy as np

# Nominaler Referenzrahmen für Zonen (Skalierung auf tatsächliche Capture-Größe).
REF_W = 1600
REF_H = 900

_GEOMETRY_SOURCE = "text_pixel_origin"


def _scale_rect(
    x1: int, y1: int, x2: int, y2: int, cap_w: int, cap_h: int
) -> tuple[int, int, int, int]:
    sx = cap_w / REF_W
    sy = cap_h / REF_H
    return (
        int(round(x1 * sx)),
        int(round(y1 * sy)),
        int(round(x2 * sx)),
        int(round(y2 * sy)),
    )


def _seat_xy(seat: int, n_seats: int = 10) -> tuple[float, float]:
    """Sitzmittelpunkt auf einer Ellipse; Sitz 0 unten (Hero)."""
    t = 2 * math.pi * (seat / n_seats)
    cx, cy = REF_W * 0.5, REF_H * 0.40
    rx, ry = REF_W * 0.41, REF_H * 0.36
    x = cx + rx * math.sin(t)
    y = cy + ry * math.cos(t)
    return x, y


def _player_field_rect(seat: int, field: str) -> tuple[int, int, int, int] | None:
    x, y = _seat_xy(seat)
    ix, iy = int(x), int(y)
    if field == "name":
        return ix - 110, iy - 58, ix + 100, iy - 8
    if field == "balance":
        return ix - 55, iy + 5, ix + 145, iy + 62
    if field == "bet":
        return ix - 70, iy - 108, ix + 88, iy - 62
    if field == "dealer":
        return ix - 118, iy - 118, ix - 38, iy - 48
    if field == "status":
        return ix - 90, iy + 58, ix + 110, iy + 112
    return None


def _default_zone_for_region(region_name: str) -> tuple[int, int, int, int] | None:
    m = re.fullmatch(r"p([0-9])(name|balance|bet|dealer|status)", region_name)
    if m:
        return _player_field_rect(int(m.group(1)), m.group(2))

    if region_name == "c0pot_total":
        return 670, 300, 930, 380
    if region_name == "c0pot_bets":
        return 670, 385, 930, 455
    if region_name == "c0pot0":
        return 720, 250, 880, 310
    if region_name == "c0smallblind":
        return 520, 265, 760, 330
    if region_name == "c0bigblind":
        return 840, 265, 1080, 330

    if region_name in ("game_id", "hand_id"):
        return 420, 6, 1180, 64
    if region_name == "street":
        return 600, 68, 1000, 118

    bm = re.fullmatch(r"board_card_(\d+)", region_name)
    if bm:
        n = int(bm.group(1))
        ox = 560 + n * 72
        return ox, 175, ox + 68, 245

    fm = re.fullmatch(r"flop_card_(\d+)", region_name)
    if fm:
        n = int(fm.group(1))
        ox = 560 + n * 72
        return ox, 175, ox + 68, 245
    if region_name == "turn_card":
        return 788, 175, 860, 245
    if region_name == "river_card":
        return 868, 175, 940, 245

    hm = re.fullmatch(r"hero_card_(\d+)", region_name)
    if hm:
        n = int(hm.group(1))
        ox = 668 + n * 62
        return ox, 718, ox + 58, 788

    im = re.fullmatch(r"i(\d+)label", region_name)
    if im:
        idx = int(im.group(1))
        bx = 180 + idx * 118
        return bx, 808, bx + 108, 878

    if region_name in ("betsize_hero", "raisesize_hero", "callsize_hero"):
        return 1050, 780, 1280, 840

    return None


def _merge_hints(
    region_zone_hints: Mapping[str, Mapping[str, int]] | None,
    region_name: str,
    rect_ref: tuple[int, int, int, int] | None,
    cap_w: int,
    cap_h: int,
) -> tuple[int, int, int, int] | None:
    if region_zone_hints and region_name in region_zone_hints:
        z = region_zone_hints[region_name]
        x1, y1, x2, y2 = int(z["x1"]), int(z["y1"]), int(z["x2"]), int(z["y2"])
        return max(0, x1), max(0, y1), max(x1 + 1, x2), max(y1 + 1, y2)
    if rect_ref is None:
        return None
    return _scale_rect(*rect_ref, cap_w, cap_h)


def _foreground_mask(gray: np.ndarray) -> np.ndarray:
    """Trennt schriftähnliche Vordergrundpixel vom Hintergrund (hell/dunkel robust)."""
    if gray.size == 0:
        return np.zeros((0, 0), dtype=np.uint8)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    mean = float(np.mean(blur))
    if mean > 130:
        _, th = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    else:
        _, th = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    th = cv2.morphologyEx(th, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)
    return th


def _measure_text_in_crop(
    crop_bgr: np.ndarray,
    zone_offset_x: int,
    zone_offset_y: int,
    anchor_x: int,
    anchor_y: int,
) -> dict[str, Any]:
    if crop_bgr.size == 0:
        return {
            "text_origin_abs": None,
            "text_origin_rel": None,
            "text_bbox_abs": None,
            "text_bbox_rel": None,
            "geometry_status": "text_pixels_not_found",
            "geometry_source": _GEOMETRY_SOURCE,
        }

    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    mask = _foreground_mask(gray)
    h, w = mask.shape[:2]
    min_area = max(28, int(0.00035 * w * h))

    nlab, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    keep = np.zeros_like(mask, dtype=np.uint8)
    for i in range(1, nlab):
        if int(stats[i, cv2.CC_STAT_AREA]) >= min_area:
            keep[labels == i] = 255

    ys, xs = np.where(keep > 0)
    if len(xs) == 0:
        return {
            "text_origin_abs": None,
            "text_origin_rel": None,
            "text_bbox_abs": None,
            "text_bbox_rel": None,
            "geometry_status": "text_pixels_not_found",
            "geometry_source": _GEOMETRY_SOURCE,
        }

    order = np.lexsort((xs, ys))
    ox_c, oy_c = int(xs[order[0]]), int(ys[order[0]])
    xa, xb = int(xs.min()), int(xs.max()) + 1
    ya, yb = int(ys.min()), int(ys.max()) + 1

    bx1 = xa + zone_offset_x
    by1 = ya + zone_offset_y
    bx2 = xb + zone_offset_x
    by2 = yb + zone_offset_y
    first_abs_x = ox_c + zone_offset_x
    first_abs_y = oy_c + zone_offset_y

    return {
        "text_origin_abs": {"x": bx1, "y": by1},
        "text_origin_rel": {"x": bx1 - anchor_x, "y": by1 - anchor_y},
        "text_bbox_abs": {"x1": bx1, "y1": by1, "x2": bx2, "y2": by2},
        "text_bbox_rel": {
            "x1": bx1 - anchor_x,
            "y1": by1 - anchor_y,
            "x2": bx2 - anchor_x,
            "y2": by2 - anchor_y,
        },
        "text_first_pixel_abs": {"x": first_abs_x, "y": first_abs_y},
        "text_first_pixel_rel": {
            "x": first_abs_x - anchor_x,
            "y": first_abs_y - anchor_y,
        },
        "geometry_status": "matched_text_pixels",
        "geometry_source": _GEOMETRY_SOURCE,
    }


def measure_text_pixel_origins(
    image_path: str,
    anchor_abs: Mapping[str, int],
    regions_detail: Iterable[Mapping[str, Any]],
    region_zone_hints: Mapping[str, Mapping[str, int]] | None = None,
) -> dict[str, dict[str, Any]]:
    """
    Bestimmt pro Region die Text-Bounding-Box; text_origin_* ist deren linker oberer Punkt (x1/y1).

    Eingabe:
    - image_path: Screenshot/Capture (z. B. unter captures/).
    - anchor_abs: absoluter Ankerpunkt im Bild, z. B. {"x": 337, "y": 167}
    - regions_detail: erkannte Regionen mit mindestens region_name (value optional, nur Echo)
    - region_zone_hints: optionale Suchzonen je Region relativ zum Anker {x1,y1,x2,y2}

    Ausgabe:
    - region_name -> Geometrie-Dict (text_origin_*, text_bbox_*, optional text_first_pixel_*, geometry_status, geometry_source)
    """
    path = Path(image_path)
    ax = int(anchor_abs["x"])
    ay = int(anchor_abs["y"])

    def _row_names() -> list[str]:
        return [
            str(row.get("region_name", "")).strip()
            for row in regions_detail
            if str(row.get("region_name", "")).strip()
        ]

    if not path.is_file():
        out: dict[str, dict[str, Any]] = {}
        for row in regions_detail:
            name = str(row.get("region_name", "")).strip()
            if not name:
                continue
            g: dict[str, Any] = {
                "text_origin_abs": None,
                "text_origin_rel": None,
                "text_bbox_abs": None,
                "text_bbox_rel": None,
                "geometry_status": "text_pixels_not_found",
                "geometry_source": _GEOMETRY_SOURCE,
            }
            if "value" in row:
                g["value"] = row.get("value")
            out[name] = g
        return out

    bgr = cv2.imread(str(path), cv2.IMREAD_COLOR)
    if bgr is None:
        out = {}
        for row in regions_detail:
            name = str(row.get("region_name", "")).strip()
            if not name:
                continue
            g = {
                "text_origin_abs": None,
                "text_origin_rel": None,
                "text_bbox_abs": None,
                "text_bbox_rel": None,
                "geometry_status": "text_pixels_not_found",
                "geometry_source": _GEOMETRY_SOURCE,
            }
            if "value" in row:
                g["value"] = row.get("value")
            out[name] = g
        return out

    img_h, img_w = bgr.shape[:2]
    out: dict[str, dict[str, Any]] = {}

    for row in regions_detail:
        name = str(row.get("region_name", "")).strip()
        if not name:
            continue
        ref_rect = _default_zone_for_region(name)
        z = _merge_hints(region_zone_hints, name, ref_rect, img_w, img_h)
        if z is None:
            out[name] = {
                "text_origin_abs": None,
                "text_origin_rel": None,
                "text_bbox_abs": None,
                "text_bbox_rel": None,
                "geometry_status": "text_pixels_not_found",
                "geometry_source": _GEOMETRY_SOURCE,
            }
            continue

        zx1, zy1, zx2, zy2 = z
        zx1 = max(0, min(zx1, img_w - 1))
        zy1 = max(0, min(zy1, img_h - 1))
        zx2 = max(zx1 + 1, min(zx2, img_w))
        zy2 = max(zy1 + 1, min(zy2, img_h))

        crop = bgr[zy1:zy2, zx1:zx2]
        geom = _measure_text_in_crop(crop, zx1, zy1, ax, ay)
        if "value" in row:
            geom = dict(geom)
            geom["value"] = row.get("value")
        out[name] = geom

    return out


__all__ = ["measure_text_pixel_origins", "REF_W", "REF_H"]
