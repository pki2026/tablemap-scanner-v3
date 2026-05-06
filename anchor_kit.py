"""
Echtes Desktop-Overlay: Vollbild, topmost, halbtransparent — roter Rahmen liegt über dem Tisch.
Separates kleines Bedienfenster nur für [Anker speichern] / [Abbrechen].
"""

from __future__ import annotations

import json
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from PIL import ImageGrab
import tkinter as tk
from tkinter import messagebox

REPO_ROOT = Path(__file__).resolve().parent
ANCHORS_DIR = REPO_ROOT / "anchors"
ANCHOR_PATCH_PATH = ANCHORS_DIR / "table_anchor_patch.png"
ANCHOR_CONFIG_PATH = ANCHORS_DIR / "table_anchor_config.json"
ANCHOR_SCHEMA = "tablemap_scanner_v3_anchor"


def _safe_grab_release(w: tk.Misc) -> None:
    try:
        w.grab_release()
    except tk.TclError:
        pass


def _destroy_pair(overlay: tk.Toplevel, toolbar: tk.Toplevel | None) -> None:
    _safe_grab_release(overlay)
    if toolbar is not None:
        _safe_grab_release(toolbar)
        try:
            toolbar.destroy()
        except tk.TclError:
            pass
    try:
        overlay.destroy()
    except tk.TclError:
        pass


def run_anchor_calibration_blocking(root: tk.Tk, *, modal: bool = True) -> bool:
    """
    Vollbild-Overlay (Desktop), Rahmen in Bildschirmkoordinaten.
    Kleines Topmost-Bedienfenster nur für Speichern/Abbrechen.
    """
    print("[V3] anchor calibration: true desktop overlay", flush=True)

    result: dict[str, bool] = {"saved": False, "closing": False}

    try:
        sw_i = int(root.winfo_screenwidth())
        sh_i = int(root.winfo_screenheight())
        print(f"[V3] overlay: screen size {sw_i}x{sh_i}", flush=True)

        overlay = tk.Toplevel(root)

        overlay.overrideredirect(True)
        overlay.geometry(f"{sw_i}x{sh_i}+0+0")
        overlay.attributes("-topmost", True)
        try:
            overlay.attributes("-alpha", 0.22)
            print("[V3] overlay: fullscreen transparent (alpha)", flush=True)
        except tk.TclError:
            print("[V3] overlay: alpha not supported, using opaque base", flush=True)

        overlay.configure(bg="#000000")

        canvas = tk.Canvas(
            overlay,
            width=sw_i,
            height=sh_i,
            highlightthickness=0,
            bd=0,
            bg="#000000",
        )
        canvas.place(x=0, y=0, width=sw_i, height=sh_i)

        tw, th = 460, 72
        tx = max(0, (sw_i - tw) // 2)
        ty = max(0, sh_i - th - 16)

        toolbar = tk.Toplevel(root)
        toolbar.title("Anker — Bedienung")
        toolbar.attributes("-topmost", True)
        toolbar.resizable(False, False)
        toolbar.geometry(f"{tw}x{th}+{tx}+{ty}")
        toolbar.configure(bg="#101010")
        bf = tk.Frame(toolbar, bg="#101010", padx=10, pady=10)
        bf.pack(fill=tk.BOTH, expand=True)
        btn_row = tk.Frame(bf, bg="#101010")
        btn_row.pack()

        print("[V3] overlay: geometry fullscreen +0+0 (no title bar)", flush=True)
        print("[V3] overlay: separate toolbar window", flush=True)
        print("[V3] overlay: widgets created", flush=True)

        HANDLE = 18
        rect_id = 0
        br_id = 0
        br_diag1 = 0
        br_diag2 = 0
        origin_cross_h = 0
        origin_cross_v = 0
        origin_ring = 0
        origin_label = 0
        ORIGIN_ARM = 10

        state: dict[str, float | int | str] = {
            "mode": "idle",
            "dx": 0.0,
            "dy": 0.0,
            "corner_xoff": 0.0,
            "corner_yoff": 0.0,
            "fix_ix": 0,
            "fix_iy": 0,
            "w": 400,
            "h": 300,
        }

        def clamp_rect(x_: float, y_: float, w_: float, h_: float) -> tuple[int, int, int, int]:
            cw, ch = sw_i, sh_i
            w_i, h_i = max(16, min(int(w_), cw)), max(16, min(int(h_), ch))
            xi, yi = int(x_), int(y_)
            if xi + w_i > cw:
                xi = cw - w_i
            if yi + h_i > ch:
                yi = ch - h_i
            xi, yi = max(0, xi), max(0, yi)
            return xi, yi, w_i, h_i

        def init_rect_from_canvas() -> None:
            nonlocal rect_id, br_id, br_diag1, br_diag2
            nonlocal origin_cross_h, origin_cross_v, origin_ring, origin_label
            rw_des = min(520, max(160, sw_i - 80))
            rh_des = min(380, max(120, sh_i - 120))
            ix0 = max(24, (sw_i - rw_des) // 2)
            iy0 = max(40, (sh_i - rh_des) // 3)
            rect_id = canvas.create_rectangle(
                ix0, iy0, ix0 + rw_des, iy0 + rh_des, outline="#ff2020", width=4
            )
            br_id = canvas.create_rectangle(0, 0, HANDLE, HANDLE, outline="#e8f0ff", fill="#2a5ad6", width=2)
            br_diag1 = canvas.create_line(0, 0, 0, 0, fill="#ffffff", width=2)
            br_diag2 = canvas.create_line(0, 0, 0, 0, fill="#ffffff", width=2)
            origin_cross_h = canvas.create_line(0, 0, 0, 0, fill="#00c853", width=3)
            origin_cross_v = canvas.create_line(0, 0, 0, 0, fill="#00c853", width=3)
            origin_ring = canvas.create_oval(0, 0, 0, 0, outline="#004d1a", width=2, fill="#b9f6ca")
            origin_label = canvas.create_text(
                0, 0, text="0,0", anchor=tk.NW, fill="#00ff88", font=("Segoe UI", 10, "bold")
            )
            wx0, wy0, ww0, hh0 = canvas_rect_to_xywh_inner()
            canvas.coords(rect_id, wx0, wy0, wx0 + ww0, wy0 + hh0)
            state["w"], state["h"] = ww0, hh0

        def sync_handle() -> None:
            bx1, by1, bx2, by2 = canvas.coords(rect_id)
            canvas.coords(br_id, bx2 - HANDLE, by2 - HANDLE, bx2, by2)
            hx1, hy1, hx2, hy2 = canvas.coords(br_id)
            canvas.coords(br_diag1, hx1 + 4, hy1 + 4, hx2 - 4, hy2 - 4)
            canvas.coords(br_diag2, hx1 + 4, hy2 - 4, hx2 - 4, hy1 + 4)

        def sync_origin_marker() -> None:
            bx1, by1, *_rest = canvas.coords(rect_id)
            bx1, by1 = float(bx1), float(by1)
            canvas.coords(origin_cross_h, bx1 - ORIGIN_ARM, by1, bx1 + ORIGIN_ARM, by1)
            canvas.coords(origin_cross_v, bx1, by1 - ORIGIN_ARM, bx1, by1 + ORIGIN_ARM)
            pr = 4.0
            canvas.coords(origin_ring, bx1 - pr, by1 - pr, bx1 + pr, by1 + pr)
            canvas.coords(origin_label, bx1 + ORIGIN_ARM + 4, by1 + 2)

        def canvas_rect_to_xywh_inner() -> tuple[int, int, int, int]:
            bx1, by1, bx2, by2 = canvas.coords(rect_id)
            xa, ya, xb, yb = int(bx1), int(by1), int(bx2), int(by2)
            if xb <= xa:
                xb = xa + 16
            if yb <= ya:
                yb = ya + 16
            return clamp_rect(float(xa), float(ya), float(xb - xa), float(yb - ya))

        def apply_rect(ix: int, iy: int, w_: int, h_: int) -> None:
            canvas.coords(rect_id, ix, iy, ix + w_, iy + h_)
            sync_handle()
            sync_origin_marker()

        def on_canvas_down(ev: tk.Event) -> None:
            cx, cy = canvas.canvasx(ev.x), canvas.canvasy(ev.y)
            bx1, by1, bx2, by2 = canvas.coords(rect_id)
            hx1, hy1, hx2, hy2 = canvas.coords(br_id)
            if hx1 <= cx <= hx2 and hy1 <= cy <= hy2:
                ix_q, iy_q, w_q, h_q = canvas_rect_to_xywh_inner()
                state["mode"] = "resize"
                state["corner_xoff"] = float(cx - bx2)
                state["corner_yoff"] = float(cy - by2)
                state["fix_ix"] = ix_q
                state["fix_iy"] = iy_q
                state["w"] = w_q
                state["h"] = h_q
            elif bx1 <= cx <= bx2 and by1 <= cy <= by2:
                _, _, wi, hi = canvas_rect_to_xywh_inner()
                state["mode"] = "move"
                state["dx"] = cx - bx1
                state["dy"] = cy - by1
                state["w"] = wi
                state["h"] = hi
            else:
                state["mode"] = "idle"

        def on_canvas_motion(ev: tk.Event) -> None:
            cx, cy = canvas.canvasx(ev.x), canvas.canvasy(ev.y)
            m = str(state["mode"])
            if m == "move":
                nx = cx - float(state["dx"])
                ny = cy - float(state["dy"])
                ni, nj, nk, nh = clamp_rect(nx, ny, float(state["w"]), float(state["h"]))
                apply_rect(ni, nj, nk, nh)
            elif m == "resize":
                target_x2 = cx - float(state["corner_xoff"])
                target_y2 = cy - float(state["corner_yoff"])
                fi, fj = int(state["fix_ix"]), int(state["fix_iy"])
                br_x, br_y = int(target_x2), int(target_y2)
                ni, nj, nk, nh = clamp_rect(float(fi), float(fj), float(br_x - fi), float(br_y - fj))
                apply_rect(ni, nj, nk, nh)

        def on_canvas_release(_ev: tk.Event) -> None:
            state["mode"] = "idle"

        _initialized = {"ok": False}

        def on_overlay_map(_ev: tk.Event | None = None) -> None:
            if _initialized["ok"]:
                return
            _initialized["ok"] = True
            init_rect_from_canvas()
            sync_handle()
            sync_origin_marker()
            canvas.bind("<ButtonPress-1>", on_canvas_down)
            canvas.bind("<B1-Motion>", on_canvas_motion)
            canvas.bind("<ButtonRelease-1>", on_canvas_release)
            print("[V3] overlay: visible (desktop frame ready)", flush=True)

        def close_cancel() -> None:
            if result["closing"]:
                return
            result["closing"] = True
            print("[V3] anchor calibration: cancelled", flush=True)
            _destroy_pair(overlay, toolbar)

        def save_anchor() -> None:
            ix, iy, w_, h_ = canvas_rect_to_xywh_inner()
            if w_ < 8 or h_ < 8:
                messagebox.showwarning(
                    "Anker", "Bereich zu klein (min. ca. 8×8 Pixel).", parent=toolbar
                )
                return
            cx0 = int(canvas.winfo_rootx())
            cy0 = int(canvas.winfo_rooty())
            left = cx0 + ix
            top = cy0 + iy
            right_ex = left + w_
            bottom_ex = top + h_
            x1, y1, x2, y2 = left, top, right_ex - 1, bottom_ex - 1
            print(f"[V3] anchor calibration: grab screen bbox left={left} top={top} w={w_} h={h_}", flush=True)
            try:
                overlay.withdraw()
                toolbar.withdraw()
                root.update_idletasks()
                root.update()
                time.sleep(0.15)
                try:
                    im = ImageGrab.grab(bbox=(left, top, right_ex, bottom_ex), all_screens=True)
                except TypeError:
                    im = ImageGrab.grab(bbox=(left, top, right_ex, bottom_ex))
            except Exception as ex:
                traceback.print_exc()
                print(f"[V3][ERROR] anchor grab failed: {ex}", flush=True)
                try:
                    overlay.deiconify()
                    toolbar.deiconify()
                except tk.TclError:
                    pass
                try:
                    messagebox.showerror("Anker", f"Screenshot fehlgeschlagen:\n{ex}", parent=toolbar)
                except tk.TclError:
                    pass
                return

            try:
                ANCHORS_DIR.mkdir(parents=True, exist_ok=True)
                im.save(str(ANCHOR_PATCH_PATH), format="PNG", optimize=True)
                cfg = {
                    "schema": ANCHOR_SCHEMA,
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "anchor_name": "table_anchor",
                    "origin": "top_left",
                    "origin_screen_hint": "top_left_of_rectangle",
                    "resize_handle": "bottom_right",
                    "origin_pixel": {"x": 0, "y": 0},
                    "anchor_width": w_,
                    "anchor_height": h_,
                    "capture_rect_screen": {"x1": x1, "y1": y1, "x2": x2, "y2": y2},
                    "calibration_mode": "desktop_overlay",
                    "note": "Nullpunkt = oben links im gespeicherten Patch (= linke obere Ecke des Rahmens).",
                }
                ANCHOR_CONFIG_PATH.write_text(
                    json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                print(f"[V3] anchor saved: {ANCHOR_PATCH_PATH}", flush=True)
                print(f"[V3] anchor config saved: {ANCHOR_CONFIG_PATH}", flush=True)
            except OSError as ex:
                traceback.print_exc()
                print(f"[V3][ERROR] anchor save failed: {ex}", flush=True)
                try:
                    messagebox.showerror("Anker", f"Speichern fehlgeschlagen:\n{ex}", parent=toolbar)
                except tk.TclError:
                    pass
                try:
                    overlay.deiconify()
                    toolbar.deiconify()
                except tk.TclError:
                    pass
                return

            result["saved"] = True
            result["closing"] = True
            _destroy_pair(overlay, toolbar)

        tk.Button(btn_row, text="Abbrechen", command=close_cancel, width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_row, text="Anker speichern", command=save_anchor, width=16).pack(side=tk.LEFT, padx=6)

        toolbar.protocol("WM_DELETE_WINDOW", close_cancel)

        overlay.bind("<Map>", on_overlay_map)

        root.update_idletasks()
        root.update()
        overlay.deiconify()
        overlay.lift()
        toolbar.deiconify()
        toolbar.lift()
        try:
            toolbar.focus_force()
        except Exception:
            pass

        print("[V3] overlay: entering wait_window", flush=True)

        if modal:
            overlay.wait_window()

        print("[V3] overlay: wait_window returned", flush=True)
        return bool(result.get("saved"))

    except Exception as exc:
        traceback.print_exc()
        print(f"[V3][ERROR] overlay failed: {exc}", flush=True, file=sys.stderr)
        return False
