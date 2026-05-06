"""
Anker-Kalibrierung: roter Rahmen auf dem Referenz-Screenshot (Tk + PIL).
Kein Desktop-Overlay, kein Template-Matching in diesem Modul.
"""

from __future__ import annotations

import json
import traceback
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageTk

import tkinter as tk
from tkinter import messagebox

REPO_ROOT = Path(__file__).resolve().parent
ANCHORS_DIR = REPO_ROOT / "anchors"
ANCHOR_PATCH_PATH = ANCHORS_DIR / "table_anchor_patch.png"
ANCHOR_CONFIG_PATH = ANCHORS_DIR / "table_anchor_config.json"

ANCHOR_SCHEMA = "tablemap_scanner_v3_anchor"


def saved_anchor_files_present() -> bool:
    """True, wenn Kalibrierung abgeschlossen (Patch + Config auf der Platte)."""
    return ANCHOR_PATCH_PATH.is_file() and ANCHOR_CONFIG_PATH.is_file()


def _safe_grab_release(w: tk.Misc) -> None:
    try:
        w.grab_release()
    except tk.TclError:
        pass


def run_anchor_calibration_blocking(root: tk.Tk, screenshot_path: Path, *, modal: bool = True) -> bool:
    """
    Tk-Calibrator: Bild aus ``screenshot_path``, Rechteck verschieben/skalieren, Patch speichern.
    Blockiert über ``dlg.wait_window()`` bis Speichern, Abbruch oder Schließen.
    """
    print("[V3] anchor calibration started", flush=True)
    img_path = screenshot_path.resolve()
    print(f"[V3] anchor calibration: screenshot path={img_path}", flush=True)

    if not img_path.is_file():
        msg = f"Screenshot fehlt:\n{img_path}"
        print(f"[V3][ERROR] anchor calibration failed: {msg}", flush=True)
        try:
            messagebox.showerror("Kalibrierung", msg, parent=root)
        except tk.TclError:
            pass
        return False

    print("[V3] anchor calibration: loading screenshot", flush=True)
    try:
        pil_img = Image.open(img_path).convert("RGB")
    except OSError as e:
        traceback.print_exc()
        print(f"[V3][ERROR] anchor calibration failed: {e}", flush=True)
        try:
            messagebox.showerror("Kalibrierung", f"Bild konnte nicht geöffnet werden:\n{e}", parent=root)
        except tk.TclError:
            pass
        return False

    iw, ih = pil_img.size
    print(f"[V3] anchor calibration: screenshot loaded size={iw}x{ih}", flush=True)

    sw = max(800, root.winfo_screenwidth())
    sh = max(600, root.winfo_screenheight())
    room_w = max(320, int(sw * 0.90) - 80)
    room_h = max(240, int(sh * 0.85) - 170)
    scale = min(1.0, room_w / max(iw, 1), room_h / max(ih, 1))
    dw, dh = int(iw * scale), int(ih * scale)

    result: dict[str, bool] = {"saved": False, "closing": False}

    print("[V3] anchor calibration: creating Tk window", flush=True)

    dlg = tk.Toplevel(root)
    dlg.title("Tablemap Scanner V3 — Anker setzen")
    dlg.transient(root)

    bx = tk.Frame(dlg)
    bx.pack(fill=tk.X, padx=10, pady=8)
    tk.Label(
        bx,
        text=(
            "Rotes Rechteck verschieben (innen ziehen) und skalieren (Ecke rechts unten ziehen).\n"
            "„Anker speichern“ schreibt den Ausschnitt aus der Original-PNG — ohne Overlay."
        ),
        justify=tk.LEFT,
    ).pack(anchor="w")
    tk.Label(
        bx,
        text="Nullpunkt: oben links (grünes Kreuz + 0,0)   |   Skalieren: blauer Griff rechts unten",
        justify=tk.LEFT,
        fg="#206020",
    ).pack(anchor="w", pady=(4, 0))

    cw = tk.Frame(dlg)
    cw.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
    vsb = tk.Scrollbar(cw, orient="vertical")
    hsb = tk.Scrollbar(cw, orient="horizontal")
    canvas_side = max(260, min(dw + 48, room_w))
    canvas_below = max(220, min(dh + 48, room_h))
    canvas = tk.Canvas(cw, width=canvas_side, height=canvas_below)
    canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set, highlightthickness=0)
    hsb.config(command=canvas.xview)
    vsb.config(command=canvas.yview)
    hsb.pack(side=tk.BOTTOM, fill=tk.X)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    _rs = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
    resized = pil_img.resize((dw, dh), _rs)
    photo = ImageTk.PhotoImage(resized, master=dlg)
    canvas.create_image(0, 0, anchor=tk.NW, image=photo)
    dlg.photo_ref = photo  # noqa: SLF001
    canvas.config(scrollregion=(0, 0, dw, dh))

    def disp_to_ix(xc: float, yc: float) -> tuple[int, int]:
        return max(0, min(iw - 1, int(xc / scale))), max(0, min(ih - 1, int(yc / scale)))

    def ix_to_canvas(ix: float, iy: float) -> tuple[float, float]:
        return ix * scale, iy * scale

    def clamp_rect(ix: float, iy: float, iw_r: float, ih_r: float) -> tuple[int, int, int, int]:
        w_, h_ = max(16, min(int(iw_r), iw)), max(16, min(int(ih_r), ih))
        x_, y_ = int(ix), int(iy)
        if x_ + w_ > iw:
            x_ = iw - w_
        if y_ + h_ > ih:
            y_ = ih - h_
        x_, y_ = max(0, x_), max(0, y_)
        return x_, y_, w_, h_

    rw_des = min(120, max(48, iw - 16))
    rh_des = min(80, max(48, ih - 16))
    ix0 = max(0, min(iw // 10, iw - rw_des))
    iy0 = max(0, min(ih // 10, ih - rh_des))
    ix0, iy0, rw, rh = clamp_rect(ix0, iy0, rw_des, rh_des)
    print(
        f"[V3] anchor calibration: rectangle initial x={ix0} y={iy0} w={rw} h={rh}",
        flush=True,
    )

    x1, y1 = ix_to_canvas(ix0, iy0)
    rect_id = canvas.create_rectangle(
        x1, y1, x1 + rw * scale, y1 + rh * scale, outline="#ff2020", width=4
    )

    HANDLE = 18
    br_id = canvas.create_rectangle(0, 0, HANDLE, HANDLE, outline="#e8f0ff", fill="#2a5ad6", width=2)
    br_diag1 = canvas.create_line(0, 0, 0, 0, fill="#ffffff", width=2)
    br_diag2 = canvas.create_line(0, 0, 0, 0, fill="#ffffff", width=2)

    ORIGIN_ARM = max(6, int(8 * scale))
    origin_cross_h = canvas.create_line(0, 0, 0, 0, fill="#00c853", width=3)
    origin_cross_v = canvas.create_line(0, 0, 0, 0, fill="#00c853", width=3)
    origin_ring = canvas.create_oval(0, 0, 0, 0, outline="#004d1a", width=2, fill="#b9f6ca")
    origin_label = canvas.create_text(
        0, 0, text="0,0", anchor=tk.NW, fill="#1b5e20", font=("Segoe UI", 9, "bold")
    )

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
        pr = max(3.0, 4.0 * scale)
        canvas.coords(origin_ring, bx1 - pr, by1 - pr, bx1 + pr, by1 + pr)
        canvas.coords(origin_label, bx1 + ORIGIN_ARM + 4, by1 + 2)

    state: dict[str, float | int | str] = {
        "mode": "idle",
        "dx": 0.0,
        "dy": 0.0,
        "corner_xoff": HANDLE / 2.0,
        "corner_yoff": HANDLE / 2.0,
        "fix_ix": 0,
        "fix_iy": 0,
        "w": rw,
        "h": rh,
    }

    def canvas_rect_to_ixy() -> tuple[int, int, int, int]:
        bx1, by1, bx2, by2 = canvas.coords(rect_id)
        xa, ya = disp_to_ix(bx1, by1)
        xb, yb = disp_to_ix(bx2, by2)
        if xb <= xa:
            xb = xa + 16
        if yb <= ya:
            yb = ya + 16
        return clamp_rect(xa, ya, xb - xa, yb - ya)

    def apply_ix(ix: int, iy: int, w_: int, h_: int) -> None:
        cx1_, cy1_ = ix_to_canvas(ix, iy)
        cx2_, cy2_ = ix_to_canvas(ix + w_, iy + h_)
        canvas.coords(rect_id, cx1_, cy1_, cx2_, cy2_)
        canvas.itemconfig(rect_id, outline="#ff2020", width=4)
        sync_handle()
        sync_origin_marker()

    def on_canvas_down(ev: tk.Event) -> None:
        cx, cy = canvas.canvasx(ev.x), canvas.canvasy(ev.y)
        bx1, by1, bx2, by2 = canvas.coords(rect_id)
        hx1, hy1, hx2, hy2 = canvas.coords(br_id)
        if hx1 <= cx <= hx2 and hy1 <= cy <= hy2:
            ix_q, iy_q, _, _ = canvas_rect_to_ixy()
            state["mode"] = "resize"
            state["corner_xoff"] = float(cx - bx2)
            state["corner_yoff"] = float(cy - by2)
            state["fix_ix"] = ix_q
            state["fix_iy"] = iy_q
        elif bx1 <= cx <= bx2 and by1 <= cy <= by2:
            _, _, ww_, hh_ = canvas_rect_to_ixy()
            state["mode"] = "move"
            state["dx"] = cx - bx1
            state["dy"] = cy - by1
            state["w"] = ww_
            state["h"] = hh_
        else:
            state["mode"] = "idle"

    def on_canvas_motion(ev: tk.Event) -> None:
        cx, cy = canvas.canvasx(ev.x), canvas.canvasy(ev.y)
        m = str(state["mode"])
        if m == "move":
            nx1_canvas = cx - float(state["dx"])
            ny1_canvas = cy - float(state["dy"])
            nix, niy = disp_to_ix(nx1_canvas, ny1_canvas)
            ni, nj, nk, nh = clamp_rect(
                float(nix), float(niy), float(int(state["w"])), float(int(state["h"]))
            )
            apply_ix(ni, nj, nk, nh)
        elif m == "resize":
            target_x2 = cx - float(state["corner_xoff"])
            target_y2 = cy - float(state["corner_yoff"])
            br_x, br_y = disp_to_ix(target_x2, target_y2)
            fi, fj = int(state["fix_ix"]), int(state["fix_iy"])
            ni, nj, nk, nh = clamp_rect(float(fi), float(fj), float(br_x - fi), float(br_y - fj))
            apply_ix(ni, nj, nk, nh)

    def on_canvas_release(_ev: tk.Event) -> None:
        state["mode"] = "idle"

    canvas.bind("<ButtonPress-1>", on_canvas_down)
    canvas.bind("<B1-Motion>", on_canvas_motion)
    canvas.bind("<ButtonRelease-1>", on_canvas_release)
    sync_handle()
    sync_origin_marker()

    bf = tk.Frame(dlg)
    bf.pack(fill=tk.X, pady=(8, 10))

    def close_cancel() -> None:
        if result["closing"]:
            return
        result["closing"] = True
        print("[V3] anchor calibration cancelled", flush=True)
        _safe_grab_release(dlg)
        try:
            dlg.destroy()
        except tk.TclError:
            pass

    def save_anchor() -> None:
        ix, iy_, w_, h_ = canvas_rect_to_ixy()
        if w_ < 8 or h_ < 8:
            messagebox.showwarning(
                "Kalibrierung", "Ankerbereich zu klein (min. ~8 Pixel).", parent=dlg
            )
            return
        canvas.itemconfigure(rect_id, state="hidden")
        canvas.itemconfigure(br_id, state="hidden")
        canvas.itemconfigure(br_diag1, state="hidden")
        canvas.itemconfigure(br_diag2, state="hidden")
        canvas.itemconfigure(origin_cross_h, state="hidden")
        canvas.itemconfigure(origin_cross_v, state="hidden")
        canvas.itemconfigure(origin_ring, state="hidden")
        canvas.itemconfigure(origin_label, state="hidden")
        dlg.update_idletasks()
        dlg.update()

        try:
            left, top, right, bottom = ix, iy_, ix + w_, iy_ + h_
            crop = pil_img.crop((left, top, right, bottom))
            ANCHORS_DIR.mkdir(parents=True, exist_ok=True)
            crop.save(str(ANCHOR_PATCH_PATH), format="PNG", optimize=True)

            cfg = {
                "schema": ANCHOR_SCHEMA,
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "anchor_name": "table_anchor",
                "anchor_width": w_,
                "anchor_height": h_,
                "origin": "top_left",
                "origin_pixel": {"x": 0, "y": 0},
                "origin_screen_hint": "top_left_of_rectangle",
                "resize_handle": "bottom_right",
                "source_screenshot": str(img_path),
                "note": "Nullpunkt ist oben links im gespeicherten Anker-Patch.",
            }
            ANCHOR_CONFIG_PATH.write_text(
                json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"[V3] anchor origin: top_left x={left} y={top}", flush=True)
            print(
                f"[V3] anchor crop: left={left} top={top} right={right} bottom={bottom}",
                flush=True,
            )
            print(f"[V3] anchor saved: {ANCHOR_PATCH_PATH}", flush=True)
            print(f"[V3] anchor config saved: {ANCHOR_CONFIG_PATH}", flush=True)
        except OSError as ex:
            traceback.print_exc()
            print(f"[V3][ERROR] anchor calibration failed: {ex}", flush=True)
            try:
                messagebox.showerror(
                    "Kalibrierung",
                    f"Konnte nicht speichern:\n{ex}",
                    parent=dlg,
                )
            except tk.TclError:
                pass
            return

        result["saved"] = True
        result["closing"] = True
        _safe_grab_release(dlg)
        try:
            dlg.destroy()
        except tk.TclError:
            pass

    tk.Button(bf, text="Abbrechen", command=close_cancel).pack(side=tk.RIGHT, padx=10)
    tk.Button(bf, text="Anker speichern", command=save_anchor).pack(side=tk.RIGHT, padx=10)

    def on_wm_close() -> None:
        print("[V3] anchor calibration window closed", flush=True)
        close_cancel()

    dlg.protocol("WM_DELETE_WINDOW", on_wm_close)

    try:
        dlg.grab_set()
    except tk.TclError:
        pass

    root.lift()
    dlg.update_idletasks()
    dlg.update()
    ww = dlg.winfo_reqwidth()
    wh = dlg.winfo_reqheight()
    gx = max(0, (sw - ww) // 2)
    gy = max(0, (sh - wh) // 6)
    dlg.geometry(f"{ww}x{wh}+{gx}+{gy}")
    dlg.lift()
    dlg.attributes("-topmost", True)

    def _dlg_top_off() -> None:
        try:
            dlg.attributes("-topmost", False)
        except tk.TclError:
            pass

    dlg.after(400, _dlg_top_off)

    try:
        dlg.focus_force()
    except Exception:
        pass

    print("[V3] anchor calibration: Tk window visible", flush=True)
    if modal:
        print("[V3] anchor calibration: entering wait_window", flush=True)
        dlg.wait_window()
        print("[V3] anchor calibration: wait_window returned", flush=True)

    return bool(result.get("saved"))
