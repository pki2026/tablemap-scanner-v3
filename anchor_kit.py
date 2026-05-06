"""
Anker: vier dünne topmost Randfenster + 0,0 + Skaliergriff — Desktop bleibt außerhalb bedienbar.
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

THICK = 4
HANDLE_SZ = 20
ORIGIN_SZ = 40


def _safe_grab_release(w: tk.Misc) -> None:
    try:
        w.grab_release()
    except tk.TclError:
        pass


def _destroy_windows(ws: list[tk.Misc]) -> None:
    for w in ws:
        try:
            _safe_grab_release(w)
        except Exception:
            pass
        try:
            w.destroy()
        except tk.TclError:
            pass


def run_anchor_calibration_blocking(root: tk.Tk, *, modal: bool = True) -> bool:
    print("[V3] anchor calibration: border frame (non-blocking)", flush=True)

    result: dict[str, bool] = {"saved": False, "closing": False}

    try:
        sw_i = int(root.winfo_screenwidth())
        sh_i = int(root.winfo_screenheight())
        print(f"[V3] anchor frame: screen {sw_i}x{sh_i}", flush=True)

        fw0 = min(480, sw_i - 120)
        fh0 = min(360, sh_i - 160)
        frame = {
            "left": max(40, (sw_i - fw0) // 2),
            "top": max(40, (sh_i - fh0) // 2),
            "w": fw0,
            "h": fh0,
        }

        interaction: dict[str, int | str | float] = {"mode": "idle"}

        def make_red_strip() -> tk.Toplevel:
            w = tk.Toplevel(root)
            w.overrideredirect(True)
            w.attributes("-topmost", True)
            fr = tk.Frame(w, bg="#ff2020", bd=0, highlightthickness=0)
            fr.pack(fill=tk.BOTH, expand=True)
            return w

        win_top = make_red_strip()
        win_bottom = make_red_strip()
        win_left = make_red_strip()
        win_right = make_red_strip()

        win_origin = tk.Toplevel(root)
        win_origin.overrideredirect(True)
        win_origin.attributes("-topmost", True)
        win_origin.configure(bg="#0d1a12")
        tk.Label(
            win_origin,
            text="0,0",
            bg="#0d1a12",
            fg="#00ff88",
            font=("Segoe UI", 9, "bold"),
        ).pack(fill=tk.BOTH, expand=True)

        win_handle = tk.Toplevel(root)
        win_handle.overrideredirect(True)
        win_handle.attributes("-topmost", True)
        hfr = tk.Frame(win_handle, bg="#2a5ad6", bd=1, highlightthickness=1, highlightbackground="#ffffff")
        hfr.pack(fill=tk.BOTH, expand=True)

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

        all_wins: list[tk.Toplevel] = [
            win_top,
            win_bottom,
            win_left,
            win_right,
            win_origin,
            win_handle,
            toolbar,
        ]

        def sync_geometry() -> None:
            fl = int(frame["left"])
            ft = int(frame["top"])
            fw = int(frame["w"])
            fh = int(frame["h"])
            win_top.geometry(f"{fw}x{THICK}+{fl}+{ft}")
            win_bottom.geometry(f"{fw}x{THICK}+{fl}+{ft + fh - THICK}")
            win_left.geometry(f"{THICK}x{fh}+{fl}+{ft}")
            win_right.geometry(f"{THICK}x{fh}+{fl + fw - THICK}+{ft}")

            ox = max(0, fl - 6)
            oy = max(0, ft - 6)
            win_origin.geometry(f"{ORIGIN_SZ}x{ORIGIN_SZ}+{ox}+{oy}")
            win_handle.geometry(f"{HANDLE_SZ}x{HANDLE_SZ}+{fl + fw - HANDLE_SZ}+{ft + fh - HANDLE_SZ}")

            for bw in (win_top, win_bottom, win_left, win_right, win_origin, win_handle):
                try:
                    bw.lift()
                except tk.TclError:
                    pass
            try:
                toolbar.lift()
            except tk.TclError:
                pass

        def on_border_press(ev: tk.Event) -> None:
            if str(interaction["mode"]) != "idle":
                return
            interaction["mode"] = "move"
            interaction["x0"] = int(ev.x_root)
            interaction["y0"] = int(ev.y_root)
            interaction["fl0"] = int(frame["left"])
            interaction["ft0"] = int(frame["top"])

        def on_border_motion(ev: tk.Event) -> None:
            if str(interaction["mode"]) != "move":
                return
            x0 = int(interaction["x0"])
            y0 = int(interaction["y0"])
            nl = int(interaction["fl0"]) + int(ev.x_root) - x0
            nt = int(interaction["ft0"]) + int(ev.y_root) - y0
            fw = int(frame["w"])
            fh = int(frame["h"])
            nl = max(0, min(nl, sw_i - fw))
            nt = max(0, min(nt, sh_i - fh))
            frame["left"] = nl
            frame["top"] = nt
            sync_geometry()

        def on_handle_press(ev: tk.Event) -> None:
            if str(interaction["mode"]) != "idle":
                return
            interaction["mode"] = "resize"
            interaction["x0"] = int(ev.x_root)
            interaction["y0"] = int(ev.y_root)
            interaction["w0"] = int(frame["w"])
            interaction["h0"] = int(frame["h"])

        def on_handle_motion(ev: tk.Event) -> None:
            if str(interaction["mode"]) != "resize":
                return
            nw = max(32, int(interaction["w0"]) + int(ev.x_root) - int(interaction["x0"]))
            nh = max(32, int(interaction["h0"]) + int(ev.y_root) - int(interaction["y0"]))
            fl = int(frame["left"])
            ft = int(frame["top"])
            nw = min(nw, sw_i - fl)
            nh = min(nh, sh_i - ft)
            frame["w"] = nw
            frame["h"] = nh
            sync_geometry()

        def on_any_release(_ev: tk.Event | None = None) -> None:
            interaction["mode"] = "idle"

        for w in (win_top, win_bottom, win_left, win_right, win_origin):
            w.bind("<ButtonPress-1>", on_border_press)
            w.bind("<B1-Motion>", on_border_motion)
            w.bind("<ButtonRelease-1>", on_any_release)

        win_handle.bind("<ButtonPress-1>", on_handle_press)
        win_handle.bind("<B1-Motion>", on_handle_motion)
        win_handle.bind("<ButtonRelease-1>", on_any_release)

        def close_cancel() -> None:
            if result["closing"]:
                return
            result["closing"] = True
            print("[V3] anchor calibration: cancelled", flush=True)
            _destroy_windows(all_wins)

        def save_anchor() -> None:
            fl, ft, fw, fh = (
                int(frame["left"]),
                int(frame["top"]),
                int(frame["w"]),
                int(frame["h"]),
            )
            if fw < 8 or fh < 8:
                messagebox.showwarning("Anker", "Bereich zu klein.", parent=toolbar)
                return
            left, top = fl, ft
            right_ex = left + fw
            bottom_ex = top + fh
            x1, y1, x2, y2 = left, top, right_ex - 1, bottom_ex - 1
            print(
                f"[V3] anchor calibration: grab screen bbox left={left} top={top} right_ex={right_ex} bottom_ex={bottom_ex}",
                flush=True,
            )
            try:
                for w in all_wins:
                    try:
                        w.withdraw()
                    except tk.TclError:
                        pass
                root.update_idletasks()
                root.update()
                time.sleep(0.15)
                try:
                    im = ImageGrab.grab(
                        bbox=(left, top, right_ex, bottom_ex), all_screens=True
                    )
                except TypeError:
                    im = ImageGrab.grab(bbox=(left, top, right_ex, bottom_ex))
            except Exception as ex:
                traceback.print_exc()
                print(f"[V3][ERROR] anchor grab failed: {ex}", flush=True)
                for w in all_wins:
                    try:
                        w.deiconify()
                    except tk.TclError:
                        pass
                sync_geometry()
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
                    "anchor_width": fw,
                    "anchor_height": fh,
                    "capture_rect_screen": {"x1": x1, "y1": y1, "x2": x2, "y2": y2},
                    "calibration_mode": "border_frame",
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
                for w in all_wins:
                    try:
                        w.deiconify()
                    except tk.TclError:
                        pass
                sync_geometry()
                return

            result["saved"] = True
            result["closing"] = True
            _destroy_windows(all_wins)

        tk.Button(btn_row, text="Abbrechen", command=close_cancel, width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_row, text="Anker speichern", command=save_anchor, width=16).pack(side=tk.LEFT, padx=6)

        toolbar.protocol("WM_DELETE_WINDOW", close_cancel)

        sync_geometry()
        root.update_idletasks()
        root.update()
        for w in all_wins:
            try:
                w.deiconify()
            except tk.TclError:
                pass
        sync_geometry()
        print("[V3] anchor frame: visible (borders + toolbar)", flush=True)

        if modal:
            print("[V3] anchor frame: wait_window(toolbar)", flush=True)
            toolbar.wait_window()

        print("[V3] anchor frame: done", flush=True)
        return bool(result.get("saved"))

    except Exception as exc:
        traceback.print_exc()
        print(f"[V3][ERROR] overlay failed: {exc}", flush=True, file=sys.stderr)
        return False
