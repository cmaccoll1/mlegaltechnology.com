#!/usr/bin/env python3
"""
MacColl Transcript Review Tool
================================
Continuous-scroll deposition transcript review with side-by-side
exhibit viewer, flagged-page sidebar, and free-form notes.
Pages load lazily as you scroll — no page-by-page flipping.

DEPENDENCIES (one-time install):
    python -m pip install pymupdf Pillow

HOW TO RUN:
    python Depo_Review.py
"""

import os, sys, re, csv, json, tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from pathlib import Path
from datetime import datetime
from collections import OrderedDict

try:
    import fitz
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install pymupdf")
try:
    from PIL import Image, ImageTk, ImageDraw
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install Pillow")

# ── Palette ────────────────────────────────────────────────────────────────────
C = dict(
    bg      = "#0d1b2e",
    bg2     = "#121f33",
    bg3     = "#1a2a42",
    bg4     = "#243450",
    bg5     = "#2d3f5c",
    fg      = "#e8e4da",
    fg2     = "#b8b2a4",
    fg3     = "#6b7885",
    gold    = "#c9a84c",
    gold2   = "#e4c77a",
    gold3   = "#8a6f2e",
    red     = "#e05c5c",
    red2    = "#8b1a1a",
    green   = "#4caf7d",
    blue    = "#5096d6",
    sash    = "#253050",
    page_bg = "#1e2e46",
    page_bd = "#2d4060",
)

NOTES_SUFFIX = "_mcrt_notes.json"
PAGE_GAP     = 10
RENDER_BUF   = 2        # extra pages above/below viewport to keep rendered
MAX_CACHE    = 30       # max cached rendered pages per viewer


# =============================================================================
# CONTINUOUS-SCROLL DOCUMENT VIEWER
# =============================================================================
class DocViewer:
    """
    Renders a whole PDF (or image file) into one tall scrollable canvas.
    Only pages near the viewport are rendered; distant ones are evicted
    from the LRU cache and replaced with lightweight placeholders.

    Public API
    ----------
    load(path)             -> bool
    close()                -> None
    go_to(n)               -> None   scroll so page n is at top
    prev() / next()        -> None   scroll to adjacent page
    search(query)          -> list   [(page_no, snippet)]
    get_page_text(page_no) -> str
    notify_scroll()        -> None   call after external scroll events
    current_page           -> int
    total_pages            -> int
    """

    def __init__(self, canvas, page_lbl, prev_btn, next_btn,
                 on_page_change=None):
        self.canvas          = canvas
        self.page_lbl        = page_lbl
        self.prev_btn        = prev_btn
        self.next_btn        = next_btn
        self._on_page_change = on_page_change

        self._doc    = None
        self._mode   = None      # "pdf" | "image"
        self._total  = 0
        self._cw     = 0

        # Layout
        self._page_y  = []
        self._page_h  = []
        self._total_h = 0

        # Rendering
        self._cache    = OrderedDict()   # pno -> PhotoImage  (LRU)
        self._items    = {}              # pno -> canvas image item id
        self._pholds   = {}              # pno -> canvas rect item id
        self._current  = 0
        self._hits     = {}              # pno -> [fitz.Rect]

        # Zoom  (1.0 = fit-to-width; >1 = larger; <1 = smaller)
        self._zoom      = 1.0
        self._zoom_step = 0.25
        self._zoom_min  = 0.25
        self._zoom_max  = 4.0
        self._zoom_lbl  = None           # optional tk.Label to keep in sync
        self._h_scrollbar = None         # optional horizontal ttk.Scrollbar

        # Highlights  {pno: [(nx0,ny0,nx1,ny1,hex_color), ...]}  normalised coords
        self._highlights  = {}
        self._hl_mode     = False
        self._hl_color    = "#ffd700"    # current highlight colour
        self._hl_start    = None         # (cx, cy) in canvas coords at press
        self._hl_start_p  = None         # page index of press point
        self._hl_rubber   = None         # rubber-band canvas item id

        # Text selection state
        self._sel_mode   = False
        self._sel_start  = None        # (cx, cy) in canvas coords
        self._sel_page   = None        # page where selection started
        self._sel_items  = []          # canvas rect item ids (overlay)
        self._sel_text   = ""          # currently selected plain text
        self._word_cache = {}          # pno -> [(x0,y0,x1,y1,word), ...]

        self.canvas.bind("<Configure>",     self._on_configure)
        self.canvas.bind("<ButtonPress-1>", self._hl_press)
        self.canvas.bind("<B1-Motion>",     self._hl_drag)
        self.canvas.bind("<ButtonRelease-1>", self._hl_release)

    @property
    def current_page(self): return self._current
    @property
    def total_pages(self):  return self._total

    # ── load / close ───────────────────────────────────────────────────────────
    def load(self, path):
        self._teardown()
        suf = path.suffix.lower()
        try:
            if suf == ".pdf":
                self._doc   = fitz.open(str(path))
                self._mode  = "pdf"
                self._total = len(self._doc)
            elif suf in (".jpg",".jpeg",".png",".tif",".tiff",
                         ".bmp",".gif",".webp"):
                raw = Image.open(str(path))
                frames = []
                try:
                    for i in range(getattr(raw, "n_frames", 1)):
                        raw.seek(i); frames.append(raw.copy().convert("RGB"))
                except Exception:
                    frames = [raw.convert("RGB")]
                self._doc   = frames
                self._mode  = "image"
                self._total = len(frames)
            else:
                return False
            self._current    = 0
            self._hits       = {}
            self._highlights = {}
            self._hl_mode    = False
            self._sel_mode   = False
            self._sel_start  = None
            self._sel_page   = None
            self._sel_items  = []
            self._sel_text   = ""
            self._word_cache = {}
            self._build_layout()
            self._render_visible()
            self._update_nav()
            return True
        except Exception as exc:
            self._show_msg(f"Cannot open file:\n{exc}")
            return False

    def close(self):
        self._teardown()
        self.canvas.delete("all")
        self._update_nav()

    # ── navigation ─────────────────────────────────────────────────────────────
    def go_to(self, n):
        if not self._page_y: return
        n = max(0, min(n, self._total - 1))
        if self._total_h == 0: return
        self.canvas.yview_moveto(self._page_y[n] / self._total_h)
        self._render_visible()
        self._update_nav()

    def prev(self): self.go_to(self._current - 1)
    def next(self): self.go_to(self._current + 1)

    def notify_scroll(self):
        self._render_visible()
        self._update_nav()

    # ── search ─────────────────────────────────────────────────────────────────
    def search(self, query):
        self._hits = {}
        results = []
        if not query or not self._doc or self._mode != "pdf":
            self._invalidate_all(); self._render_visible(); return results
        for pno in range(self._total):
            rects = self._doc[pno].search_for(query)
            if rects:
                self._hits[pno] = rects
                txt = self._doc[pno].get_text()
                q   = query.lower()
                idx = txt.lower().find(q)
                snip = (txt[max(0,idx-30):idx+len(query)+40]
                        .replace("\n"," ").strip()) if idx>=0 else ""
                results.append((pno, snip))
        self._invalidate_all()
        self._render_visible()
        return results

    def get_page_text(self, pno):
        if not self._doc or self._mode != "pdf": return ""
        try:    return self._doc[pno].get_text()
        except: return ""

    # ── layout ─────────────────────────────────────────────────────────────────
    # ── zoom helpers ───────────────────────────────────────────────────────────
    def _render_width(self):
        """Pixel width of a rendered page at current zoom."""
        return max(100, int((self._cw - 8) * self._zoom))

    def zoom_in(self):
        self._zoom = min(self._zoom_max,
                         round(self._zoom + self._zoom_step, 2))
        self._apply_zoom()

    def zoom_out(self):
        self._zoom = max(self._zoom_min,
                         round(self._zoom - self._zoom_step, 2))
        self._apply_zoom()

    def _apply_zoom(self):
        if self._zoom_lbl:
            self._zoom_lbl.config(text=f"{int(self._zoom*100)}%")
        if not self._doc: return
        old = self._current
        self._invalidate_all()
        self._build_layout()
        self.go_to(old)
        self._update_h_scroll()

    def _update_h_scroll(self):
        """Connect horizontal scrollbar when zoomed past fit-width."""
        if not self._h_scrollbar: return
        rw = self._render_width()
        cw = self._cw
        if rw > cw:
            # Scrollable: compute fraction
            self._h_scrollbar.config(state="normal")
        else:
            self._h_scrollbar.set(0, 1)

    # ── highlight helpers ───────────────────────────────────────────────────────
    def set_hl_mode(self, on: bool):
        if on:
            self._sel_mode = False
            self._clear_selection()
        self._hl_mode = on
        self.canvas.config(cursor="crosshair" if on else "")

    def set_sel_mode(self, on: bool):
        if on:
            self._hl_mode = False
        else:
            self._clear_selection()
        self._sel_mode = on
        self.canvas.config(cursor="xterm" if on else "")

    def get_selected_text(self):
        return self._sel_text

    def set_hl_color(self, color: str):
        self._hl_color = color

    def undo_last_highlight(self):
        """Remove the most recently added highlight across all pages."""
        # Find the page with highlights and pop the last one
        for pno in sorted(self._highlights, reverse=True):
            if self._highlights[pno]:
                self._highlights[pno].pop()
                if not self._highlights[pno]:
                    del self._highlights[pno]
                self._invalidate_page(pno)
                return

    def clear_highlights(self):
        pages = list(self._highlights.keys())
        self._highlights.clear()
        for pno in pages:
            self._invalidate_page(pno)

    def get_highlights(self):
        return {str(k): v for k, v in self._highlights.items()}

    def load_highlights(self, data: dict):
        self._highlights = {int(k): v for k, v in data.items()}
        self._invalidate_all()
        self._render_visible()

    def _invalidate_page(self, pno: int):
        """Drop one cached page so it re-renders on next scroll."""
        if pno in self._cache:
            del self._cache[pno]
        if pno in self._items:
            self.canvas.delete(self._items.pop(pno))
        if pno in self._pholds:
            self.canvas.itemconfig(self._pholds[pno], state="normal")
            if self._page_y:
                self.canvas.create_text(
                    self._cw//2,
                    self._page_y[pno]+max(20, self._page_h[pno]//2),
                    text=f"Page {pno+1}", fill=C["bg5"],
                    font=("Georgia",10), tags=f"plbl_{pno}")
        self._render_visible()

    # ── highlight mouse events ──────────────────────────────────────────────────
    def _hl_press(self, event):
        if self._sel_mode:
            self._sel_press(event); return
        if not self._hl_mode: return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        self._hl_start   = (cx, cy)
        self._hl_start_p = self._canvas_to_page(cx, cy)

    def _hl_drag(self, event):
        if self._sel_mode:
            self._sel_drag(event); return
        if not self._hl_mode or self._hl_start is None: return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        x0, y0 = self._hl_start
        if self._hl_rubber:
            self.canvas.coords(self._hl_rubber, x0, y0, cx, cy)
        else:
            self._hl_rubber = self.canvas.create_rectangle(
                x0, y0, cx, cy,
                outline=self._hl_color, width=2, fill="", dash=(4, 2))

    def _hl_release(self, event):
        if self._sel_mode:
            self._sel_release(event); return
        if not self._hl_mode or self._hl_start is None: return
        cx1 = self.canvas.canvasx(event.x)
        cy1 = self.canvas.canvasy(event.y)
        cx0, cy0 = self._hl_start

        # Remove rubber-band
        if self._hl_rubber:
            self.canvas.delete(self._hl_rubber)
            self._hl_rubber = None

        self._hl_start = None

        # Reject tiny drags (mis-clicks)
        if abs(cx1-cx0) < 4 or abs(cy1-cy0) < 4:
            return

        pno = self._hl_start_p
        if pno is None: return

        # Convert both corners to normalised page coords
        rw = self._render_width()
        py = self._page_y[pno]
        ph = self._page_h[pno]
        pw = rw   # rendered pixel width of this page

        def _norm_xy(cx, cy):
            nx = max(0.0, min(1.0, (cx - 4) / pw))
            ny = max(0.0, min(1.0, (cy - py) / ph))
            return nx, ny

        nx0, ny0 = _norm_xy(cx0, cy0)
        nx1, ny1 = _norm_xy(cx1, cy1)
        # Ensure (nx0,ny0) is top-left
        nx0, nx1 = sorted([nx0, nx1])
        ny0, ny1 = sorted([ny0, ny1])

        if pno not in self._highlights:
            self._highlights[pno] = []
        self._highlights[pno].append((nx0, ny0, nx1, ny1, self._hl_color))
        self._invalidate_page(pno)

    # ── text selection ─────────────────────────────────────────────────────
    def _sel_press(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        self._clear_selection()
        self._sel_start = (cx, cy)
        self._sel_page  = self._canvas_to_page(cx, cy)
        self.canvas.focus_set()

    def _sel_drag(self, event):
        if self._sel_start is None: return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        cx0, cy0 = self._sel_start
        # rubber-band
        if self._hl_rubber:
            self.canvas.coords(self._hl_rubber, cx0, cy0, cx, cy)
        else:
            self._hl_rubber = self.canvas.create_rectangle(
                cx0, cy0, cx, cy,
                outline=C["blue"], width=1, fill="", dash=(2, 2))
        self._update_sel_words(cx0, cy0, cx, cy)

    def _sel_release(self, event):
        if self._sel_start is None: return
        cx1 = self.canvas.canvasx(event.x)
        cy1 = self.canvas.canvasy(event.y)
        cx0, cy0 = self._sel_start
        if self._hl_rubber:
            self.canvas.delete(self._hl_rubber)
            self._hl_rubber = None
        self._sel_start = None
        self._update_sel_words(cx0, cy0, cx1, cy1)

    def _update_sel_words(self, cx0, cy0, cx1, cy1):
        """Recompute selection: find words in drag rect, draw overlays."""
        for item in self._sel_items:
            self.canvas.delete(item)
        self._sel_items.clear()
        self._sel_text = ""

        pno = self._sel_page
        if pno is None or self._mode != "pdf": return

        rw    = self._render_width()
        pw    = self._doc[pno].rect.width
        ratio = rw / pw
        py    = self._page_y[pno]

        # Selection rect in PDF coords
        sx0 = (min(cx0, cx1) - 4) / ratio
        sy0 = (min(cy0, cy1) - py) / ratio
        sx1 = (max(cx0, cx1) - 4) / ratio
        sy1 = (max(cy0, cy1) - py) / ratio

        words = self._get_words(pno)
        selected = []
        for (wx0, wy0, wx1, wy1, word) in words:
            # axis-aligned intersection
            if wx0 < sx1 and wx1 > sx0 and wy0 < sy1 and wy1 > sy0:
                selected.append((wy0, wx0, wx1, wy1, word))

        # Sort reading order: top-to-bottom then left-to-right
        selected.sort(key=lambda w: (round(w[0], 1), w[1]))

        prev_y = None
        parts  = []
        for (wy0, wx0, wx1, wy1, word) in selected:
            # Draw blue overlay on canvas
            ccx0 = int(wx0 * ratio) + 4
            ccy0 = int(wy0 * ratio) + py
            ccx1 = int(wx1 * ratio) + 4
            ccy1 = int(wy1 * ratio) + py
            item = self.canvas.create_rectangle(
                ccx0, ccy0, ccx1, ccy1,
                fill=C["blue"], outline="", stipple="gray50")
            self._sel_items.append(item)
            # Accumulate text
            if prev_y is not None and abs(wy0 - prev_y) > 2:
                parts.append("\n")
            elif parts:
                parts.append(" ")
            parts.append(word)
            prev_y = wy0

        self._sel_text = "".join(parts)

    def _get_words(self, pno):
        """Return cached word list [(x0,y0,x1,y1,word)] in PDF coords."""
        if pno not in self._word_cache and self._mode == "pdf":
            try:
                raw = self._doc[pno].get_text("words")
                self._word_cache[pno] = [
                    (w[0], w[1], w[2], w[3], w[4]) for w in raw]
            except Exception:
                self._word_cache[pno] = []
        return self._word_cache.get(pno, [])

    def _clear_selection(self):
        for item in self._sel_items:
            self.canvas.delete(item)
        self._sel_items.clear()
        self._sel_text = ""
        if self._hl_rubber:
            self.canvas.delete(self._hl_rubber)
            self._hl_rubber = None

    def _canvas_to_page(self, cx, cy):
        """Return page index that contains canvas point (cx, cy), or None."""
        for i in range(self._total):
            if self._page_y[i] <= cy <= self._page_y[i] + self._page_h[i]:
                return i
        return None

    def _build_layout(self):
        self.canvas.update_idletasks()
        self._cw = max(self.canvas.winfo_width(), 400)
        rw = self._render_width()
        self._page_y = []; self._page_h = []
        y = PAGE_GAP
        for i in range(self._total):
            h = self._nat_height(i)
            self._page_y.append(y); self._page_h.append(h)
            y += h + PAGE_GAP
        self._total_h = y
        # Canvas content width: at least cw, but wider when zoomed in
        content_w = max(self._cw, rw + 8)
        self.canvas.delete("all")
        self._cache.clear(); self._items.clear(); self._pholds.clear()
        for i in range(self._total):
            ph = self.canvas.create_rectangle(
                4, self._page_y[i], rw + 4,
                self._page_y[i]+self._page_h[i],
                fill=C["page_bg"], outline=C["page_bd"])
            self._pholds[i] = ph
            self.canvas.create_text(
                (rw + 8)//2,
                self._page_y[i] + max(20, self._page_h[i]//2),
                text=f"Page {i+1}", fill=C["bg5"],
                font=("Georgia", 10), tags=f"plbl_{i}")
        self.canvas.config(scrollregion=(0, 0, content_w, self._total_h))
        if self._h_scrollbar:
            cmd = self._h_scrollbar.cget("command")
            self.canvas.configure(xscrollcommand=self._h_scrollbar.set)

    def _nat_height(self, i):
        rw = self._render_width()
        if self._mode == "pdf":
            p = self._doc[i]
            return max(20, int(p.rect.height * rw / p.rect.width))
        img = self._doc[i]
        return max(20, int(img.height * rw / img.width))

    # ── rendering ──────────────────────────────────────────────────────────────
    def _visible_range(self):
        if not self._page_y: return 0, 0
        t, b  = self.canvas.yview()
        buf   = self._page_h[0] * RENDER_BUF if self._page_h else 200
        vt    = max(0, t*self._total_h - buf)
        vb    = min(self._total_h, b*self._total_h + buf)
        first = 0
        for i, y in enumerate(self._page_y):
            if y + self._page_h[i] >= vt: first = i; break
        last = self._total - 1
        for i in range(first, self._total):
            if self._page_y[i] > vb: last = max(first, i-1); break
        return first, last

    def _render_visible(self):
        if not self._doc or not self._page_y: return
        first, last = self._visible_range()
        for pno in range(first, last+1):
            if pno not in self._cache:
                self._render_one(pno)
        self._evict(first, last)
        self._update_current()

    def _render_one(self, pno):
        try:
            rw = self._render_width()
            if self._mode == "pdf":
                page  = self._doc[pno]
                ratio = rw / page.rect.width
                pix   = page.get_pixmap(
                    matrix=fitz.Matrix(ratio, ratio), alpha=False)
                img = Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
            else:
                src   = self._doc[pno]
                ratio = rw / src.width
                img   = src.resize(
                    (int(src.width*ratio),int(src.height*ratio)),
                    Image.LANCZOS)

            # Draw overlays: manual highlights first, search hits on top
            has_hl = pno in self._highlights and self._highlights[pno]
            has_hit= pno in self._hits and self._hits[pno]
            if has_hl or has_hit:
                draw = ImageDraw.Draw(img, "RGBA")
                for (nx0,ny0,nx1,ny1,hcol) in self._highlights.get(pno, []):
                    r = tuple(int(hcol.lstrip("#")[i:i+2], 16) for i in (0,2,4))
                    draw.rectangle(
                        [int(nx0*img.width), int(ny0*img.height),
                         int(nx1*img.width), int(ny1*img.height)],
                        fill=(*r, 110))
                if self._mode == "pdf":
                    for r in self._hits.get(pno, []):
                        draw.rectangle(
                            [int(r.x0*ratio), int(r.y0*ratio),
                             int(r.x1*ratio), int(r.y1*ratio)],
                            fill=(255, 230, 0, 120))

            photo = ImageTk.PhotoImage(img)
            self._cache[pno] = photo
            self._cache.move_to_end(pno)
            self.canvas.delete(f"plbl_{pno}")
            if pno in self._pholds:
                self.canvas.itemconfig(self._pholds[pno], state="hidden")
            if pno in self._items:
                self.canvas.itemconfig(self._items[pno], image=photo)
            else:
                item = self.canvas.create_image(
                    4, self._page_y[pno], anchor="nw", image=photo)
                self._items[pno] = item
        except Exception:
            pass

    def _evict(self, first, last):
        while len(self._cache) > MAX_CACHE:
            old, _ = self._cache.popitem(last=False)
            if old in self._items:
                self.canvas.delete(self._items.pop(old))
            if old in self._pholds:
                rw = self._render_width()
                self.canvas.itemconfig(self._pholds[old], state="normal")
                self.canvas.create_text(
                    (rw+8)//2,
                    self._page_y[old]+max(20,self._page_h[old]//2),
                    text=f"Page {old+1}", fill=C["bg5"],
                    font=("Georgia",10), tags=f"plbl_{old}")

    def _invalidate_all(self):
        for pno in list(self._items):
            self.canvas.delete(self._items.pop(pno))
        self._cache.clear()
        rw = self._render_width()
        for pno, ph in self._pholds.items():
            self.canvas.itemconfig(ph, state="normal")
            self.canvas.create_text(
                (rw+8)//2,
                self._page_y[pno]+max(20,self._page_h[pno]//2),
                text=f"Page {pno+1}", fill=C["bg5"],
                font=("Georgia",10), tags=f"plbl_{pno}")

    # ── page indicator ─────────────────────────────────────────────────────────
    def _update_current(self):
        if not self._page_y: return
        t, b   = self.canvas.yview()
        vp_mid = (t+b)/2 * self._total_h
        best, best_d = 0, float("inf")
        for i in range(self._total):
            d = abs(self._page_y[i]+self._page_h[i]/2 - vp_mid)
            if d < best_d: best_d, best = d, i
        if best != self._current:
            self._current = best
            if self._on_page_change:
                self._on_page_change(best)
        self._update_nav()

    def _update_nav(self):
        p, t = self._current, self._total
        self.page_lbl.config(text=f"{p+1} / {t}" if t else "— / —")
        self.prev_btn.config(state="normal" if p>0       else "disabled")
        self.next_btn.config(state="normal" if t and p<t-1 else "disabled")

    def _teardown(self):
        if self._doc and self._mode == "pdf":
            try: self._doc.close()
            except: pass
        self._doc=None; self._mode=None; self._total=0
        self._cache.clear(); self._items.clear(); self._pholds.clear()
        self._page_y=[]; self._page_h=[]; self._total_h=0
        self._current=0; self._hits={}; self._highlights={}
        self._sel_start=None; self._sel_page=None
        self._sel_items=[]; self._sel_text=""; self._word_cache={}

    def _on_configure(self, event):
        if self._doc and abs(event.width - self._cw) > 8:
            old = self._current
            self._cw = event.width
            self._invalidate_all()
            self._build_layout()
            self.go_to(old)
            self._update_h_scroll()

    def _show_msg(self, msg):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 400)
        self.canvas.create_text(cw//2, 140, text=msg,
                                fill=C["fg3"], font=("Georgia",10),
                                justify="center", width=cw-40)
        self._update_nav()


# =============================================================================
# WIDGET HELPERS
# =============================================================================
def _btn(p, text, cmd, bg=None, fg=None, size=9, padx=8, pady=3, **kw):
    return tk.Button(p, text=text, command=cmd,
                     bg=bg or C["bg4"], fg=fg or C["fg"], relief="flat",
                     font=("Georgia",size,"bold"), padx=padx, pady=pady,
                     cursor="hand2", activebackground=C["bg5"],
                     activeforeground=C["fg"], **kw)

def _entry(p, w=8, val=""):
    e = tk.Entry(p, width=w, bg=C["bg3"], fg=C["fg"],
                 insertbackground=C["fg"], relief="flat",
                 font=("Courier New",10))
    if val: e.insert(0, str(val))
    return e

def _sep(p):
    return tk.Frame(p, bg=C["bg4"], height=1)

def _tv_style(name):
    s = ttk.Style()
    s.configure(f"{name}.Treeview", background=C["bg3"], foreground=C["fg"],
                fieldbackground=C["bg3"], rowheight=24, font=("Georgia",9))
    s.configure(f"{name}.Treeview.Heading", background=C["bg4"],
                foreground=C["gold"], font=("Georgia",9,"bold"))
    s.map(f"{name}.Treeview", background=[("selected",C["bg5"])])
    return f"{name}.Treeview"


# =============================================================================
# MAIN APPLICATION
# =============================================================================
class DepoReview(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MacColl Transcript Review Tool")
        self.configure(bg=C["bg"])
        self.geometry("1440x900")
        self.minsize(1100, 700)
        self.resizable(True, True)

        self._transcript_path = None
        self._notes_path      = None
        self._exhibit_paths   = []
        self._exhibit_index   = 0
        self._flags           = {}
        self._notes_dirty     = False
        self._t_viewer        = None
        self._e_viewer        = None
        self._search_results  = []
        self._search_idx      = -1

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-f>", lambda _: self._focus_search())
        self.bind("<F5>",        lambda _: self._flag_page())

    # ── UI ─────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        tk.Frame(self, bg=C["gold"], height=3).pack(fill="x")
        hdr = tk.Frame(self, bg=C["bg"], padx=20, pady=9); hdr.pack(fill="x")

        badge = tk.Frame(hdr, bg=C["gold3"], width=44, height=44)
        badge.pack(side="left", padx=(0,14)); badge.pack_propagate(False)
        tk.Label(badge, text="M", bg=C["gold3"], fg=C["gold2"],
                 font=("Georgia",22,"bold")).place(relx=.5,rely=.5,anchor="center")

        tf = tk.Frame(hdr, bg=C["bg"]); tf.pack(side="left")
        tk.Label(tf, text="MacColl Transcript Review Tool",
                 bg=C["bg"], fg=C["gold"],
                 font=("Georgia",16,"bold")).pack(anchor="w")
        tk.Label(tf, text="ZUCKERMAN SPAEDER LLP  ·  DEPOSITION ANALYSIS",
                 bg=C["bg"], fg=C["fg3"], font=("Courier New",7)).pack(anchor="w")

        rf = tk.Frame(hdr, bg=C["bg"]); rf.pack(side="right")
        self._hdr_info = tk.Label(rf, text="No transcript loaded",
                                  bg=C["bg"], fg=C["fg3"],
                                  font=("Courier New",8), anchor="e")
        self._hdr_info.pack(anchor="e", pady=(0,4))
        br = tk.Frame(rf, bg=C["bg"]); br.pack()
        _btn(br, "💾  Save Notes",   self._save_notes,
             bg=C["gold3"], fg=C["gold2"]).pack(side="left", padx=(0,4))
        _btn(br, "📤  Export Notes", self._export_notes).pack(side="left")

        tk.Frame(self, bg=C["gold3"], height=1).pack(fill="x")

        main = tk.PanedWindow(self, orient="horizontal", bg=C["sash"],
                              sashwidth=5, sashrelief="flat")
        main.pack(fill="both", expand=True)
        lf = tk.Frame(main, bg=C["bg"])
        rf2= tk.Frame(main, bg=C["bg"])
        main.add(lf,  minsize=420, stretch="always")
        main.add(rf2, minsize=360, stretch="always")
        main.paneconfig(lf,  width=720)
        main.paneconfig(rf2, width=720)

        self._build_transcript_panel(lf)
        self._build_right_panel(rf2)

        sb = tk.Frame(self, bg=C["bg4"], pady=3); sb.pack(fill="x", side="bottom")
        self._status = tk.Label(sb, text="Ready.", bg=C["bg4"], fg=C["fg3"],
                                font=("Courier New",8), anchor="w", padx=10)
        self._status.pack(side="left")
        tk.Label(sb, text="Ctrl+F search  ·  F5 flag page  ·  scroll continuously  ·  zoom ± then drag to highlight",
                 bg=C["bg4"], fg=C["fg3"],
                 font=("Courier New",8), anchor="e", padx=10).pack(side="right")

    # ── Transcript panel ───────────────────────────────────────────────────────
    def _build_transcript_panel(self, p):
        self._sec_lbl(p, "TRANSCRIPT")

        tb1 = tk.Frame(p, bg=C["bg2"], pady=5); tb1.pack(fill="x")
        _btn(tb1, "📂  Load Transcript", self._load_transcript,
             bg=C["gold3"], fg=C["gold2"]).pack(side="left", padx=(8,4))
        self._flag_btn = _btn(tb1, "🚩  Flag This Page", self._flag_page,
                              bg=C["red2"], fg=C["red"])
        self._flag_btn.pack(side="left", padx=4)
        self._flag_btn.config(state="disabled")
        _btn(tb1, "Export PDF w/ Highlights",
             self._export_highlighted_pdf,
             bg=C["bg5"], fg=C["gold2"], size=8).pack(side="right", padx=(0,8))

        tb2 = tk.Frame(p, bg=C["bg2"], pady=4); tb2.pack(fill="x")
        self._t_prev_btn = _btn(tb2, "◄", self._t_prev, size=10)
        self._t_prev_btn.pack(side="left", padx=(8,2))
        self._t_prev_btn.config(state="disabled")
        self._t_page_lbl = tk.Label(tb2, text="— / —",
                                    bg=C["bg2"], fg=C["fg2"],
                                    font=("Courier New",9), width=9)
        self._t_page_lbl.pack(side="left", padx=2)
        self._t_next_btn = _btn(tb2, "►", self._t_next, size=10)
        self._t_next_btn.pack(side="left", padx=(2,6))
        self._t_next_btn.config(state="disabled")

        tk.Label(tb2, text="Go:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left")
        self._t_jump = _entry(tb2, w=5)
        self._t_jump.pack(side="left", padx=(2,8))
        self._t_jump.bind("<Return>", lambda _: self._jump_transcript())

        tk.Label(tb2, text="🔍", bg=C["bg2"], fg=C["fg2"],
                 font=("Georgia",10)).pack(side="left")
        self._search_var = tk.StringVar()
        se = tk.Entry(tb2, textvariable=self._search_var,
                      bg=C["bg3"], fg=C["fg"], insertbackground=C["fg"],
                      relief="flat", font=("Courier New",9), width=20)
        se.pack(side="left", padx=(2,2))
        se.bind("<Return>", lambda _: self._do_search())
        _btn(tb2, "Find", self._do_search, size=8).pack(side="left", padx=2)
        self._srch_prev = _btn(tb2, "▲", self._search_prev, size=8)
        self._srch_prev.pack(side="left", padx=1)
        self._srch_next = _btn(tb2, "▼", self._search_next, size=8)
        self._srch_next.pack(side="left", padx=1)
        self._srch_count = tk.Label(tb2, text="", bg=C["bg2"], fg=C["fg3"],
                                    font=("Courier New",8))
        self._srch_count.pack(side="left", padx=(4,0))

        # ── Zoom toolbar ──────────────────────────────────────────────────────
        tz = tk.Frame(p, bg=C["bg2"], pady=3); tz.pack(fill="x")
        tk.Label(tz, text="Zoom:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left", padx=(8,4))
        _btn(tz, "−", self._t_zoom_out, size=11, width=2).pack(side="left", padx=1)
        self._t_zoom_lbl = tk.Label(tz, text="100%", bg=C["bg3"], fg=C["gold"],
                                    font=("Courier New",8,"bold"),
                                    width=5, relief="flat", pady=2)
        self._t_zoom_lbl.pack(side="left", padx=2)
        _btn(tz, "+", self._t_zoom_in, size=11, width=2).pack(side="left", padx=1)
        _btn(tz, "Reset", self._t_zoom_reset, size=8, padx=4).pack(side="left", padx=(4,0))

        # ── Highlight toolbar ─────────────────────────────────────────────────
        th = tk.Frame(p, bg=C["bg2"], pady=3); th.pack(fill="x")
        self._hl_toggle_btn = _btn(th, "🖊  Highlight Mode",
                                   self._toggle_hl_mode, size=8)
        self._hl_toggle_btn.pack(side="left", padx=(8,4))
        self._t_sel_btn = _btn(th, "🔤  Select Text",
                               self._toggle_t_sel, size=8)
        self._t_sel_btn.pack(side="left", padx=(0,6))
        tk.Label(th, text="Color:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left")
        HL_COLORS = [
            ("#ffd700", "🟡"), ("#90ee90", "🟢"),
            ("#ffb6c1", "🩷"), ("#87ceeb", "🔵"),
        ]
        self._hl_color_btns = []
        for col, lbl in HL_COLORS:
            b = tk.Button(th, text=lbl, bg=C["bg4"], relief="flat",
                          font=("Georgia",11), cursor="hand2",
                          command=lambda c=col: self._set_hl_color(c))
            b.pack(side="left", padx=1)
            self._hl_color_btns.append((col, b))
        self._hl_color_btns[0][1].config(bg=C["bg5"])   # yellow selected by default
        tk.Frame(th, bg=C["bg4"], width=1).pack(side="left", fill="y", padx=6)
        _btn(th, "↩ Undo", self._undo_hl, size=8).pack(side="left", padx=2)
        _btn(th, "✕ Clear", self._clear_hl, size=8, fg=C["red"]).pack(side="left", padx=2)

        _sep(p).pack(fill="x")

        cf = tk.Frame(p, bg=C["bg2"]); cf.pack(fill="both", expand=True)
        self._t_canvas = tk.Canvas(cf, bg=C["bg2"], highlightthickness=0)
        t_vsb = ttk.Scrollbar(cf, orient="vertical")
        t_vsb.pack(side="right", fill="y")
        t_hsb = ttk.Scrollbar(cf, orient="horizontal", command=self._t_canvas.xview)
        t_hsb.pack(side="bottom", fill="x")
        self._t_canvas.pack(fill="both", expand=True)
        self._t_canvas.configure(
            yscrollcommand=self._make_scroll_cmd(t_vsb, "t"),
            xscrollcommand=t_hsb.set)
        t_vsb.configure(command=self._t_canvas.yview)
        self._t_canvas.bind("<MouseWheel>", self._t_scroll)
        self._t_canvas.bind("<Button-4>",   self._t_scroll)
        self._t_canvas.bind("<Button-5>",   self._t_scroll)

        self._t_viewer = DocViewer(
            self._t_canvas, self._t_page_lbl,
            self._t_prev_btn, self._t_next_btn,
            on_page_change=self._on_t_page_change)
        self._t_viewer._zoom_lbl    = self._t_zoom_lbl
        self._t_viewer._h_scrollbar = t_hsb
        self._attach_ctx(self._t_canvas, "t")
        self._t_canvas.bind("<Control-c>", lambda e: self._copy_t_sel())
        self._t_canvas.bind("<Control-C>", lambda e: self._copy_t_sel())

    def _make_scroll_cmd(self, scrollbar, which):
        def handler(first, last):
            scrollbar.set(first, last)
            v = self._t_viewer if which=="t" else self._e_viewer
            if v: v.notify_scroll()
        return handler

    def _t_scroll(self, event):
        if event.num==4:   self._t_canvas.yview_scroll(-3,"units")
        elif event.num==5: self._t_canvas.yview_scroll(3,"units")
        else:              self._t_canvas.yview_scroll(
                               int(-1*(event.delta/120)),"units")
        if self._t_viewer: self._t_viewer.notify_scroll()

    def _e_scroll(self, event):
        if event.num==4:   self._e_canvas.yview_scroll(-3,"units")
        elif event.num==5: self._e_canvas.yview_scroll(3,"units")
        else:              self._e_canvas.yview_scroll(
                               int(-1*(event.delta/120)),"units")
        if self._e_viewer: self._e_viewer.notify_scroll()

    # ── Right panel ────────────────────────────────────────────────────────────
    def _build_right_panel(self, p):
        rs = tk.PanedWindow(p, orient="vertical", bg=C["sash"],
                            sashwidth=5, sashrelief="flat")
        rs.pack(fill="both", expand=True)
        ef = tk.Frame(rs, bg=C["bg"])
        nf = tk.Frame(rs, bg=C["bg"])
        rs.add(ef, minsize=200, stretch="always")
        rs.add(nf, minsize=160, stretch="always")
        rs.paneconfig(ef, height=530)
        rs.paneconfig(nf, height=270)
        self._build_exhibit_panel(ef)
        self._build_notes_panel(nf)

    # ── Exhibit panel ──────────────────────────────────────────────────────────
    def _build_exhibit_panel(self, p):
        self._sec_lbl(p, "EXHIBITS")

        tb = tk.Frame(p, bg=C["bg2"], pady=5); tb.pack(fill="x")
        _btn(tb, "📁  Load Folder",
             self._load_exhibit_folder).pack(side="left", padx=(8,4))
        _btn(tb, "📄  Load PDF",
             self._load_exhibit_pdf).pack(side="left", padx=4)

        tb2 = tk.Frame(p, bg=C["bg2"], pady=3); tb2.pack(fill="x")
        self._ex_prev_btn = _btn(tb2, "◄", self._e_prev, size=10)
        self._ex_prev_btn.pack(side="left", padx=(8,2))
        self._ex_prev_btn.config(state="disabled")
        self._ex_page_lbl = tk.Label(tb2, text="— / —",
                                     bg=C["bg2"], fg=C["fg2"],
                                     font=("Courier New",9), width=9)
        self._ex_page_lbl.pack(side="left", padx=2)
        self._ex_next_btn = _btn(tb2, "►", self._e_next, size=10)
        self._ex_next_btn.pack(side="left", padx=(2,4))
        self._ex_next_btn.config(state="disabled")
        tk.Label(tb2, text="Go:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left")
        self._e_jump = _entry(tb2, w=5)
        self._e_jump.pack(side="left", padx=(2,8))
        self._e_jump.bind("<Return>", lambda _: self._jump_exhibit())
        tk.Frame(tb2, bg=C["bg4"], width=1).pack(side="left", fill="y", padx=6)
        self._e_sel_btn = _btn(tb2, "🔤  Select Text",
                               self._toggle_e_sel, size=8)
        self._e_sel_btn.pack(side="left", padx=2)

        sf = tk.Frame(p, bg=C["bg2"], pady=3); sf.pack(fill="x")
        tk.Label(sf, text="File:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left", padx=(8,4))
        self._ex_selector = ttk.Combobox(sf, state="readonly",
                                          font=("Courier New",8), width=34)
        self._ex_selector.pack(side="left")
        self._ex_selector.bind("<<ComboboxSelected>>",
                               self._on_exhibit_select)

        # ── Zoom toolbar ──────────────────────────────────────────────────────
        ez = tk.Frame(p, bg=C["bg2"], pady=3); ez.pack(fill="x")
        tk.Label(ez, text="Zoom:", bg=C["bg2"], fg=C["fg3"],
                 font=("Georgia",8)).pack(side="left", padx=(8,4))
        _btn(ez, "−", self._e_zoom_out, size=11, width=2).pack(side="left", padx=1)
        self._e_zoom_lbl = tk.Label(ez, text="100%", bg=C["bg3"], fg=C["gold"],
                                    font=("Courier New",8,"bold"),
                                    width=5, relief="flat", pady=2)
        self._e_zoom_lbl.pack(side="left", padx=2)
        _btn(ez, "+", self._e_zoom_in, size=11, width=2).pack(side="left", padx=1)
        _btn(ez, "Reset", self._e_zoom_reset, size=8, padx=4).pack(side="left", padx=(4,0))

        _sep(p).pack(fill="x")

        cf = tk.Frame(p, bg=C["bg2"]); cf.pack(fill="both", expand=True)
        self._e_canvas = tk.Canvas(cf, bg=C["bg2"], highlightthickness=0)
        e_vsb = ttk.Scrollbar(cf, orient="vertical")
        e_vsb.pack(side="right", fill="y")
        e_hsb = ttk.Scrollbar(cf, orient="horizontal", command=self._e_canvas.xview)
        e_hsb.pack(side="bottom", fill="x")
        self._e_canvas.pack(fill="both", expand=True)
        self._e_canvas.configure(
            yscrollcommand=self._make_scroll_cmd(e_vsb, "e"),
            xscrollcommand=e_hsb.set)
        e_vsb.configure(command=self._e_canvas.yview)
        self._e_canvas.bind("<MouseWheel>", self._e_scroll)
        self._e_canvas.bind("<Button-4>",   self._e_scroll)
        self._e_canvas.bind("<Button-5>",   self._e_scroll)
        self._e_canvas.create_text(200, 120,
            text="Load an exhibit folder or PDF above.",
            fill=C["fg3"], font=("Georgia",10), justify="center")

        self._e_viewer = DocViewer(
            self._e_canvas, self._ex_page_lbl,
            self._ex_prev_btn, self._ex_next_btn)
        self._e_viewer._zoom_lbl    = self._e_zoom_lbl
        self._e_viewer._h_scrollbar = e_hsb
        self._attach_ctx(self._e_canvas, "e")
        self._e_canvas.bind("<Control-c>", lambda e: self._copy_e_sel())
        self._e_canvas.bind("<Control-C>", lambda e: self._copy_e_sel())

    # ── Notes panel ────────────────────────────────────────────────────────────
    def _build_notes_panel(self, p):
        self._sec_lbl(p, "NOTES & FLAGS")
        sty = ttk.Style()
        sty.configure("Gold.TNotebook",     background=C["bg"], borderwidth=0)
        sty.configure("Gold.TNotebook.Tab", background=C["bg4"],
                      foreground=C["fg2"], padding=(10,4),
                      font=("Georgia",8,"bold"))
        sty.map("Gold.TNotebook.Tab",
                background=[("selected",C["bg3"])],
                foreground=[("selected",C["gold"])])
        nb = ttk.Notebook(p, style="Gold.TNotebook")
        nb.pack(fill="both", expand=True, padx=4, pady=(2,4))

        ft = tk.Frame(nb, bg=C["bg3"])
        nb.add(ft, text="  🚩 Flagged Pages  ")
        ftb = tk.Frame(ft, bg=C["bg3"], pady=3); ftb.pack(fill="x")
        _btn(ftb,"Remove",    self._remove_flag,    bg=C["bg4"],size=8).pack(side="left",padx=(6,4))
        _btn(ftb,"Edit Note", self._edit_flag_note, bg=C["bg4"],size=8).pack(side="left",padx=2)
        _btn(ftb,"Jump →",self._jump_to_flag, bg=C["bg5"],fg=C["gold"],size=8).pack(side="right",padx=6)
        ff = tk.Frame(ft, bg=C["bg3"]); ff.pack(fill="both", expand=True, padx=4, pady=(0,4))
        self._flag_tv = ttk.Treeview(ff, columns=("Page","Note"),
                                     show="headings", style=_tv_style("flag"))
        self._flag_tv.heading("Page", text="Page")
        self._flag_tv.column("Page", width=50, minwidth=40, stretch=False)
        self._flag_tv.heading("Note", text="Note")
        self._flag_tv.column("Note", width=280, minwidth=80)
        fvsb = ttk.Scrollbar(ff, orient="vertical", command=self._flag_tv.yview)
        self._flag_tv.configure(yscrollcommand=fvsb.set)
        fvsb.pack(side="right", fill="y")
        self._flag_tv.pack(fill="both", expand=True)
        self._flag_tv.bind("<Double-1>", lambda _: self._jump_to_flag())

        nt = tk.Frame(nb, bg=C["bg3"])
        nb.add(nt, text="  📝 Notes  ")

        # toolbar: timestamp | copy | clear
        ntb = tk.Frame(nt, bg=C["bg3"], pady=3); ntb.pack(fill="x")
        _btn(ntb, "🕔 Timestamp", self._insert_timestamp,
             bg=C["bg4"], size=8).pack(side="left", padx=(6,2))
        _btn(ntb, "📋 Copy All",  self._copy_notes,
             bg=C["bg4"], size=8).pack(side="left", padx=2)
        _btn(ntb, "🗑 Clear", self._clear_notes_confirm,
             bg=C["bg4"], size=8, fg=C["red"]).pack(side="right", padx=6)

        # main text area
        nf = tk.Frame(nt, bg=C["bg3"]); nf.pack(fill="both", expand=True, padx=4, pady=(0,4))
        self._notes_text = tk.Text(
            nf, bg=C["bg3"], fg=C["fg"],
            insertbackground=C["gold"],
            selectbackground=C["bg5"], selectforeground=C["fg"],
            relief="flat", font=("Georgia", 10),
            wrap="word", undo=True,
            spacing1=2, spacing3=2,
            pady=8, padx=10)
        nvsb = ttk.Scrollbar(nf, orient="vertical", command=self._notes_text.yview)
        self._notes_text.configure(yscrollcommand=nvsb.set)
        nvsb.pack(side="right", fill="y")
        self._notes_text.pack(fill="both", expand=True)
        self._notes_text.bind("<Control-a>",
            lambda e: (self._notes_text.tag_add("sel","1.0","end"), "break"))
        self._notes_text.bind("<Key>",
            lambda _: setattr(self,"_notes_dirty",True))

    def _sec_lbl(self, parent, text):
        f = tk.Frame(parent, bg=C["bg4"], pady=5); f.pack(fill="x")
        tk.Label(f, text=f"  {text}", bg=C["bg4"], fg=C["gold"],
                 font=("Courier New",7,"bold"), anchor="w").pack(side="left",padx=4)

    # ══════════════════════════════════════════════════════════════════════════
    # TRANSCRIPT ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _load_transcript(self):
        fp = filedialog.askopenfilename(
            title="Open Transcript",
            filetypes=[("PDF","*.pdf"),
                       ("Images","*.jpg *.jpeg *.png *.tif *.tiff"),
                       ("All","*.*")])
        if not fp: return
        fp = Path(fp)
        if self._t_viewer.load(fp):
            self._transcript_path = fp
            self._notes_path = fp.parent / (fp.stem + NOTES_SUFFIX)
            self._load_notes_file()
            self._flag_btn.config(state="normal")
            self._hdr_info.config(
                text=f"{fp.name}  ·  {self._t_viewer.total_pages} pages",
                fg=C["fg2"])
            self._set_status(
                f"Loaded: {fp.name}  ({self._t_viewer.total_pages} pages)")
        else:
            messagebox.showerror("Cannot Open", f"Could not open:\n{fp.name}")

    def _t_prev(self):
        if self._t_viewer: self._t_viewer.prev()
    def _t_next(self):
        if self._t_viewer: self._t_viewer.next()

    def _jump_transcript(self):
        try: self._t_viewer.go_to(int(self._t_jump.get().strip())-1)
        except: pass
        self._t_jump.delete(0,"end")

    def _on_t_page_change(self, pno):
        txt = "🚩  Edit Flag" if pno in self._flags else "🚩  Flag This Page"
        self._flag_btn.config(text=txt)

    def _focus_search(self):
        self._search_var.set("")

    def _do_search(self):
        q = self._search_var.get().strip()
        if not q or not self._t_viewer: return
        self._search_results = self._t_viewer.search(q)
        self._search_idx     = -1
        n = len(self._search_results)
        self._srch_count.config(
            text=(f"{n} match{'es' if n!=1 else ''}" if n else "No matches"))
        if n: self._search_next()
        self._set_status(f"Search \"{q}\": {n} result{'s' if n!=1 else ''}.")

    def _search_next(self):
        if not self._search_results: return
        self._search_idx = (self._search_idx+1) % len(self._search_results)
        self._t_viewer.go_to(self._search_results[self._search_idx][0])
        self._srch_count.config(
            text=f"{self._search_idx+1}/{len(self._search_results)}")

    def _search_prev(self):
        if not self._search_results: return
        self._search_idx = (self._search_idx-1) % len(self._search_results)
        self._t_viewer.go_to(self._search_results[self._search_idx][0])
        self._srch_count.config(
            text=f"{self._search_idx+1}/{len(self._search_results)}")

    # ══════════════════════════════════════════════════════════════════════════
    # FLAGS
    # ══════════════════════════════════════════════════════════════════════════
    def _flag_page(self):
        if not self._t_viewer or not self._t_viewer.total_pages: return
        pno  = self._t_viewer.current_page
        note = simpledialog.askstring(
            "Flag Page",
            f"Page {pno+1} — add a note (optional):",
            initialvalue=self._flags.get(pno,""), parent=self)
        if note is None: return
        self._flags[pno] = note
        self._refresh_flags(); self._notes_dirty = True
        self._on_t_page_change(pno)
        self._set_status(f"Page {pno+1} flagged.")

    def _remove_flag(self):
        sel = self._flag_tv.selection()
        if not sel: return
        pno = int(self._flag_tv.item(sel[0],"values")[0]) - 1
        if pno in self._flags:
            del self._flags[pno]
            self._refresh_flags(); self._notes_dirty = True
            if self._t_viewer: self._on_t_page_change(self._t_viewer.current_page)

    def _edit_flag_note(self):
        sel = self._flag_tv.selection()
        if not sel: return
        pno  = int(self._flag_tv.item(sel[0],"values")[0]) - 1
        note = simpledialog.askstring(
            "Edit Flag Note", f"Page {pno+1}:",
            initialvalue=self._flags.get(pno,""), parent=self)
        if note is None: return
        self._flags[pno] = note
        self._refresh_flags(); self._notes_dirty = True

    def _jump_to_flag(self):
        sel = self._flag_tv.selection()
        if not sel: return
        pno = int(self._flag_tv.item(sel[0],"values")[0]) - 1
        self._t_viewer.go_to(pno)

    def _refresh_flags(self):
        self._flag_tv.delete(*self._flag_tv.get_children())
        for pno in sorted(self._flags):
            self._flag_tv.insert("","end",
                values=(pno+1,(self._flags[pno] or "")[:120]))

    # ══════════════════════════════════════════════════════════════════════════
    # EXHIBITS
    # ══════════════════════════════════════════════════════════════════════════
    def _load_exhibit_folder(self):
        folder = filedialog.askdirectory(title="Select Exhibit Folder")
        if not folder: return
        folder = Path(folder)
        files  = sorted([p for p in folder.iterdir()
                         if p.suffix.lower() in
                         (".pdf",".jpg",".jpeg",".png",".tif",".tiff",
                          ".bmp",".gif",".webp")],
                        key=lambda p: p.name.lower())
        if not files:
            messagebox.showinfo("No Files",
                                "No supported files found in that folder.")
            return
        self._exhibit_paths = files
        self._exhibit_index = 0
        self._ex_selector.config(values=[p.name for p in files])
        self._ex_selector.set(files[0].name)
        self._open_exhibit(0)
        self._set_status(f"Loaded {len(files)} exhibit files from {folder.name}/")

    def _load_exhibit_pdf(self):
        fp = filedialog.askopenfilename(
            title="Open Exhibit PDF",
            filetypes=[("PDF","*.pdf"),
                       ("Images","*.jpg *.jpeg *.png *.tif *.tiff"),
                       ("All","*.*")])
        if not fp: return
        fp = Path(fp)
        self._exhibit_paths = [fp]
        self._exhibit_index = 0
        self._ex_selector.config(values=[fp.name])
        self._ex_selector.set(fp.name)
        self._open_exhibit(0)
        self._set_status(f"Loaded exhibit: {fp.name}")

    def _on_exhibit_select(self, _=None):
        idx = self._ex_selector.current()
        if idx >= 0: self._open_exhibit(idx)

    def _open_exhibit(self, idx):
        if not self._exhibit_paths: return
        idx  = max(0, min(idx, len(self._exhibit_paths)-1))
        self._exhibit_index = idx
        path = self._exhibit_paths[idx]
        if not self._e_viewer.load(path):
            self._e_canvas.delete("all")
            self._e_canvas.create_text(200,120,
                text=f"Cannot render:\n{path.name}",
                fill=C["fg3"], font=("Georgia",10), justify="center")

    def _e_prev(self):
        if self._e_viewer: self._e_viewer.prev()
    def _e_next(self):
        if self._e_viewer: self._e_viewer.next()

    def _jump_exhibit(self):
        try:    self._e_viewer.go_to(int(self._e_jump.get().strip())-1)
        except: pass
        self._e_jump.delete(0,"end")

    # ══════════════════════════════════════════════════════════════════════════
    # ZOOM ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _t_zoom_in(self):
        if self._t_viewer: self._t_viewer.zoom_in()
    def _t_zoom_out(self):
        if self._t_viewer: self._t_viewer.zoom_out()
    def _t_zoom_reset(self):
        if self._t_viewer:
            self._t_viewer._zoom = 1.0
            self._t_viewer._apply_zoom()

    def _e_zoom_in(self):
        if self._e_viewer: self._e_viewer.zoom_in()
    def _e_zoom_out(self):
        if self._e_viewer: self._e_viewer.zoom_out()
    def _e_zoom_reset(self):
        if self._e_viewer:
            self._e_viewer._zoom = 1.0
            self._e_viewer._apply_zoom()

    # ══════════════════════════════════════════════════════════════════════════
    # HIGHLIGHT ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _toggle_hl_mode(self):
        if not self._t_viewer: return
        on = not self._t_viewer._hl_mode
        self._t_viewer.set_hl_mode(on)
        if on:
            self._hl_toggle_btn.config(bg=C["gold3"], fg=C["gold2"],
                                       text="🖊  Highlighting ON")
        else:
            self._hl_toggle_btn.config(bg=C["bg4"], fg=C["fg"],
                                       text="🖊  Highlight Mode")

    def _set_hl_color(self, color: str):
        if self._t_viewer:
            self._t_viewer.set_hl_color(color)
        # Update button backgrounds
        for col, btn in self._hl_color_btns:
            btn.config(bg=C["bg5"] if col == color else C["bg4"])

    def _undo_hl(self):
        if self._t_viewer:
            self._t_viewer.undo_last_highlight()
            self._notes_dirty = True

    def _clear_hl(self):
        if not self._t_viewer: return
        if messagebox.askyesno("Clear Highlights",
                               "Remove all highlights from this transcript?"):
            self._t_viewer.clear_highlights()
            self._notes_dirty = True

        # ══════════════════════════════════════════════════════════════════════════
    # NOTES PERSISTENCE
    # ══════════════════════════════════════════════════════════════════════════
    def _load_notes_file(self):
        if not self._notes_path or not self._notes_path.exists():
            self._flags={}; self._notes_text.delete("1.0","end")
            self._refresh_flags(); return
        try:
            with open(self._notes_path, encoding="utf-8") as f:
                data = json.load(f)
            self._flags = {int(k):v for k,v in data.get("flags",{}).items()}
            self._notes_text.delete("1.0","end")
            self._notes_text.insert("1.0", data.get("notes",""))
            if self._t_viewer and data.get("highlights"):
                self._t_viewer.load_highlights(data["highlights"])
            self._refresh_flags(); self._notes_dirty = False
            self._set_status(f"Notes loaded from {self._notes_path.name}")
        except Exception as e:
            self._set_status(f"Could not load notes: {e}")

    def _save_notes(self):
        if not self._notes_path:
            fp = filedialog.asksaveasfilename(
                title="Save Notes", defaultextension=".json",
                filetypes=[("JSON","*.json")])
            if not fp: return
            self._notes_path = Path(fp)
        try:
            data = dict(saved_at=datetime.now().isoformat(),
                        transcript=str(self._transcript_path or ""),
                        flags={str(k):v for k,v in self._flags.items()},
                        notes=self._notes_text.get("1.0","end").rstrip(),
                        highlights=(self._t_viewer.get_highlights()
                                    if self._t_viewer else {}))
            with open(self._notes_path,"w",encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self._notes_dirty = False
            self._set_status(f"Notes saved → {self._notes_path.name}")
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    def _export_notes(self):
        if not self._flags and not self._notes_text.get("1.0","end").strip():
            messagebox.showinfo("Nothing to Export","No flags or notes yet.")
            return
        fp = filedialog.asksaveasfilename(
            title="Export Notes", defaultextension=".txt",
            filetypes=[("Text","*.txt"),("CSV","*.csv")])
        if not fp: return
        fp = Path(fp)
        try:
            if fp.suffix.lower() == ".csv":
                with open(fp,"w",newline="",encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(["Type","Page","Content"])
                    for pno in sorted(self._flags):
                        w.writerow(["Flag",pno+1,self._flags[pno] or ""])
                    fr = self._notes_text.get("1.0","end").strip()
                    if fr: w.writerow(["Notes","",fr])
            else:
                lines = [
                    "MacColl Transcript Review Tool — Export",
                    "="*60,
                    f"Transcript : {self._transcript_path or 'unknown'}",
                    f"Exported   : {datetime.now().strftime('%Y-%m-%d %H:%M')}",""]
                if self._flags:
                    lines += ["FLAGGED PAGES","-"*40]
                    for pno in sorted(self._flags):
                        lines.append(f"  Page {pno+1:>4} :  "
                                     f"{self._flags[pno] or '(no note)'}")
                    lines.append("")
                fr = self._notes_text.get("1.0","end").strip()
                if fr: lines += ["FREE-FORM NOTES","-"*40, fr]
                with open(fp,"w",encoding="utf-8") as f:
                    f.write("\n".join(lines))
            self._set_status(f"Exported → {fp.name}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def _clear_notes_confirm(self):
        if messagebox.askyesno("Clear Notes",
                               "Clear all free-form notes?\nFlags will not be affected."):
            self._notes_text.delete("1.0","end"); self._notes_dirty=True

    def _set_status(self, msg):
        self._status.config(text=msg)


    # ══════════════════════════════════════════════════════════════════════════
    # CONTEXT MENUS & TEXT COPY
    # ══════════════════════════════════════════════════════════════════════════
    def _attach_ctx(self, canvas, which):
        """Right-click context menu on viewer canvases."""
        menu = tk.Menu(canvas, tearoff=0,
                       bg=C["bg3"], fg=C["fg"],
                       activebackground=C["bg5"], activeforeground=C["fg"],
                       font=("Georgia", 9))
        if which == "t":
            menu.add_command(label="Export PDF with highlights",
                             command=self._export_highlighted_pdf)

        def _show(event):
            try:    menu.tk_popup(event.x_root, event.y_root)
            finally: menu.grab_release()

        canvas.bind("<Button-3>", _show)
        canvas.bind("<Button-2>", _show)


    # ── notes helpers ──────────────────────────────────────────────────────────
    def _insert_timestamp(self):
        ts  = datetime.now().strftime("%Y-%m-%d %H:%M")
        pno = (self._t_viewer.current_page + 1
               if self._t_viewer and self._t_viewer.total_pages else None)
        tag = "[" + ts
        if pno:
            tag += "  |  Transcript p." + str(pno)
        tag += "]\n"
        self._notes_text.insert("insert", tag)
        self._notes_text.see("insert")
        self._notes_dirty = True

    def _copy_notes(self):
        text = self._notes_text.get("1.0", "end").strip()
        if not text:
            self._set_status("Notes are empty."); return
        self.clipboard_clear()
        self.clipboard_append(text)
        self._set_status("Notes copied to clipboard.")

    # ══════════════════════════════════════════════════════════════════════════
    # EXPORT PDF WITH HIGHLIGHTS
    # ══════════════════════════════════════════════════════════════════════════
    def _export_highlighted_pdf(self):
        if not self._transcript_path:
            messagebox.showinfo("No Transcript",
                "Please load a transcript PDF first.")
            return
        if self._transcript_path.suffix.lower() != ".pdf":
            messagebox.showinfo("PDF Only",
                "Export with highlights is only available for PDF transcripts.")
            return

        highlights = self._t_viewer.get_highlights() if self._t_viewer else {}
        stem      = self._transcript_path.stem
        save_path = filedialog.asksaveasfilename(
            title="Export Transcript PDF with Highlights",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=stem + "_highlighted.pdf")
        if not save_path: return

        try:
            doc = fitz.open(str(self._transcript_path))

            for pno_key, page_hl in highlights.items():
                pno  = int(pno_key)
                if pno >= len(doc): continue
                page = doc[pno]
                pw, ph = page.rect.width, page.rect.height

                for (nx0, ny0, nx1, ny1, color) in page_hl:
                    rect = fitz.Rect(
                        nx0 * pw, ny0 * ph,
                        nx1 * pw, ny1 * ph)
                    hx = color.lstrip("#")
                    r  = int(hx[0:2], 16) / 255
                    g  = int(hx[2:4], 16) / 255
                    b  = int(hx[4:6], 16) / 255
                    page.draw_rect(rect,
                                   color=(r, g, b),
                                   fill=(r, g, b),
                                   fill_opacity=0.35,
                                   overlay=True)

            doc.save(str(save_path), garbage=4, deflate=True)
            doc.close()

            n_hl    = sum(len(v) for v in highlights.values())
            n_pages = len(highlights)
            if n_hl:
                detail = (str(n_hl) + " highlight" + ("s" if n_hl != 1 else "")
                          + " across " + str(n_pages)
                          + " page" + ("s" if n_pages != 1 else "") + ".")
            else:
                detail = "Clean copy — no highlights applied."
            messagebox.showinfo("Export Complete",
                "Saved to:\n" + str(Path(save_path).name) + "\n\n" + detail)
            self._set_status("Exported to " + str(Path(save_path).name))

        except Exception as exc:
            messagebox.showerror("Export Failed",
                "Could not write PDF:\n" + str(exc))

    # ── text selection (app layer) ─────────────────────────────────────────────
    def _toggle_t_sel(self):
        if not self._t_viewer: return
        on = not self._t_viewer._sel_mode
        self._t_viewer.set_sel_mode(on)
        if on:
            self._hl_toggle_btn.config(
                bg=C["bg4"], fg=C["fg"], text="🖊  Highlight Mode")
        self._t_sel_btn.config(
            bg=C["blue"] if on else C["bg4"],
            fg=C["bg"]   if on else C["fg"],
            text=("🔤  Selecting..." if on else "🔤  Select Text"))

    def _toggle_e_sel(self):
        if not self._e_viewer: return
        on = not self._e_viewer._sel_mode
        self._e_viewer.set_sel_mode(on)
        self._e_sel_btn.config(
            bg=C["blue"] if on else C["bg4"],
            fg=C["bg"]   if on else C["fg"],
            text=("🔤  Selecting..." if on else "🔤  Select Text"))

    def _copy_sel(self, viewer):
        text = viewer.get_selected_text() if viewer else ""
        if not text.strip(): return
        self.clipboard_clear()
        self.clipboard_append(text)
        snip = text[:60].strip().replace("\n", " ")
        self._set_status("Copied: " + repr(snip) + ("..." if len(text) > 60 else ""))

    def _copy_t_sel(self): self._copy_sel(self._t_viewer)
    def _copy_e_sel(self): self._copy_sel(self._e_viewer)

    def _on_close(self):
        if self._notes_dirty:
            ans = messagebox.askyesnocancel(
                "Unsaved Notes",
                "You have unsaved notes.\n\nSave before closing?")
            if ans is None: return
            if ans: self._save_notes()
        if self._t_viewer: self._t_viewer.close()
        if self._e_viewer: self._e_viewer.close()
        self.destroy()


if __name__ == "__main__":
    app = DepoReview()
    sty = ttk.Style(app); sty.theme_use("clam")
    sty.configure("TScrollbar",
                  background=C["bg4"], troughcolor=C["bg2"],
                  bordercolor=C["bg2"], arrowcolor=C["fg3"])
    sty.configure("Vertical.TScrollbar",
                  background=C["bg4"], troughcolor=C["bg2"])
    app.mainloop()
