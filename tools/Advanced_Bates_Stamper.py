"""
Bates Stamper GUI
==================
Stamps Bates numbers onto PDF and image files, and produces renamed
native placeholders for all other file types (Word, Excel, CSV, etc.).

Supports:
  - PDF (.pdf)         → stamped PDF, pages numbered sequentially
  - Images (.jpg, .jpeg, .png, .gif, .bmp, .tiff, .tif)
                       → converted to stamped PDF
  - Natives (.docx, .xlsx, .csv, .pptx, and all others)
                       → single-page "PRODUCED AS NATIVE" placeholder PDF
                          + renamed copy of the original file
  - PST (.pst)         → blocked with a clear message (see below)

Stamps applied to every page:
  - Bottom right: Bates number  (e.g., SMITH_00000042)
  - Bottom left:  Optional confidentiality text  (e.g., CONFIDENTIAL)

Output:
  - All stamped/placeholder PDFs and renamed natives saved to a folder
    you choose
  - A production_index.csv listing every document with its Bates range,
    original filename, type, page count, and whether produced as native

--------------------------------------------------------------
HOW TO RUN (Windows)
--------------------------------------------------------------
Do NOT double-click this file — it won't work that way.
Instead, launch it through PowerShell:

  1. Open File Explorer and navigate to the folder where
     this file lives.

  2. Click the address bar, type  powershell  and press Enter.
     (Use PowerShell, not CMD — especially on network drives \\)

  3. In the PowerShell window, type:
         python Advanced_Bates_Stamper.py
     and press Enter.

--------------------------------------------------------------
IF YOU GET "Missing dependency" OR AN IMPORT ERROR
--------------------------------------------------------------
Run this in PowerShell, then try again:

    python -m pip install pymupdf Pillow

--------------------------------------------------------------
A NOTE ON PST FILES
--------------------------------------------------------------
If the source folder contains .pst files (Outlook archive files),
this tool will detect them and show a warning. PST files are
containers that can hold thousands of emails and attachments —
they require special processing outside the scope of this tool.

To include PST content in your production:
  * Export the emails from Outlook as individual .msg or .pdf files, OR
  * Use a dedicated e-discovery tool to process the PST first.
Then place the extracted files in the source folder and stamp normally.

--------------------------------------------------------------
"""

import os
import sys
import csv
import io
import queue
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime

# ── Required dependencies ──────────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install pymupdf")

try:
    from PIL import Image as PILImage, ImageTk
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install Pillow")


# ── File type definitions ──────────────────────────────────────────────────
PDF_EXTS   = {'.pdf'}
IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif'}
PST_EXTS   = {'.pst'}
SKIP_EXTS  = {'.py', '.json'}


def file_category(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in PDF_EXTS:    return 'PDF'
    if ext in IMAGE_EXTS:  return 'Image'
    if ext in PST_EXTS:    return 'PST'
    if ext in SKIP_EXTS:   return 'Skip'
    return 'Native'


TYPE_TAG = {'PDF': '[PDF]   ', 'Image': '[IMG]   ', 'Native': '[NATIVE]'}


# ── Bates formatting ───────────────────────────────────────────────────────

def fmt(prefix: str, number: int, padding: int) -> str:
    return f"{prefix}_{str(number).zfill(padding)}"


# ── Stamp visual constants ─────────────────────────────────────────────────
_FS   = 7
_MARG = 5
_PAD  = 3


def stamp_page(page, bates_str: str, conf_str: str) -> None:
    """Stamp one page: bold 14pt text, no border boxes."""
    r    = page.rect
    FONT = "Helvetica-Bold"
    FS   = 14

    # Bates number -- bottom right
    bw = fitz.get_text_length(bates_str, fontname=FONT, fontsize=FS)
    page.insert_text(
        fitz.Point(r.x1 - bw - _MARG, r.y1 - _MARG),
        bates_str, fontname=FONT, fontsize=FS, color=(0, 0, 0))

    # Confidentiality text -- bottom left
    if conf_str.strip():
        page.insert_text(
            fitz.Point(_MARG, r.y1 - _MARG),
            conf_str, fontname=FONT, fontsize=FS, color=(0, 0, 0))


def make_native_placeholder(bates_str: str, orig_name: str,
                            conf_str: str) -> fitz.Document:
    doc  = fitz.open()
    page = doc.new_page(width=612, height=792)

    title = "PRODUCED AS NATIVE"
    t_fs  = 20
    t_w   = fitz.get_text_length(title, fontname="Helvetica-Bold", fontsize=t_fs)
    page.insert_text(fitz.Point((612 - t_w) / 2, 375),
                     title, fontname="Helvetica-Bold", fontsize=t_fs,
                     color=(0.1, 0.1, 0.1))

    b_fs = 14
    b_w  = fitz.get_text_length(bates_str, fontname="helv", fontsize=b_fs)
    page.insert_text(fitz.Point((612 - b_w) / 2, 405),
                     bates_str, fontname="helv", fontsize=b_fs,
                     color=(0.2, 0.2, 0.2))

    n_fs    = 10
    display = orig_name
    while (fitz.get_text_length(display, fontname="helv", fontsize=n_fs) > 520
           and len(display) > 12):
        display = display[:-1]
    if display != orig_name:
        display = display[:-2] + "..."
    n_w = fitz.get_text_length(display, fontname="helv", fontsize=n_fs)
    page.insert_text(fitz.Point((612 - n_w) / 2, 428),
                     display, fontname="helv", fontsize=n_fs,
                     color=(0.45, 0.45, 0.45))

    stamp_page(page, bates_str, conf_str)
    return doc


def image_to_pdf(img_path: Path) -> fitz.Document:
    pil_img    = PILImage.open(str(img_path)).convert("RGB")
    w_px, h_px = pil_img.size
    PAGE_W, PAGE_H = 612, 792
    scale = min(PAGE_W / w_px, PAGE_H / h_px)
    dw, dh = w_px * scale, h_px * scale
    x0, y0 = (PAGE_W - dw) / 2, (PAGE_H - dh) / 2
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    doc  = fitz.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    page.insert_image(fitz.Rect(x0, y0, x0 + dw, y0 + dh),
                      stream=buf.getvalue())
    return doc


# ── Preview helpers ────────────────────────────────────────────────────────

PREV_H = 640   # target pixel height for preview renders


def _pdf_page_to_pil(doc: fitz.Document, page_no: int) -> PILImage.Image:
    page = doc[page_no]
    zoom = PREV_H / page.rect.height
    pix  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)


def _image_to_pil(path: Path) -> PILImage.Image:
    img = PILImage.open(str(path)).convert("RGB")
    if img.height > PREV_H:
        scale = PREV_H / img.height
        img   = img.resize((int(img.width * scale), PREV_H), PILImage.LANCZOS)
    return img


def _native_info_card(path: Path) -> PILImage.Image:
    """Clean info card rendered via fitz for native (non-PDF/image) files."""
    W, H = 520, 290
    doc  = fitz.open()
    page = doc.new_page(width=W, height=H)

    # Blue accent bar at top
    page.draw_rect(fitz.Rect(0, 0, W, 6),
                   color=None, fill=(0.537, 0.706, 0.98), width=0)

    # "NATIVE FILE" heading
    title = "NATIVE FILE"
    tw = fitz.get_text_length(title, fontname="Helvetica-Bold", fontsize=20)
    page.insert_text(fitz.Point((W - tw) / 2, 68),
                     title, fontname="Helvetica-Bold", fontsize=20,
                     color=(0.2, 0.35, 0.65))

    # Filename (truncated if needed)
    fn = path.name
    while (fitz.get_text_length(fn, fontname="helv", fontsize=11) > W - 40
           and len(fn) > 10):
        fn = fn[:-1]
    if fn != path.name:
        fn = fn[:-2] + "..."
    fw = fitz.get_text_length(fn, fontname="helv", fontsize=11)
    page.insert_text(fitz.Point((W - fw) / 2, 100),
                     fn, fontname="helv", fontsize=11,
                     color=(0.15, 0.15, 0.15))

    # Extension row
    ext_text = f"Extension:  {path.suffix.lower() or '(none)'}"
    ew = fitz.get_text_length(ext_text, fontname="helv", fontsize=9)
    page.insert_text(fitz.Point((W - ew) / 2, 128),
                     ext_text, fontname="helv", fontsize=9,
                     color=(0.4, 0.4, 0.4))

    # Divider
    page.draw_line(fitz.Point(60, 152), fitz.Point(W - 60, 152),
                   color=(0.82, 0.82, 0.82), width=0.5)

    # Production note
    for y, line, fs, col in [
        (172, "This file will be produced as:",          9,  (0.35, 0.35, 0.35)),
        (191, "  * A renamed native copy  (BATES" + path.suffix + ")",
                                                          8,  (0.22, 0.22, 0.22)),
        (207, "  * A placeholder PDF  (BATES.pdf)",      8,  (0.22, 0.22, 0.22)),
    ]:
        lw = fitz.get_text_length(line, fontname="helv", fontsize=fs)
        page.insert_text(fitz.Point((W - lw) / 2, y),
                         line, fontname="helv", fontsize=fs, color=col)

    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def _error_card(msg: str) -> PILImage.Image:
    doc  = fitz.open()
    page = doc.new_page(width=420, height=180)
    page.insert_text(fitz.Point(20, 70), "Preview unavailable",
                     fontname="Helvetica-Bold", fontsize=14,
                     color=(0.7, 0.2, 0.2))
    em = msg[:130] + ("..." if len(msg) > 130 else "")
    page.insert_text(fitz.Point(20, 100), em,
                     fontname="helv", fontsize=8, color=(0.4, 0.4, 0.4))
    pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
    img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


# ── Main Application ───────────────────────────────────────────────────────

class BatesStamper(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Bates Stamper")
        self.configure(bg="#1e1e2e")
        self.geometry("1560x860")
        self.minsize(1000, 640)
        self.resizable(True, True)

        self.file_entries = []
        self._drag_idx    = None
        self._q           = queue.Queue()

        # Preview state
        self._prev_doc   = None   # open fitz.Document for PDF previews
        self._prev_img   = None   # PhotoImage ref (must stay alive)
        self._prev_page  = 0
        self._prev_total = 1
        self._prev_idx   = -1     # list index currently shown in preview

        self._build_ui()

    # ══════════════════════════════════════════════════════════════════════
    # UI construction
    # ══════════════════════════════════════════════════════════════════════

    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg="#181825", pady=8)
        top.pack(fill="x")

        _lbl(top, "Source Folder:").pack(side="left", padx=(14, 4))
        self.var_source = tk.StringVar()
        tk.Entry(top, textvariable=self.var_source,
                 **_ent_cfg(), width=62).pack(side="left", padx=(0, 4))
        _btn(top, "Browse...", self._browse_source).pack(side="left", padx=2)
        _btn(top, "Load Files", self._load_files,
             bg="#89b4fa", fg="#1e1e2e").pack(side="left", padx=(8, 0))

        # Three-pane layout
        pw = tk.PanedWindow(self, orient="horizontal", bg="#1e1e2e",
                            sashwidth=5, sashrelief="flat")
        pw.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        file_pane    = tk.Frame(pw, bg="#1e1e2e")
        preview_pane = tk.Frame(pw, bg="#181825")
        config_pane  = tk.Frame(pw, bg="#1e1e2e")

        pw.add(file_pane,    minsize=220, stretch="never")
        pw.add(preview_pane, minsize=280, stretch="always")
        pw.add(config_pane,  minsize=400, stretch="never")

        pw.paneconfig(file_pane,    width=300)
        pw.paneconfig(preview_pane, width=540)
        pw.paneconfig(config_pane,  width=640)

        self._build_file_panel(file_pane)
        self._build_preview_panel(preview_pane)
        self._build_config_panel(config_pane)
        self._build_results_panel(config_pane)

    # ── File list ─────────────────────────────────────────────────────────

    def _build_file_panel(self, parent):
        hdr = tk.Frame(parent, bg="#1e1e2e")
        hdr.pack(fill="x", pady=(6, 4), padx=4)
        _lbl(hdr, "FILES  (drag to reorder)").pack(side="left")
        self.lbl_count = tk.Label(hdr, text="", bg="#1e1e2e",
                                  fg="#6c7086", font=("Helvetica", 9))
        self.lbl_count.pack(side="right")

        lf = tk.Frame(parent, bg="#181825")
        lf.pack(fill="both", expand=True, padx=4)

        vsb = ttk.Scrollbar(lf, orient="vertical")
        self.listbox = tk.Listbox(
            lf, bg="#181825", fg="#cdd6f4",
            selectbackground="#313244", selectforeground="#89b4fa",
            font=("Courier", 9), relief="flat", activestyle="none",
            cursor="hand2", yscrollcommand=vsb.set)
        vsb.config(command=self.listbox.yview)
        vsb.pack(side="right", fill="y")
        self.listbox.pack(fill="both", expand=True)

        self.listbox.bind("<Button-1>",        self._drag_start)
        self.listbox.bind("<B1-Motion>",       self._drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self._drag_release)

        btn_row = tk.Frame(parent, bg="#1e1e2e")
        btn_row.pack(fill="x", padx=4, pady=(4, 0))
        _btn(btn_row, "Up",     self._move_up  ).pack(side="left", padx=(0, 2))
        _btn(btn_row, "Down",   self._move_down).pack(side="left", padx=2)
        _btn(btn_row, "Remove", self._remove_selected,
             bg="#f38ba8", fg="#1e1e2e").pack(side="right")

    # ── Preview panel ─────────────────────────────────────────────────────

    def _build_preview_panel(self, parent):
        # Header: filename label + page navigation
        hdr = tk.Frame(parent, bg="#181825", pady=4)
        hdr.pack(fill="x")

        sm = dict(bg="#313244", fg="#cdd6f4", relief="flat",
                  font=("Helvetica", 8), padx=5, pady=2,
                  cursor="hand2", activebackground="#45475a",
                  activeforeground="#cdd6f4")

        self.btn_prev_pg = tk.Button(hdr, text="<- pg",
                                     command=self._prev_pg_back, **sm)
        self.btn_prev_pg.pack(side="right", padx=(2, 6))

        self.lbl_prev_page = tk.Label(hdr, text="", bg="#181825",
                                      fg="#6c7086", font=("Helvetica", 8))
        self.lbl_prev_page.pack(side="right")

        self.btn_next_pg = tk.Button(hdr, text="pg ->",
                                     command=self._prev_pg_fwd, **sm)
        self.btn_next_pg.pack(side="right", padx=2)

        self.lbl_prev_name = tk.Label(
            hdr, text="Click a file in the list to preview it",
            bg="#181825", fg="#6c7086",
            font=("Helvetica", 9), anchor="w")
        self.lbl_prev_name.pack(side="left", padx=8)

        # Scrollable canvas
        cf = tk.Frame(parent, bg="#181825")
        cf.pack(fill="both", expand=True)

        self.prev_canvas = tk.Canvas(cf, bg="#181825", highlightthickness=0)
        vsb = ttk.Scrollbar(cf, orient="vertical",
                            command=self.prev_canvas.yview)
        hsb = ttk.Scrollbar(cf, orient="horizontal",
                            command=self.prev_canvas.xview)
        self.prev_canvas.configure(yscrollcommand=vsb.set,
                                   xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.prev_canvas.pack(fill="both", expand=True)

        self.prev_canvas.bind("<MouseWheel>", self._prev_scroll)
        self.prev_canvas.bind("<Button-4>",   self._prev_scroll)
        self.prev_canvas.bind("<Button-5>",   self._prev_scroll)

        # Disable page nav buttons initially
        self.btn_prev_pg.config(state="disabled")
        self.btn_next_pg.config(state="disabled")

    # ── Config / progress / start ─────────────────────────────────────────

    def _build_config_panel(self, parent):
        cfg = tk.Frame(parent, bg="#1e1e2e")
        cfg.pack(fill="x", padx=10, pady=(8, 0))

        _lbl(cfg, "BATES CONFIGURATION", size=10).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        _lbl(cfg, "Prefix:", size=9).grid(row=1, column=0, sticky="w", padx=(0, 4))
        self.var_prefix = tk.StringVar(value="SMITH")
        self.var_prefix.trace_add("write", self._update_bates_preview)
        tk.Entry(cfg, textvariable=self.var_prefix,
                 **_ent_cfg(), width=16).grid(row=1, column=1, sticky="w")

        _lbl(cfg, "Starting Number:", size=9).grid(
            row=1, column=2, sticky="w", padx=(16, 4))
        self.var_start = tk.StringVar(value="1")
        self.var_start.trace_add("write", self._update_bates_preview)
        tk.Entry(cfg, textvariable=self.var_start,
                 **_ent_cfg(), width=14).grid(row=1, column=3, sticky="w")

        _lbl(cfg, "Zero-padding digits:", size=9).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var_padding = tk.StringVar(value="8")
        self.var_padding.trace_add("write", self._update_bates_preview)
        tk.Entry(cfg, textvariable=self.var_padding,
                 **_ent_cfg(), width=5).grid(row=2, column=2, sticky="w",
                                              pady=(8, 0))

        _lbl(cfg, "First stamp preview:", size=9).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var_bpreview = tk.StringVar(value="SMITH_00000001")
        tk.Label(cfg, textvariable=self.var_bpreview,
                 bg="#313244", fg="#a6e3a1",
                 font=("Courier", 11, "bold"), padx=8, pady=3).grid(
            row=3, column=2, columnspan=2, sticky="w", pady=(8, 0))

        tk.Frame(cfg, bg="#313244", height=1).grid(
            row=4, column=0, columnspan=4, sticky="ew", pady=12)

        _lbl(cfg, "Bottom-left stamp text  (optional):", size=9).grid(
            row=5, column=0, columnspan=4, sticky="w")
        self.var_conf = tk.StringVar()
        tk.Entry(cfg, textvariable=self.var_conf,
                 **_ent_cfg(), width=52).grid(
            row=6, column=0, columnspan=4, sticky="w", pady=(4, 2))
        tk.Label(cfg, text="e.g.  CONFIDENTIAL  or  CONFIDENTIAL - ATTORNEYS EYES ONLY",
                 bg="#1e1e2e", fg="#6c7086",
                 font=("Helvetica", 8)).grid(
            row=7, column=0, columnspan=4, sticky="w")

        tk.Frame(cfg, bg="#313244", height=1).grid(
            row=8, column=0, columnspan=4, sticky="ew", pady=12)

        _lbl(cfg, "Output Folder:", size=9).grid(
            row=9, column=0, sticky="w", padx=(0, 4))
        self.var_output = tk.StringVar()
        tk.Entry(cfg, textvariable=self.var_output,
                 **_ent_cfg(), width=34).grid(
            row=9, column=1, columnspan=2, sticky="ew", padx=4)
        _btn(cfg, "Browse...", self._browse_output).grid(
            row=9, column=3, sticky="w")

        cfg.columnconfigure(1, weight=1)

        pb_frame = tk.Frame(parent, bg="#1e1e2e")
        pb_frame.pack(fill="x", padx=10, pady=(14, 0))

        self.var_status = tk.StringVar(value="")
        tk.Label(pb_frame, textvariable=self.var_status,
                 bg="#1e1e2e", fg="#6c7086",
                 font=("Helvetica", 9), anchor="w").pack(fill="x")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("stamp.Horizontal.TProgressbar",
                        background="#a6e3a1", troughcolor="#313244",
                        bordercolor="#1e1e2e",
                        lightcolor="#a6e3a1", darkcolor="#a6e3a1")
        self.pb_var = tk.IntVar(value=0)
        ttk.Progressbar(pb_frame, variable=self.pb_var, maximum=100,
                        style="stamp.Horizontal.TProgressbar").pack(
            fill="x", pady=(4, 0))

        self.btn_start = tk.Button(
            parent, text="START BATES STAMPING",
            bg="#a6e3a1", fg="#1e1e2e", relief="flat",
            font=("Helvetica", 12, "bold"), pady=10,
            cursor="hand2", activebackground="#94e2d5",
            activeforeground="#1e1e2e",
            command=self._start_stamping)
        self.btn_start.pack(fill="x", padx=10, pady=(10, 0))

    def _build_results_panel(self, parent):
        self.results_frame = tk.Frame(parent, bg="#1e1e2e")
        self.results_frame.pack(fill="both", expand=True,
                                padx=10, pady=(12, 0))

    # ══════════════════════════════════════════════════════════════════════
    # File loading
    # ══════════════════════════════════════════════════════════════════════

    def _browse_source(self):
        d = filedialog.askdirectory(title="Select Source Folder")
        if d:
            self.var_source.set(d)
            self._load_files()

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.var_output.set(d)

    def _load_files(self):
        src = self.var_source.get().strip()
        if not src or not Path(src).is_dir():
            messagebox.showerror("Error",
                                 "Please enter a valid source folder path.")
            return

        self._clear_preview()
        self.file_entries = []
        pst_found = []
        own_name  = Path(__file__).name

        for entry in sorted(os.scandir(src), key=lambda e: e.name.lower()):
            if not entry.is_file():
                continue
            path = Path(entry.path)
            if path.name == own_name:
                continue
            cat = file_category(path)
            if cat == 'Skip':
                continue
            if cat == 'PST':
                pst_found.append(path.name)
                continue
            self.file_entries.append({
                'path':     path,
                'category': cat,
                'display':  f"{TYPE_TAG.get(cat, '[?]'):<10} {path.name}",
            })

        self._refresh_listbox()

        if pst_found:
            msg = (
                "This tool cannot process PST files.\n\n"
                "PST files are Outlook archive containers that can hold thousands "
                "of emails and attachments. Opening and rendering them requires "
                "separate processing outside the scope of this stamper.\n\n"
                "The following PST file(s) were found and will be skipped:\n\n"
                + "".join(f"   *  {f}\n" for f in pst_found)
                + "\nTo include this content in your production, first export the "
                "emails from Outlook as individual files (e.g., PDF or MSG), "
                "or process the PST with a dedicated e-discovery tool. Then place "
                "the extracted files in the source folder and stamp normally."
            )
            messagebox.showwarning("PST Files Detected - Action Required", msg)

    def _refresh_listbox(self):
        self.listbox.delete(0, "end")
        for fe in self.file_entries:
            self.listbox.insert("end", f"  {fe['display']}")
        n = len(self.file_entries)
        self.lbl_count.config(text=f"{n} file{'s' if n != 1 else ''}")

    # ══════════════════════════════════════════════════════════════════════
    # Drag / reorder
    # ══════════════════════════════════════════════════════════════════════

    def _drag_start(self, event):
        self._drag_idx = self.listbox.nearest(event.y)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(self._drag_idx)

    def _drag_motion(self, event):
        idx = self.listbox.nearest(event.y)
        if self._drag_idx is not None and idx != self._drag_idx:
            fe = self.file_entries
            fe[self._drag_idx], fe[idx] = fe[idx], fe[self._drag_idx]
            self._refresh_listbox()
            self.listbox.selection_set(idx)
            self._drag_idx = idx

    def _drag_release(self, event):
        """Show preview of whatever item is selected at the end of the click/drag."""
        sel = self.listbox.curselection()
        if sel:
            self._preview_file(sel[0])
        self._drag_idx = None

    def _move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        fe = self.file_entries
        fe[i - 1], fe[i] = fe[i], fe[i - 1]
        self._refresh_listbox()
        self.listbox.selection_set(i - 1)
        self._preview_file(i - 1)

    def _move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.file_entries) - 1:
            return
        i = sel[0]
        fe = self.file_entries
        fe[i], fe[i + 1] = fe[i + 1], fe[i]
        self._refresh_listbox()
        self.listbox.selection_set(i + 1)
        self._preview_file(i + 1)

    def _remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        self.file_entries.pop(sel[0])
        self._refresh_listbox()
        self._clear_preview()

    # ══════════════════════════════════════════════════════════════════════
    # Preview
    # ══════════════════════════════════════════════════════════════════════

    def _preview_file(self, idx: int):
        if idx < 0 or idx >= len(self.file_entries):
            return
        if idx == self._prev_idx:
            return  # already showing this file

        fe   = self.file_entries[idx]
        path = fe['path']
        cat  = fe['category']

        self._prev_idx  = idx
        self._prev_page = 0

        if self._prev_doc:
            self._prev_doc.close()
            self._prev_doc = None

        self.lbl_prev_name.config(text=path.name, fg="#89b4fa")

        try:
            if cat == 'PDF':
                self._prev_doc   = fitz.open(str(path))
                self._prev_total = len(self._prev_doc)
                img = _pdf_page_to_pil(self._prev_doc, 0)
            elif cat == 'Image':
                self._prev_total = 1
                img = _image_to_pil(path)
            else:
                self._prev_total = 1
                img = _native_info_card(path)
        except Exception as e:
            self._prev_total = 1
            img = _error_card(str(e))

        self._show_preview_image(img)
        self._update_prev_nav()

    def _show_preview_image(self, img: PILImage.Image):
        self._prev_img = ImageTk.PhotoImage(img)
        self.prev_canvas.delete("all")
        self.prev_canvas.create_image(4, 4, anchor="nw",
                                      image=self._prev_img)
        self.prev_canvas.config(
            scrollregion=(0, 0, img.width + 8, img.height + 8))
        self.prev_canvas.yview_moveto(0)

    def _clear_preview(self):
        if self._prev_doc:
            self._prev_doc.close()
            self._prev_doc = None
        self._prev_idx   = -1
        self._prev_page  = 0
        self._prev_total = 1
        self.prev_canvas.delete("all")
        self.lbl_prev_name.config(
            text="Click a file in the list to preview it", fg="#6c7086")
        self.lbl_prev_page.config(text="")
        self.btn_prev_pg.config(state="disabled")
        self.btn_next_pg.config(state="disabled")

    def _update_prev_nav(self):
        n = self._prev_total
        p = self._prev_page
        if n > 1:
            self.lbl_prev_page.config(text=f" {p + 1} / {n} ")
            self.btn_prev_pg.config(state="normal" if p > 0     else "disabled")
            self.btn_next_pg.config(state="normal" if p < n - 1 else "disabled")
        else:
            self.lbl_prev_page.config(text="")
            self.btn_prev_pg.config(state="disabled")
            self.btn_next_pg.config(state="disabled")

    def _prev_pg_back(self):
        if self._prev_doc and self._prev_page > 0:
            self._prev_page -= 1
            img = _pdf_page_to_pil(self._prev_doc, self._prev_page)
            self._show_preview_image(img)
            self._update_prev_nav()

    def _prev_pg_fwd(self):
        if self._prev_doc and self._prev_page < self._prev_total - 1:
            self._prev_page += 1
            img = _pdf_page_to_pil(self._prev_doc, self._prev_page)
            self._show_preview_image(img)
            self._update_prev_nav()

    def _prev_scroll(self, event):
        if event.num == 4:
            self.prev_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.prev_canvas.yview_scroll(1, "units")
        else:
            self.prev_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ══════════════════════════════════════════════════════════════════════
    # Bates stamp preview label
    # ══════════════════════════════════════════════════════════════════════

    def _update_bates_preview(self, *_):
        try:
            prefix  = self.var_prefix.get().strip()
            start   = int(self.var_start.get().strip() or "1")
            padding = int(self.var_padding.get().strip() or "8")
            self.var_bpreview.set(fmt(prefix, start, padding))
        except (ValueError, tk.TclError):
            self.var_bpreview.set("(invalid input)")

    # ══════════════════════════════════════════════════════════════════════
    # Stamping
    # ══════════════════════════════════════════════════════════════════════

    def _start_stamping(self):
        errors = []
        if not self.file_entries:
            errors.append("No files loaded — use Load Files first.")
        if not self.var_prefix.get().strip():
            errors.append("Bates prefix cannot be empty.")
        try:
            s = int(self.var_start.get().strip())
            if s < 1:
                errors.append("Starting number must be at least 1.")
        except ValueError:
            errors.append("Starting number must be a whole number.")
        try:
            p = int(self.var_padding.get().strip())
            if not 1 <= p <= 14:
                errors.append("Padding digits must be between 1 and 14.")
        except ValueError:
            errors.append("Padding digits must be a whole number.")
        if not self.var_output.get().strip():
            errors.append("Please select an output folder.")

        if errors:
            messagebox.showerror(
                "Cannot Start",
                "Please fix the following before stamping:\n\n"
                + "\n".join(f"  *  {e}" for e in errors))
            return

        self.btn_start.config(state="disabled", text="Stamping in progress...")
        self.pb_var.set(0)
        self.var_status.set("")

        threading.Thread(target=self._worker, daemon=True).start()
        self.after(120, self._poll)

    def _worker(self):
        prefix  = self.var_prefix.get().strip()
        start   = int(self.var_start.get().strip())
        padding = int(self.var_padding.get().strip())
        conf    = self.var_conf.get().strip()
        out_dir = Path(self.var_output.get().strip())

        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self._q.put(('error', f"Cannot create output folder:\n{e}"))
            return

        counter = start
        records = []
        errs    = []
        total   = len(self.file_entries)

        for i, fe in enumerate(self.file_entries):
            path = fe['path']
            cat  = fe['category']
            self._q.put(('status', f"({i+1}/{total})  {path.name}"))

            try:
                first = fmt(prefix, counter, padding)

                if cat == 'PDF':
                    doc = fitz.open(str(path))
                    n   = len(doc)
                    for pg in doc:
                        stamp_page(pg, fmt(prefix, counter, padding), conf)
                        counter += 1
                    last     = fmt(prefix, counter - 1, padding)
                    out_name = first + '.pdf'
                    doc.save(str(out_dir / out_name))
                    doc.close()
                    records.append(_rec(first, last, path.name,
                                        out_name, 'PDF', n, 'No'))

                elif cat == 'Image':
                    doc = image_to_pdf(path)
                    n   = len(doc)
                    for pg in doc:
                        stamp_page(pg, fmt(prefix, counter, padding), conf)
                        counter += 1
                    last     = fmt(prefix, counter - 1, padding)
                    out_name = first + '.pdf'
                    doc.save(str(out_dir / out_name))
                    doc.close()
                    records.append(_rec(first, last, path.name,
                                        out_name, 'Image->PDF', n, 'No'))

                else:
                    ph      = make_native_placeholder(first, path.name, conf)
                    out_pdf = first + '.pdf'
                    out_nat = first + path.suffix
                    ph.save(str(out_dir / out_pdf))
                    ph.close()
                    shutil.copy2(str(path), str(out_dir / out_nat))
                    counter += 1
                    records.append(_rec(first, first, path.name,
                                        out_pdf, 'Native', 1, 'Yes'))

            except Exception as exc:
                errs.append(f"{path.name}: {exc}")

            self._q.put(('progress', int((i + 1) / total * 100)))

        stamp    = datetime.now().strftime('%Y%m%d_%H%M%S')
        csv_path = out_dir / f"production_index_{stamp}.csv"
        try:
            fields = ['bates_begin', 'bates_end', 'original_filename',
                      'stamped_filename', 'file_type', 'page_count',
                      'produced_native']
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                w = csv.DictWriter(f, fieldnames=fields)
                w.writeheader()
                w.writerows(records)
        except Exception as exc:
            errs.append(f"CSV write error: {exc}")

        self._q.put(('done', records, str(csv_path), errs))

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                if msg[0] == 'progress':
                    self.pb_var.set(msg[1])
                elif msg[0] == 'status':
                    self.var_status.set(msg[1])
                elif msg[0] == 'error':
                    messagebox.showerror("Error", msg[1])
                    self.btn_start.config(state="normal",
                                         text="START BATES STAMPING")
                    return
                elif msg[0] == 'done':
                    _, records, csv_path, errs = msg
                    self._stamping_done(records, csv_path, errs)
                    return
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _stamping_done(self, records, csv_path, errs):
        self.btn_start.config(state="normal",
                              text="START BATES STAMPING")
        self.pb_var.set(100)
        self.var_status.set(
            f"Done -- {len(records)} document(s) stamped. "
            f"Index saved to {Path(csv_path).name}")

        if errs:
            messagebox.showwarning(
                "Completed with errors",
                f"{len(errs)} file(s) could not be stamped:\n\n"
                + "\n".join(f"  *  {e}" for e in errs[:12])
                + ("\n  ...and more." if len(errs) > 12 else ""))

        self._show_results(records, csv_path)

    # ══════════════════════════════════════════════════════════════════════
    # Results table
    # ══════════════════════════════════════════════════════════════════════

    def _show_results(self, records, csv_path):
        for w in self.results_frame.winfo_children():
            w.destroy()

        _lbl(self.results_frame, "PRODUCTION INDEX").pack(
            fill="x", pady=(0, 4))
        tk.Label(self.results_frame,
                 text=f"CSV saved -> {csv_path}",
                 bg="#1e1e2e", fg="#6c7086",
                 font=("Helvetica", 8), anchor="w",
                 wraplength=570).pack(fill="x", pady=(0, 6))

        cols    = ('begin', 'end', 'orig', 'type', 'pgs', 'native')
        headers = ('Bates Begin', 'Bates End', 'Original Filename',
                   'Type', 'Pgs', 'Native')
        widths  = [120, 120, 200, 80, 40, 50]

        style = ttk.Style()
        style.configure("idx.Treeview",
                        background="#181825", foreground="#cdd6f4",
                        fieldbackground="#181825", rowheight=19,
                        font=("Courier", 8))
        style.configure("idx.Treeview.Heading",
                        background="#313244", foreground="#89b4fa",
                        font=("Helvetica", 8, "bold"))
        style.map("idx.Treeview", background=[("selected", "#313244")])

        frame = tk.Frame(self.results_frame, bg="#181825")
        frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        tv  = ttk.Treeview(frame, columns=cols, show="headings",
                           style="idx.Treeview",
                           yscrollcommand=vsb.set,
                           xscrollcommand=hsb.set)
        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        for col, hdr, w in zip(cols, headers, widths):
            tv.heading(col, text=hdr)
            tv.column(col, width=w, minwidth=30)

        for r in records:
            tv.insert("", "end", values=(
                r['bates_begin'], r['bates_end'],
                r['original_filename'], r['file_type'],
                r['page_count'],  r['produced_native']))

        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        tv.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════
    # Lifecycle
    # ══════════════════════════════════════════════════════════════════════

    def destroy(self):
        if self._prev_doc:
            self._prev_doc.close()
        super().destroy()


# ── Small UI helpers ───────────────────────────────────────────────────────

def _lbl(parent, text, size=10):
    return tk.Label(parent, text=text, bg="#1e1e2e", fg="#a6adc8",
                    font=("Helvetica", size, "bold"), anchor="w")


def _ent_cfg():
    return dict(bg="#313244", fg="#cdd6f4",
                insertbackground="#cdd6f4", relief="flat",
                font=("Helvetica", 10))


def _btn(parent, text, cmd, bg="#313244", fg="#cdd6f4"):
    return tk.Button(parent, text=text, command=cmd,
                     bg=bg, fg=fg, relief="flat",
                     font=("Helvetica", 9, "bold"), padx=8, pady=3,
                     cursor="hand2",
                     activebackground="#45475a" if bg == "#313244" else "#74c7ec",
                     activeforeground=fg)


def _rec(begin, end, orig, stamped, ftype, pages, native):
    return {'bates_begin': begin, 'bates_end': end,
            'original_filename': orig, 'stamped_filename': stamped,
            'file_type': ftype, 'page_count': pages,
            'produced_native': native}


# ── Entry point ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = BatesStamper()
    app.mainloop()
