#!/usr/bin/env python3
"""
Advanced_PDF_Splitter.py
MacColl's Advanced PDF Document Splitter — v2.0
Detect · Review · Extract · Own It
"""

import os, re, sys, json, math, threading, io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

# ── Optional dependencies ────────────────────────────────────────────────────
HAS_FITZ = HAS_PIL = HAS_PYPDF = HAS_PDFPLUMBER = HAS_ANTHROPIC = False
try:
    import fitz;            HAS_FITZ       = True   # PyMuPDF
except ImportError: pass
try:
    from PIL import Image, ImageTk; HAS_PIL = True
except ImportError: pass
try:
    from pypdf import PdfReader, PdfWriter; HAS_PYPDF = True
except ImportError: pass
try:
    import pdfplumber;      HAS_PDFPLUMBER = True
except ImportError: pass
try:
    import anthropic;       HAS_ANTHROPIC  = True
except ImportError: pass


# ── Palette ──────────────────────────────────────────────────────────────────
C = {
    "bg":        "#0b0e14",
    "surface":   "#131720",
    "panel":     "#1a2030",
    "panel2":    "#1f2840",
    "border":    "#2a3550",
    "gold":      "#c9a42a",
    "gold_lt":   "#f0c84a",
    "gold_dk":   "#7a6010",
    "gold_pale": "#e8d48a",
    "accent":    "#4a90d9",
    "success":   "#3db86a",
    "warning":   "#d4831a",
    "danger":    "#c03030",
    "text":      "#dce6f0",
    "text_dim":  "#7a8ea8",
    "text_mute": "#3a4a60",
    "entry":     "#0b0e14",
    "row_a":     "#131720",
    "row_b":     "#181e2c",
    "sel":       "#1e3460",
    "vbg":       "#090c12",
    "vpage":     "#1a2030",
}

SEG_PALETTE = [
    "#3a1a6a","#1a4a2a","#1a3060","#6a2a1a",
    "#1a5a4a","#5a1a3a","#3a4a1a","#4a1a1a",
]


# ── Date extraction ───────────────────────────────────────────────────────────
_MON = {
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
    'january':1,'february':2,'march':3,'april':4,
    'june':6,'july':7,'august':8,'september':9,
    'october':10,'november':11,'december':12,
}

def _parse_dates(text: str) -> list:
    found = []
    t = text[:6000]
    # ISO YYYY-MM-DD
    for m in re.finditer(r'\b(20\d{2})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])\b', t):
        try: found.append(datetime(int(m[1]),int(m[2]),int(m[3])))
        except: pass
    # Month DD, YYYY
    pat = (r'\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|'
           r'Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
           r'\.?\s+(\d{1,2}),?\s+((?:19|20)\d{2})\b')
    for m in re.finditer(pat, t, re.I):
        try:
            mo = _MON.get(m[1].lower().rstrip('.'))
            if mo: found.append(datetime(int(m[3]),mo,int(m[2])))
        except: pass
    # MM/DD/YYYY  MM-DD-YYYY
    for m in re.finditer(r'\b(0?[1-9]|1[0-2])[/\-](0?[1-9]|[12]\d|3[01])[/\-]((?:19|20)\d{2})\b', t):
        try: found.append(datetime(int(m[3]),int(m[1]),int(m[2])))
        except: pass
    # DD Month YYYY
    pat2 = (r'\b(\d{1,2})\s+(January|February|March|April|May|June|July|August|'
            r'September|October|November|December)\s+((?:19|20)\d{2})\b')
    for m in re.finditer(pat2, t, re.I):
        try:
            mo = _MON.get(m[2].lower())
            if mo: found.append(datetime(int(m[3]),mo,int(m[1])))
        except: pass
    return found

def latest_date(text: str) -> str:
    dates = _parse_dates(text)
    return max(dates).strftime("%m/%d/%Y") if dates else ""


# ── Text extraction ───────────────────────────────────────────────────────────
def extract_pages(pdf_path: str) -> list:
    pages = []
    if HAS_PDFPLUMBER:
        with pdfplumber.open(pdf_path) as pdf:
            for i, pg in enumerate(pdf.pages):
                txt = pg.extract_text() or ""
                pages.append({"page": i+1, "text": txt, "chars": len(txt)})
    elif HAS_PYPDF:
        r = PdfReader(pdf_path)
        for i, pg in enumerate(r.pages):
            txt = pg.extract_text() or ""
            pages.append({"page": i+1, "text": txt, "chars": len(txt)})
    return pages


# ── Heuristic detection ───────────────────────────────────────────────────────
def _first_meaningful_line(text: str) -> str:
    for ln in text.strip().split('\n')[:10]:
        ln = ln.strip()
        if 3 < len(ln) < 120 and not ln.isdigit():
            return re.sub(r'\s+',' ',ln)[:70]
    return ""

def heuristic_detect(pages: list) -> list:
    total = len(pages)
    if not total: return []
    scores  = [0]*total
    reasons = [[] for _ in range(total)]
    PATS = [
        (r'\bpage\s+1\b',                    8,  "Page 1 marker"),
        (r'\bpage\s+1\s+of\s+\d+',          13,  "Page 1 of N"),
        (r'^(to:|from:|subject:|re:|date:)',  6,  "Memo header"),
        (r'\b(agreement|contract|invoice|receipt|memorandum|addendum|amendment|exhibit)\b',
                                              4,  "Doc keyword"),
        (r'^(certificate|declaration|affidavit|report|proposal|notice)\b',
                                              6,  "Doc-type heading"),
    ]
    for i, p in enumerate(pages):
        txt   = p["text"].strip()
        first = txt.lower()[:500]
        if p["chars"] < 50:
            scores[i]  += 5; reasons[i].append("Blank separator")
        if i == 0:
            scores[i]  += 20; reasons[i].append("First page"); continue
        for pat,sc,rsn in PATS:
            if re.search(pat,first,re.I|re.M):
                scores[i]+=sc; reasons[i].append(rsn)
        if pages[i-1]["chars"] < 50 and p["chars"] > 200:
            scores[i]+=7; reasons[i].append("After blank page")
        lines=[ln.strip() for ln in txt.split('\n') if ln.strip()]
        if lines and len(lines[0])<80 and lines[0]==lines[0].upper() and len(lines[0])>3:
            scores[i]+=5; reasons[i].append("All-caps title")

    THRESH = 7
    bounds = [i for i in range(total) if scores[i]>=THRESH]
    if not bounds or bounds[0]!=0: bounds.insert(0,0)

    segs=[]
    for n,(j,start) in enumerate([(j,bounds[j]) for j in range(len(bounds))]):
        end  = bounds[j+1]-1 if j+1<len(bounds) else total-1
        conf = min(scores[start]/20.0,1.0) if start>0 else 1.0
        pg_txt = "\n".join(p["text"] for p in pages[start:end+1])
        segs.append({"start":start+1,"end":end+1,
                     "name":f"Exhibit_{n+1}",
                     "date":latest_date(pg_txt),
                     "confidence":conf,
                     "reason":", ".join(reasons[start]) if reasons[start] else "Heuristic",
                     "source":"heuristic"})
    return segs


# ── AI detection ──────────────────────────────────────────────────────────────
def ai_detect(pages, api_key=None, progress_cb=None):
    if not HAS_ANTHROPIC: raise RuntimeError("anthropic package not installed")
    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
    total = len(pages); CHUNK=60; all_segs=[]
    for c0 in range(0,total,CHUNK):
        chunk = pages[c0:c0+CHUNK]
        if progress_cb:
            progress_cb(f"AI analyzing pages {chunk[0]['page']}–{chunk[-1]['page']}…")
        sums = "\n".join(
            f"[P{p['page']}]({p['chars']}ch): {p['text'][:240].replace(chr(10),' ')}"
            for p in chunk)
        prompt=(f"Analyze pages {chunk[0]['page']}–{chunk[-1]['page']} of {total}.\n"
                f"Find document boundaries. Signals: 'Page 1', 'Page 1 of N', memo headers, "
                f"blank pages, topic shifts, doc-type keywords.\n\nSummaries:\n{sums}\n\n"
                f"Reply ONLY with JSON (no markdown):\n"
                f'{{"segments":[{{"start_page":N,"end_page":N,"suggested_name":"...",'
                f'"confidence":0.0,"reason":"..."}}]}}\nCover ALL pages with no gaps.')
        try:
            resp = client.messages.create(model="claude-sonnet-4-20250514",
                                          max_tokens=2000,
                                          messages=[{"role":"user","content":prompt}])
            raw = re.sub(r'^```json\s*|\s*```$','',resp.content[0].text.strip())
            data = json.loads(raw)
            for s in data.get("segments",[]):
                n = len(all_segs)+1
                pg_txt="\n".join(p["text"] for p in pages[s["start_page"]-1:s["end_page"]])
                all_segs.append({"start":s["start_page"],"end":s["end_page"],
                                 "name":f"Exhibit_{n}","date":latest_date(pg_txt),
                                 "confidence":s.get("confidence",0.7),
                                 "reason":s.get("reason","AI"),"source":"ai"})
        except Exception as e:
            if progress_cb: progress_cb(f"AI error chunk: {e}")
            fb=heuristic_detect(chunk)
            off=chunk[0]['page']-1
            for s in fb:
                s["start"]+=off; s["end"]+=off
                s["name"]=f"Exhibit_{len(all_segs)+1}"
                s["source"]="fallback"
                all_segs.append(s)
    all_segs.sort(key=lambda s:s["start"])
    for n,s in enumerate(all_segs,1): s["name"]=f"Exhibit_{n}"
    return all_segs


# ── PDF splitting ─────────────────────────────────────────────────────────────
def split_pdf(src, segs, outdir, progress_cb=None):
    reader=PdfReader(src); total=len(reader.pages); files=[]
    for idx,seg in enumerate(segs):
        s,e=max(1,seg["start"]),min(total,seg["end"])
        safe=re.sub(r'[<>:"/\\|?*]','_',seg["name"]).strip('. ')[:80] or f"doc_{idx+1}"
        out=os.path.join(outdir,f"{safe}.pdf")
        if progress_cb: progress_cb(f"Writing {safe}.pdf  (p{s}–{e})")
        w=PdfWriter()
        for pn in range(s-1,e): w.add_page(reader.pages[pn])
        with open(out,"wb") as f: w.write(f)
        files.append(out)
    return files


# ─────────────────────────────────────────────────────────────────────────────
#  GUI APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class App:
    # ── Init ─────────────────────────────────────────────────────────────────
    def __init__(self, root):
        self.root       = root
        self.root.title("MacColl's Advanced PDF Splitter")
        self.root.geometry("1400x860")
        self.root.minsize(1100,700)
        self.root.configure(bg=C["bg"])

        # State
        self.pdf_path    = None
        self.total_pages = 0
        self.page_texts  = []
        self.segments    = []
        self.page_images = {}          # 1-indexed -> PhotoImage
        self.page_y      = {}          # 1-indexed -> canvas y
        self._sel_idx    = None        # selected segment index
        self._sort_col   = None
        self._sort_rev   = False
        self._anim_step  = 0
        self._anim_msg   = 0

        # Edit StringVars (must exist before _build_ui)
        self.ev_name  = tk.StringVar()
        self.ev_start = tk.StringVar()
        self.ev_end   = tk.StringVar()
        self.ev_date  = tk.StringVar()

        self._build_ui()
        self._ttk_styles()
        self._start_anim()

    # ── Top-level layout ──────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        tk.Frame(self.root, bg=C["gold"], height=2).pack(fill="x")
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=6, pady=(4,2))
        self._build_viewer(body)
        self._build_right(body)
        self._build_statusbar()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C["panel"], height=68)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        # ── Hexagonal M logo ──
        lc = tk.Canvas(hdr, width=56, height=56, bg=C["panel"], highlightthickness=0)
        lc.pack(side="left", padx=(12,8), pady=6)
        pts=[]
        for a in range(0,360,60):
            ang=math.radians(a-90)
            pts += [28+24*math.cos(ang), 28+24*math.sin(ang)]
        lc.create_polygon(pts, fill=C["gold"], outline=C["gold_lt"], width=2)
        # Inner dark hex
        pts2=[]
        for a in range(0,360,60):
            ang=math.radians(a-90)
            pts2 += [28+20*math.cos(ang), 28+20*math.sin(ang)]
        lc.create_polygon(pts2, fill=C["panel2"], outline="")
        lc.create_text(28,30, text="M", fill=C["gold_lt"],
                       font=("Georgia",22,"bold"))

        # ── Title ──
        tf = tk.Frame(hdr, bg=C["panel"])
        tf.pack(side="left", pady=10)
        tk.Label(tf, text="MacColl's PDF Splitter",
                 bg=C["panel"], fg=C["gold"],
                 font=("Georgia",19,"bold")).pack(anchor="w")
        tk.Label(tf, text="Advanced Document Intelligence  ·  v2.0",
                 bg=C["panel"], fg=C["text_dim"],
                 font=("Courier New",8)).pack(anchor="w")

        # ── Right: animated tagline ──
        self.tagline_lbl = tk.Label(hdr, text="✦  MacColl's PDF Splitter  ✦",
                                    bg=C["panel"], fg=C["gold"],
                                    font=("Courier New",11,"bold"))
        self.tagline_lbl.pack(side="right", padx=20)

        # ── Right: decoration bars ──
        dbar = tk.Canvas(hdr, width=6, height=56, bg=C["panel"], highlightthickness=0)
        dbar.pack(side="right", padx=(0,6), pady=6)
        for i,cl in enumerate([C["gold"],C["gold_lt"],C["gold"],C["gold_dk"],C["gold"]]):
            dbar.create_rectangle(0,i*11,6,i*11+9, fill=cl, outline="")

    # ── Left: PDF viewer ──────────────────────────────────────────────────────
    def _build_viewer(self, parent):
        vf = tk.Frame(parent, bg=C["surface"], width=384)
        vf.pack(side="left", fill="y", padx=(0,5))
        vf.pack_propagate(False)

        # Sub-header
        vh = tk.Frame(vf, bg=C["panel2"])
        vh.pack(fill="x")
        tk.Label(vh, text="PDF VIEWER", bg=C["panel2"], fg=C["gold"],
                 font=("Courier New",9,"bold")).pack(side="left",padx=10,pady=5)
        self.viewer_info = tk.Label(vh, text="No document",
                                    bg=C["panel2"], fg=C["text_dim"],
                                    font=("Courier New",8))
        self.viewer_info.pack(side="right", padx=8)

        # Canvas + scrollbar
        cf = tk.Frame(vf, bg=C["vbg"])
        cf.pack(fill="both", expand=True)
        self.vcv = tk.Canvas(cf, bg=C["vbg"], highlightthickness=0, cursor="crosshair")
        vsb = ttk.Scrollbar(cf, orient="vertical", command=self.vcv.yview)
        self.vcv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.vcv.pack(side="left", fill="both", expand=True)
        self.vcv.bind("<MouseWheel>", lambda e: self.vcv.yview_scroll(int(-1*(e.delta/120)),"units"))
        self.vcv.bind("<Button-4>",   lambda e: self.vcv.yview_scroll(-1,"units"))
        self.vcv.bind("<Button-5>",   lambda e: self.vcv.yview_scroll(1,"units"))

        # Render progress
        self.rpbar = ttk.Progressbar(vf, mode="determinate", style="gold.Horizontal.TProgressbar")
        self.rpbar.pack_forget()

    # ── Right panel ──────────────────────────────────────────────────────────
    def _build_right(self, parent):
        rp = tk.Frame(parent, bg=C["bg"])
        rp.pack(side="left", fill="both", expand=True)
        self._build_topbar(rp)
        self._build_table(rp)
        self._build_edit_panel(rp)
        self._build_output_bar(rp)

    def _build_topbar(self, parent):
        bar = tk.Frame(parent, bg=C["panel"])
        bar.pack(fill="x", pady=(0,4))

        # File controls
        self._btn(bar,"📂  Open PDF", self._open_pdf,"gold").pack(side="left",padx=(8,6),pady=6)
        self.file_lbl = tk.Label(bar, text="No file selected",
                                 bg=C["panel"], fg=C["text_dim"],
                                 font=("Courier New",9), width=32, anchor="w")
        self.file_lbl.pack(side="left")
        self.pages_lbl = tk.Label(bar, text="", bg=C["panel"], fg=C["success"],
                                  font=("Courier New",9,"bold"))
        self.pages_lbl.pack(side="left",padx=(4,12))

        tk.Frame(bar, bg=C["border"], width=1).pack(side="left", fill="y", pady=4)

        # Detect mode
        self.dmode = tk.StringVar(value="heuristic")
        for lbl,val in [("Heuristic","heuristic"),("AI — Claude","ai")]:
            tk.Radiobutton(bar,text=lbl,variable=self.dmode,value=val,
                           bg=C["panel"],fg=C["text"],selectcolor=C["panel2"],
                           activebackground=C["panel"],activeforeground=C["gold"],
                           font=("Courier New",9)).pack(side="left",padx=4)
        self._btn(bar,"🔍  Detect",self._run_detect,"accent").pack(side="left",padx=(8,4),pady=6)

        tk.Frame(bar, bg=C["border"], width=1).pack(side="left", fill="y", pady=4)

        self._btn(bar,"＋ Add",    self._add_segment).pack(side="left",padx=3,pady=6)
        self._btn(bar,"✕ Remove",  self._remove_sel,"danger").pack(side="left",padx=3,pady=6)
        self._btn(bar,"⟳ Reset",   self._reset,"dim").pack(side="left",padx=3,pady=6)

        # API key far right
        tk.Label(bar,text="API Key:",bg=C["panel"],fg=C["text_dim"],
                 font=("Courier New",8)).pack(side="right",padx=(0,4))
        self.api_var = tk.StringVar()
        tk.Entry(bar,textvariable=self.api_var,show="*",width=22,
                 bg=C["entry"],fg=C["text"],insertbackground=C["gold"],
                 font=("Courier New",8),relief="flat",bd=3).pack(side="right",padx=(0,8),pady=6)

    def _build_table(self, parent):
        tf = tk.Frame(parent, bg=C["bg"])
        tf.pack(fill="both", expand=True)

        # Columns: (id, header, width, anchor)
        self.COLS = [
            ("num",   "#",            38,  "center"),
            ("name",  "Document Name",195, "w"),
            ("start", "Start Page",   76,  "center"),
            ("end",   "End Page",     72,  "center"),
            ("pages", "Pages",        52,  "center"),
            ("date",  "Est. Date",    92,  "center"),
            ("source","Source",       76,  "center"),
            ("conf",  "Conf.",        52,  "center"),
            ("reason","Reason",       280, "w"),
        ]
        cids = [c[0] for c in self.COLS]

        self.tree = ttk.Treeview(tf, columns=cids, show="headings",
                                 style="MacColl.Treeview", selectmode="browse")
        for cid,hdr,w,anc in self.COLS:
            self.tree.heading(cid, text=hdr,
                              command=lambda c=cid,h=hdr: self._sort(c,h))
            self.tree.column(cid, width=w, minwidth=28, anchor=anc,
                             stretch=(cid in ("name","reason")))

        ys = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        xs = ttk.Scrollbar(tf, orient="horizontal",  command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom",fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.tag_configure("even", background=C["row_a"], foreground=C["text"])
        self.tree.tag_configure("odd",  background=C["row_b"], foreground=C["text"])
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def _build_edit_panel(self, parent):
        ep = tk.Frame(parent, bg=C["panel"])
        ep.pack(fill="x", pady=(4,3))

        tk.Label(ep,text="EDIT SELECTED SEGMENT",bg=C["panel"],fg=C["gold"],
                 font=("Courier New",8,"bold")).grid(row=0,column=0,columnspan=10,
                                                     sticky="w",padx=10,pady=(5,2))

        fields=[("Name:",self.ev_name,28),("Start:",self.ev_start,6),
                ("End:",self.ev_end,6),("Est. Date:",self.ev_date,12)]
        for col,(lbl,var,w) in enumerate(fields):
            tk.Label(ep,text=lbl,bg=C["panel"],fg=C["text_dim"],
                     font=("Courier New",8)).grid(row=1,column=col*2,padx=(10,2),pady=4,sticky="e")
            tk.Entry(ep,textvariable=var,width=w,bg=C["entry"],fg=C["text"],
                     insertbackground=C["gold"],font=("Courier New",9),
                     relief="flat",bd=3).grid(row=1,column=col*2+1,padx=(0,8),pady=4,sticky="w")

        self._btn(ep,"✔  Apply",self._apply_edit,"success").grid(row=1,column=8,padx=4,pady=4)
        self._btn(ep,"✕  Cancel",self._cancel_edit,"dim").grid(row=1,column=9,padx=(0,10),pady=4)
        tk.Label(ep,text="← Click any row to edit",bg=C["panel"],fg=C["text_mute"],
                 font=("Courier New",8,"italic")).grid(row=1,column=10,padx=8,sticky="w")

    def _build_output_bar(self, parent):
        bar = tk.Frame(parent, bg=C["panel"])
        bar.pack(fill="x")
        tk.Label(bar,text="Output folder:",bg=C["panel"],fg=C["text_dim"],
                 font=("Courier New",9)).pack(side="left",padx=(10,4),pady=7)
        self.out_var = tk.StringVar(value=str(Path.home()/"Downloads"))
        tk.Entry(bar,textvariable=self.out_var,width=42,bg=C["entry"],fg=C["text"],
                 insertbackground=C["gold"],font=("Courier New",9),
                 relief="flat",bd=3).pack(side="left",pady=7)
        self._btn(bar,"…",self._choose_out).pack(side="left",padx=4,pady=7)
        self._btn(bar,"✂   Extract All Documents",self._extract,"gold").pack(
            side="right",padx=12,pady=7)

    def _build_statusbar(self):
        sb = tk.Frame(self.root, bg=C["panel2"], height=24)
        sb.pack(fill="x",side="bottom"); sb.pack_propagate(False)
        self.svar = tk.StringVar(value="Ready — Open a PDF to begin.")
        tk.Label(sb,textvariable=self.svar,bg=C["panel2"],fg=C["text_dim"],
                 font=("Courier New",8),anchor="w").pack(side="left",padx=10,pady=3)
        tk.Label(sb,text="MacColl's PDF Splitter  ✦  v2.0",
                 bg=C["panel2"],fg=C["gold_dk"],
                 font=("Courier New",8)).pack(side="right",padx=12)

    # ── ttk styles ────────────────────────────────────────────────────────────
    def _ttk_styles(self):
        s = ttk.Style(); s.theme_use("default")
        s.configure("MacColl.Treeview",
                    background=C["row_a"], foreground=C["text"],
                    fieldbackground=C["row_a"], rowheight=23,
                    font=("Courier New",9))
        s.configure("MacColl.Treeview.Heading",
                    background=C["panel2"], foreground=C["gold"],
                    font=("Courier New",9,"bold"), relief="flat")
        s.map("MacColl.Treeview",
              background=[("selected",C["sel"])],
              foreground=[("selected",C["text"])])
        s.map("MacColl.Treeview.Heading",
              background=[("active",C["gold_dk"])])
        s.configure("Vertical.TScrollbar",
                    background=C["panel"], troughcolor=C["bg"],
                    arrowcolor=C["text_dim"])
        s.configure("Horizontal.TScrollbar",
                    background=C["panel"], troughcolor=C["bg"],
                    arrowcolor=C["text_dim"])
        s.configure("gold.Horizontal.TProgressbar",
                    background=C["gold"], troughcolor=C["panel2"])

    # ── Animation ─────────────────────────────────────────────────────────────
    def _start_anim(self):
        self._msgs = [
            "✦  MacColl's PDF Splitter  ✦",
            "◈  AI-Powered Detection  ◈",
            "⬡  Split · Extract · Own It  ⬡",
            "✦  Intelligent Boundaries  ✦",
            "◈  Document Intelligence  ◈",
            "⬡  MacColl's PDF Splitter  ⬡",
        ]
        self._clrs = [
            C["gold"],C["gold_lt"],"#ffffff",C["gold_pale"],
            C["gold_lt"],C["gold"],C["gold_dk"],C["gold"],
            C["gold_lt"],"#ffffff",C["gold_pale"],C["gold"],
        ]
        self._tick()

    def _tick(self):
        step=self._anim_step
        self.tagline_lbl.config(fg=self._clrs[step % len(self._clrs)])
        if step % 30 == 0:
            self._anim_msg = (self._anim_msg+1) % len(self._msgs)
            self.tagline_lbl.config(text=self._msgs[self._anim_msg])
        self._anim_step+=1
        self.root.after(110, self._tick)

    # ── Helper: button factory ────────────────────────────────────────────────
    def _btn(self, parent, text, cmd, style="normal", **kw):
        pal={"normal":(C["panel"],C["text"]),"gold":(C["gold"],C["bg"]),
             "accent":(C["accent"],C["bg"]),"success":(C["success"],C["bg"]),
             "danger":(C["danger"],C["text"]),"dim":(C["border"],C["text_dim"])}
        bg,fg=pal.get(style,pal["normal"])
        return tk.Button(parent,text=text,command=cmd,bg=bg,fg=fg,
                         activebackground=fg,activeforeground=bg,
                         font=("Courier New",8,"bold"),relief="flat",
                         cursor="hand2",pady=3,padx=7,**kw)

    def _log(self,msg): self.svar.set(msg)

    # ── Open PDF ──────────────────────────────────────────────────────────────
    def _open_pdf(self):
        path=filedialog.askopenfilename(
            title="Select Combined PDF",
            filetypes=[("PDF files","*.pdf"),("All files","*.*")])
        if not path: return
        self.pdf_path=path
        name=Path(path).name
        try:
            r=PdfReader(path)
            self.total_pages=len(r.pages)
            self.file_lbl.config(text=name[:36],fg=C["text"])
            self.pages_lbl.config(text=f"({self.total_pages} pages)")
            self.out_var.set(str(Path(path).parent))
            self._log(f"Loaded: {name}  —  {self.total_pages} pages")
            self._render_viewer()
        except Exception as e:
            messagebox.showerror("Error",f"Could not open PDF:\n{e}")

    # ── PDF Viewer rendering ──────────────────────────────────────────────────
    def _render_viewer(self):
        if not (HAS_FITZ and HAS_PIL):
            self.viewer_info.config(text="pip install pymupdf Pillow for preview")
            return
        self.page_images.clear(); self.page_y.clear()
        self.vcv.delete("all")
        self.rpbar.pack(fill="x",padx=6,pady=3)
        self.rpbar["maximum"]=self.total_pages
        self.rpbar["value"]=0

        def run():
            try:
                doc=fitz.open(self.pdf_path)
                # Determine scale to fit viewer width
                vw=self.vcv.winfo_width() or 360
                scale=max(0.38,min((vw-18)/612.0,0.95))
                y_off=8
                for pn in range(self.total_pages):
                    pg=doc[pn]
                    pix=pg.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False)
                    img=Image.open(io.BytesIO(pix.tobytes("ppm")))
                    photo=ImageTk.PhotoImage(img)
                    self.page_images[pn+1]=photo
                    self.page_y[pn+1]=y_off
                    # Capture loop vars
                    _pn,_y,_w,_h=pn+1,y_off,img.width,img.height
                    def draw(_p=_pn,_yy=_y,_ww=_w,_hh=_h):
                        self.vcv.create_image(9,_yy,anchor="nw",
                                              image=self.page_images.get(_p))
                        # Page badge
                        self.vcv.create_rectangle(9,_yy,9+_ww,_yy+_hh,
                                                  outline=C["border"],width=1)
                        self.vcv.create_rectangle(9,_yy+_hh-17,50,_yy+_hh,
                                                  fill=C["panel"],outline="")
                        self.vcv.create_text(29,_yy+_hh-8,text=str(_p),
                                             fill=C["text_dim"],
                                             font=("Courier New",7))
                        self.vcv.configure(
                            scrollregion=self.vcv.bbox("all") or (0,0,400,600))
                        self.rpbar["value"]=_p
                    self.root.after(0,draw)
                    y_off+=img.height+6
                doc.close()
                self.root.after(200,lambda:(
                    self.rpbar.pack_forget(),
                    self.viewer_info.config(
                        text=f"{self.total_pages} pages  |  scroll to browse")))
            except Exception as ex:
                self.root.after(0,lambda:self._log(f"Viewer error: {ex}"))

        threading.Thread(target=run,daemon=True).start()

    def _color_viewer_overlays(self):
        """Draw colored left-edge bars showing which segment each page belongs to."""
        self.vcv.delete("seg_overlay")
        for si,seg in enumerate(self.segments):
            col=SEG_PALETTE[si % len(SEG_PALETTE)]
            for pn in range(seg["start"],seg["end"]+1):
                if pn in self.page_y:
                    y=self.page_y[pn]
                    h=self.page_images[pn].height() if pn in self.page_images else 50
                    self.vcv.create_rectangle(0,y,7,y+h,
                                              fill=col,outline="",tags="seg_overlay")

    def _jump_to(self,page_num):
        if page_num not in self.page_y: return
        bbox=self.vcv.bbox("all")
        if not bbox: return
        total_h=bbox[3]
        if total_h>0:
            frac=self.page_y[page_num]/total_h
            self.vcv.yview_moveto(max(0,frac-0.02))

    # ── Table refresh ─────────────────────────────────────────────────────────
    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for idx,seg in enumerate(self.segments):
            pages=seg["end"]-seg["start"]+1
            tag="even" if idx%2==0 else "odd"
            self.tree.insert("","end",iid=str(idx),tags=(tag,),values=(
                idx+1,seg["name"],seg["start"],seg["end"],pages,
                seg.get("date",""),seg.get("source","?"),
                f"{seg.get('confidence',0):.0%}",seg.get("reason",""),
            ))
        self._color_viewer_overlays()

    # ── Sort ──────────────────────────────────────────────────────────────────
    def _sort(self, col, base_hdr):
        rev = (self._sort_col==col and not self._sort_rev)
        self._sort_col=col; self._sort_rev=rev
        data=[(self.tree.set(k,col),k) for k in self.tree.get_children("")]
        NUM={"start","end","pages"}
        try:
            if col in NUM:
                data.sort(key=lambda t:int(t[0] or 0),reverse=rev)
            elif col=="conf":
                data.sort(key=lambda t:float(t[0].rstrip('%') or 0),reverse=rev)
            elif col=="date":
                def dkey(s):
                    try: return datetime.strptime(s,"%m/%d/%Y")
                    except: return datetime.min
                data.sort(key=lambda t:dkey(t[0]),reverse=rev)
            else:
                data.sort(key=lambda t:t[0].lower(),reverse=rev)
        except: data.sort(reverse=rev)
        for i,(_,k) in enumerate(data):
            self.tree.move(k,"",i)
            self.tree.item(k,tags=("even" if i%2==0 else "odd",))
        # Update heading indicators
        for cid,hdr,_,_ in self.COLS:
            arrow=(" ↑" if not rev else " ↓") if cid==col else ""
            self.tree.heading(cid,text=hdr+arrow,
                              command=lambda c=cid,h=hdr:self._sort(c,h))

    # ── Tree selection ────────────────────────────────────────────────────────
    def _on_select(self, _event):
        sel=self.tree.selection()
        if not sel: return
        idx=int(sel[0])
        if idx>=len(self.segments): return
        self._sel_idx=idx
        seg=self.segments[idx]
        self.ev_name.set(seg["name"])
        self.ev_start.set(str(seg["start"]))
        self.ev_end.set(str(seg["end"]))
        self.ev_date.set(seg.get("date",""))
        self._jump_to(seg["start"])

    # ── Segment actions ───────────────────────────────────────────────────────
    def _run_detect(self):
        if not self.pdf_path:
            messagebox.showwarning("No PDF","Open a PDF first."); return
        def worker():
            self.root.after(0,lambda:self._log("Extracting page text…"))
            self.page_texts=extract_pages(self.pdf_path)
            mode=self.dmode.get()
            if mode=="heuristic":
                segs=heuristic_detect(self.page_texts)
            else:
                key=self.api_var.get().strip() or None
                segs=ai_detect(self.page_texts,api_key=key,
                               progress_cb=lambda m:self.root.after(0,
                               lambda msg=m:self._log(msg)))
            self.root.after(0,lambda:self._done_detect(segs))
        threading.Thread(target=worker,daemon=True).start()

    def _done_detect(self, segs):
        self.segments=segs; self._sel_idx=None
        self._refresh()
        self._log(f"Detected {len(segs)} document(s). Review & adjust, then Extract.")

    def _add_segment(self):
        if not self.total_pages:
            messagebox.showwarning("No PDF","Open a PDF first."); return
        n=len(self.segments)+1
        start=(self.segments[-1]["end"]+1) if self.segments else 1
        start=min(start,self.total_pages)
        seg={"start":start,"end":self.total_pages,"name":f"Exhibit_{n}",
             "date":"","confidence":1.0,"reason":"Manually added","source":"manual"}
        self.segments.append(seg)
        self._refresh()
        self._sel_idx=len(self.segments)-1
        self.tree.selection_set(str(self._sel_idx))
        self.ev_name.set(seg["name"]); self.ev_start.set(str(seg["start"]))
        self.ev_end.set(str(seg["end"])); self.ev_date.set("")

    def _remove_sel(self):
        if self._sel_idx is None:
            messagebox.showinfo("Select","Click a segment row first."); return
        self.segments.pop(self._sel_idx)
        # Renumber Exhibit_N names
        for n,s in enumerate(self.segments,1):
            if re.match(r'^Exhibit_\d+$',s["name"]):
                s["name"]=f"Exhibit_{n}"
        self._sel_idx=None; self._refresh()

    def _reset(self):
        if messagebox.askyesno("Reset","Clear all segments?"):
            self.segments=[]; self._sel_idx=None
            self._refresh(); self._log("Segments cleared.")

    def _apply_edit(self):
        if self._sel_idx is None: return
        try:
            s,e=int(self.ev_start.get()),int(self.ev_end.get())
            if not(1<=s<=e<=self.total_pages): raise ValueError
        except:
            messagebox.showerror("Invalid",
                f"Pages must be between 1 and {self.total_pages},\nand start ≤ end.")
            return
        self.segments[self._sel_idx].update({
            "name":  self.ev_name.get().strip() or f"Exhibit_{self._sel_idx+1}",
            "start": s, "end": e,
            "date":  self.ev_date.get().strip(),
            "source":"manual",
        })
        self._refresh(); self._log("Segment updated.")

    def _cancel_edit(self):
        self._sel_idx=None
        self.tree.selection_remove(*self.tree.selection())

    def _choose_out(self):
        d=filedialog.askdirectory(title="Choose Output Folder")
        if d: self.out_var.set(d)

    def _extract(self):
        if not self.pdf_path:
            messagebox.showwarning("No PDF","Open a PDF first."); return
        if not self.segments:
            messagebox.showwarning("No Segments","Detect or add segments first."); return
        out=self.out_var.get().strip()
        if not out:
            messagebox.showwarning("No Output","Set an output folder."); return
        os.makedirs(out,exist_ok=True)
        def worker():
            try:
                files=split_pdf(self.pdf_path,self.segments,out,
                    progress_cb=lambda m:self.root.after(0,lambda msg=m:self._log(msg)))
                self.root.after(0,lambda:self._done_extract(files,out))
            except Exception as ex:
                self.root.after(0,lambda:messagebox.showerror("Error",str(ex)))
        threading.Thread(target=worker,daemon=True).start()

    def _done_extract(self, files, out):
        self._log(f"✅  {len(files)} files extracted  →  {out}")
        if messagebox.askyesno("Done!",f"{len(files)} PDFs saved to:\n{out}\n\nOpen folder?"):
            import subprocess, platform
            try:
                if   platform.system()=="Darwin":  subprocess.Popen(["open",out])
                elif platform.system()=="Windows": os.startfile(out)
                else:                              subprocess.Popen(["xdg-open",out])
            except: pass


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    if not HAS_PYPDF:
        sys.exit("ERROR: pypdf is required.\nRun:  pip install pypdf")
    if not (HAS_FITZ and HAS_PIL):
        print("NOTE: PDF preview requires PyMuPDF + Pillow.\n"
              "Run:  pip install pymupdf Pillow")
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__=="__main__":
    main()
