#!/usr/bin/env python3
"""
Case Command  —  Multi-Case Litigation Management Tool
=======================================================
Manages multiple cases from a single Case Master view.
Each case has its own database: docket, court hearings & deadlines,
written discovery (set-based), depositions, parties & counsel.
Imports PACER docket PDFs automatically.  Full backup/restore support.

HOW TO RUN (Windows / PowerShell):
    python Case_Command.py

DEPENDENCIES (one-time install):
    python -m pip install pymupdf Pillow
"""

import os, sys, re, csv, sqlite3, json, zipfile, shutil, io, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime, date, timedelta

try:
    import fitz
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install pymupdf")
try:
    from PIL import Image, ImageTk
except ImportError:
    sys.exit("Missing dependency.\nRun: python -m pip install Pillow")

# ── palette ───────────────────────────────────────────────────────────────────
C = dict(
    bg='#1e1e2e', bg2='#181825', bg3='#313244', bg4='#45475a',
    fg='#cdd6f4',  fg2='#a6adc8', fg3='#6c7086',
    acc='#89b4fa', green='#a6e3a1', yellow='#f9e2af',
    red='#f38ba8', orange='#fab387', mauve='#cba6f7',
)

DB_FILE  = None   # legacy stub — per-case paths managed by CaseMasterDB
PDF_H    = 720   # target preview height in pixels
TODAY    = date.today

# ── tiny UI helpers ───────────────────────────────────────────────────────────
def _L(p, text, size=9, bold=False, fg=None):
    return tk.Label(p, text=text, bg=C['bg'], fg=fg or C['fg2'],
                    font=('Helvetica', size, 'bold' if bold else 'normal'),
                    anchor='w')

def _E(p, val='', w=None):
    kw = dict(bg=C['bg3'], fg=C['fg'], insertbackground=C['fg'],
              relief='flat', font=('Helvetica', 10))
    if w: kw['width'] = w
    e = tk.Entry(p, **kw)
    if val: e.insert(0, str(val))
    return e

def _T(p, val='', h=3):
    t = tk.Text(p, bg=C['bg3'], fg=C['fg'], insertbackground=C['fg'],
                relief='flat', font=('Helvetica', 10), height=h, wrap='word')
    if val: t.insert('1.0', str(val))
    return t

def _CB(p, values, val=''):
    c = ttk.Combobox(p, values=values, state='readonly',
                     font=('Helvetica', 10))
    if val in values: c.set(val)
    elif values:      c.set(values[0])
    return c

def _Btn(p, text, cmd, bg=None, fg=None, **kw):
    return tk.Button(p, text=text, command=cmd,
                     bg=bg or C['bg3'], fg=fg or C['fg'], relief='flat',
                     font=('Helvetica', 9, 'bold'), padx=8, pady=3,
                     cursor='hand2',
                     activebackground=C['bg4'],
                     activeforeground=C['fg'], **kw)

def _tv_style(name):
    s = ttk.Style()
    s.configure(f'{name}.Treeview',
                background=C['bg2'], foreground=C['fg'],
                fieldbackground=C['bg2'], rowheight=22,
                font=('Helvetica', 9))
    s.configure(f'{name}.Treeview.Heading',
                background=C['bg3'], foreground=C['acc'],
                font=('Helvetica', 9, 'bold'))
    s.map(f'{name}.Treeview', background=[('selected', C['bg3'])])
    return f'{name}.Treeview'

# ── Sortable Treeview helper ─────────────────────────────────────────────────
# Click a column heading to sort rows alphanumerically by that column.
# Click the same heading again to reverse the direction. A small arrow
# indicator (▲ / ▼) is appended to the active column's heading text.
#
# Sort state is stored on the Treeview instance itself (as `_sort_state`)
# so it survives refreshes: after repopulating the tree, call
# `_apply_sort(tv)` to re-apply the active sort.
def _sort_key(val):
    """Return a key suitable for natural alphanumeric sorting.
    - Recognises MM/DD/YYYY dates and sorts them chronologically.
    - Recognises pure numbers (including with ECF-style hyphens like '12-3').
    - Falls back to case-insensitive string comparison.
    """
    s = ('' if val is None else str(val)).strip()
    if not s:
        # Empty strings sort last (both asc and desc) by using a sentinel tuple
        return (2, '')
    # Try date in MM/DD/YYYY
    try:
        d = datetime.strptime(s, '%m/%d/%Y')
        return (0, d.timestamp())
    except ValueError:
        pass
    # Try date in YYYY-MM-DD
    try:
        d = datetime.strptime(s, '%Y-%m-%d')
        return (0, d.timestamp())
    except ValueError:
        pass
    # Try plain number (allow leading/trailing text? keep simple: pure numeric)
    try:
        return (0, float(s))
    except ValueError:
        pass
    # ECF-style "12-3" -> (12, 3) tuple of ints, preserves hierarchy
    if '-' in s:
        parts = s.split('-')
        if all(p.strip().isdigit() for p in parts):
            return (0, tuple(int(p) for p in parts))
    # Fallback: case-insensitive string
    return (1, s.lower())

def _apply_sort(tv):
    """Sort the rows currently in the Treeview according to its stored
    `_sort_state` (a dict with keys 'col' and 'desc'). If no sort is active,
    does nothing. For trees with children (e.g. the docket's parent/child
    rows), only top-level rows are reordered; children stay in order under
    their parent."""
    state = getattr(tv, '_sort_state', None)
    if not state or not state.get('col'):
        _update_sort_arrows(tv)
        return
    col  = state['col']
    desc = state['desc']
    # Only reorder top-level items. Children keep their relative order under
    # their parent (so docket attachments stay nested under their parent entry).
    top_items = list(tv.get_children(''))
    try:
        col_idx = list(tv['columns']).index(col)
    except ValueError:
        return

    def _get_val(iid):
        vals = tv.item(iid, 'values')
        return vals[col_idx] if col_idx < len(vals) else ''

    top_items.sort(key=lambda iid: _sort_key(_get_val(iid)), reverse=desc)
    for i, iid in enumerate(top_items):
        tv.move(iid, '', i)
    _update_sort_arrows(tv)

def _update_sort_arrows(tv):
    """Refresh the ▲/▼ indicator shown on column headings."""
    state     = getattr(tv, '_sort_state', None) or {}
    active    = state.get('col')
    desc      = state.get('desc', False)
    base      = getattr(tv, '_sort_base_headings', {})
    for col in tv['columns']:
        text = base.get(col, col)
        if col == active:
            text = f'{text}  {"▼" if desc else "▲"}'
        tv.heading(col, text=text)

def _make_sortable(tv):
    """Attach click-to-sort behaviour to every column heading of `tv`.
    Stores the original heading text (so the arrow can be appended without
    growing) and a `_sort_state` dict on the widget."""
    # Capture the original heading texts BEFORE we start appending arrows,
    # so toggling doesn't accumulate '▲▲▲' over repeated clicks.
    tv._sort_base_headings = {col: tv.heading(col)['text'] for col in tv['columns']}
    tv._sort_state         = {'col': None, 'desc': False}

    def _on_click(col):
        st = tv._sort_state
        if st['col'] == col:
            st['desc'] = not st['desc']   # toggle direction
        else:
            st['col']  = col
            st['desc'] = False            # first click = ascending
        _apply_sort(tv)

    for col in tv['columns']:
        # NOTE: bind a *copy* of `col` via default-arg trick
        tv.heading(col, command=lambda c=col: _on_click(c))

def iso(d):
    """Display date string -> ISO YYYY-MM-DD, tolerant of multiple formats."""
    if not d: return ''
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%Y/%m/%d'):
        try: return datetime.strptime(d.strip(), fmt).strftime('%Y-%m-%d')
        except: pass
    return d.strip()

def disp(d):
    """ISO YYYY-MM-DD -> MM/DD/YYYY for display."""
    if not d: return ''
    try: return datetime.strptime(d, '%Y-%m-%d').strftime('%m/%d/%Y')
    except: return d

def deadline_tag(due_iso):
    """Return treeview tag based on deadline urgency."""
    if not due_iso: return 'ok'
    try:
        dt = datetime.strptime(due_iso, '%Y-%m-%d').date()
        if dt < TODAY():              return 'overdue'
        if dt <= TODAY() + timedelta(14): return 'soon'
    except: pass
    return 'ok'

# ═════════════════════════════════════════════════════════════════════════════
# PATH CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════
_SCRIPT_DIR  = Path(__file__).parent
MASTER_DB    = _SCRIPT_DIR / 'case_master.db'
DEFAULT_CASES_DIR = _SCRIPT_DIR / 'cases'

# ═════════════════════════════════════════════════════════════════════════════
# CASE MASTER DATABASE  (registry of all cases)
# ═════════════════════════════════════════════════════════════════════════════
class CaseMasterDB:
    def __init__(self):
        self.cx = sqlite3.connect(str(MASTER_DB))
        self.cx.row_factory = sqlite3.Row
        self.cx.execute("PRAGMA journal_mode=WAL")
        self._init()

    def _init(self):
        self.cx.executescript("""
        CREATE TABLE IF NOT EXISTS cases(
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            case_name   TEXT    DEFAULT '',
            case_number TEXT    DEFAULT '',
            court       TEXT    DEFAULT '',
            judge       TEXT    DEFAULT '',
            db_path     TEXT    UNIQUE NOT NULL,
            created_at  TEXT,
            last_opened TEXT,
            status      TEXT    DEFAULT 'Active',
            notes       TEXT    DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
        """)
        self.cx.commit()

    def add_case(self, case_name, case_number, court, judge, db_path):
        now = datetime.now().isoformat()
        self.cx.execute(
            "INSERT OR IGNORE INTO cases(case_name,case_number,court,judge,"
            "db_path,created_at,last_opened) VALUES(?,?,?,?,?,?,?)",
            (case_name, case_number, court, judge, str(db_path), now, now))
        self.cx.commit()
        return self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]

    def all_cases(self):
        return self.cx.execute(
            "SELECT * FROM cases ORDER BY last_opened DESC").fetchall()

    def get_case(self, id):
        return self.cx.execute(
            "SELECT * FROM cases WHERE id=?", (id,)).fetchone()

    def sync_meta(self, id, meta: dict):
        """Push updated case info from a CaseDB.meta() back into master."""
        self.cx.execute(
            "UPDATE cases SET case_name=?,case_number=?,court=?,judge=? WHERE id=?",
            (meta.get('case_name',''), meta.get('case_number',''),
             meta.get('court',''),     meta.get('judge',''), id))
        self.cx.commit()

    def touch(self, id):
        self.cx.execute("UPDATE cases SET last_opened=? WHERE id=?",
                        (datetime.now().isoformat(), id))
        self.cx.commit()

    def set_status(self, id, status):
        self.cx.execute("UPDATE cases SET status=? WHERE id=?", (status, id))
        self.cx.commit()

    def remove(self, id):
        self.cx.execute("DELETE FROM cases WHERE id=?", (id,))
        self.cx.commit()

    def get_setting(self, k, default=''):
        r = self.cx.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
        return r['v'] if r else default

    def set_setting(self, k, v):
        self.cx.execute("INSERT OR REPLACE INTO settings VALUES(?,?)", (k, v))
        self.cx.commit()

    def close(self): self.cx.close()

# ═════════════════════════════════════════════════════════════════════════════
# DATABASE  (per-case)
# ═════════════════════════════════════════════════════════════════════════════
class CaseDB:
    def __init__(self, path=None):
        if path is None:
            # Legacy fallback — creates in script dir (shouldn't be hit in normal use)
            path = _SCRIPT_DIR / 'case_command.db'
        self.path = Path(path)
        self.cx = sqlite3.connect(str(self.path))
        self.cx.row_factory = sqlite3.Row
        self.cx.execute("PRAGMA journal_mode=WAL")
        self._init()

    def _init(self):
        self.cx.executescript("""
        CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);

        CREATE TABLE IF NOT EXISTS docket(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_no TEXT, date_filed TEXT, description TEXT,
            doc_type TEXT, filing_party TEXT, file_path TEXT, notes TEXT,
            parent_id INTEGER REFERENCES docket(id),
            attachment_no TEXT DEFAULT '');

        CREATE TABLE IF NOT EXISTS deadlines(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, dtype TEXT, due_date TEXT,
            status TEXT DEFAULT 'Upcoming',
            docket_ref TEXT, notes TEXT);

        CREATE TABLE IF NOT EXISTS disc_sets(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rtype TEXT, set_no INTEGER, direction TEXT,
            prop_party TEXT, resp_party TEXT,
            date_served TEXT, due_date TEXT, resp_date TEXT,
            num_reqs INTEGER DEFAULT 0,
            status TEXT DEFAULT 'Pending',
            deficiency TEXT DEFAULT 'None',
            notes TEXT);

        CREATE TABLE IF NOT EXISTS depos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            witness TEXT, role TEXT, affiliation TEXT,
            notice_date TEXT, depo_date TEXT, location TEXT,
            format TEXT, status TEXT DEFAULT 'Not Noticed',
            errata_deadline TEXT, transcript_bates TEXT,
            topics TEXT, issues TEXT, notes TEXT);

        CREATE TABLE IF NOT EXISTS parties(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, role TEXT, affiliation TEXT,
            contact TEXT, notes TEXT);

        CREATE TABLE IF NOT EXISTS counsel(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, firm TEXT, phone TEXT, email TEXT,
            party_id INTEGER, is_lead INTEGER DEFAULT 0,
            term_date TEXT, notes TEXT);
        """)
        self.cx.commit()
        # ── Migration: add columns that post-date the original schema ─────────
        for col, defn in [('parent_id',     'INTEGER'),
                          ('attachment_no', "TEXT DEFAULT ''")]:
            try:
                self.cx.execute(f'ALTER TABLE docket ADD COLUMN {col} {defn}')
                self.cx.commit()
            except Exception:
                pass   # column already exists — safe to ignore

    # ── meta ──────────────────────────────────────────────────────────────────
    def meta(self):
        return {r['k']: r['v'] for r in
                self.cx.execute("SELECT k,v FROM meta").fetchall()}
    def set_meta(self, d):
        for k, v in d.items():
            self.cx.execute("INSERT OR REPLACE INTO meta VALUES(?,?)", (k, v))
        self.cx.commit()

    # ── docket ────────────────────────────────────────────────────────────────
    def docket_all(self):
        return self.cx.execute(
            "SELECT * FROM docket ORDER BY date_filed,"
            " CAST(entry_no AS INT), CAST(attachment_no AS INT)").fetchall()
    def docket_add(self, **r):
        r.setdefault('parent_id', None)
        r.setdefault('attachment_no', '')
        self.cx.execute(
            "INSERT INTO docket(entry_no,date_filed,description,doc_type,"
            "filing_party,file_path,notes,parent_id,attachment_no) "
            "VALUES(:entry_no,:date_filed,:description,:doc_type,"
            ":filing_party,:file_path,:notes,:parent_id,:attachment_no)", r)
        self.cx.commit()
        return self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
    def docket_upd(self, id, **r):
        r['id'] = id
        r.setdefault('parent_id', None)
        r.setdefault('attachment_no', '')
        self.cx.execute(
            "UPDATE docket SET entry_no=:entry_no,date_filed=:date_filed,"
            "description=:description,doc_type=:doc_type,"
            "filing_party=:filing_party,file_path=:file_path,"
            "notes=:notes,parent_id=:parent_id,"
            "attachment_no=:attachment_no WHERE id=:id", r)
        self.cx.commit()
    def docket_del(self, id):
        self.cx.execute("DELETE FROM docket WHERE id=?", (id,)); self.cx.commit()
    def docket_set_path(self, id, file_path):
        self.cx.execute("UPDATE docket SET file_path=? WHERE id=?",
                        (str(file_path), id))
        self.cx.commit()
    def docket_existing_keys(self):
        return {(r['date_filed'], r['entry_no'])
                for r in self.cx.execute("SELECT date_filed,entry_no FROM docket")}

    # ── deadlines ─────────────────────────────────────────────────────────────
    def dl_all(self):
        return self.cx.execute(
            "SELECT * FROM deadlines ORDER BY due_date").fetchall()
    def dl_add(self, **r):
        self.cx.execute(
            "INSERT INTO deadlines(name,dtype,due_date,status,docket_ref,notes)"
            " VALUES(:name,:dtype,:due_date,:status,:docket_ref,:notes)", r)
        self.cx.commit()
    def dl_upd(self, id, **r):
        r['id'] = id
        self.cx.execute(
            "UPDATE deadlines SET name=:name,dtype=:dtype,due_date=:due_date,"
            "status=:status,docket_ref=:docket_ref,notes=:notes WHERE id=:id", r)
        self.cx.commit()
    def dl_del(self, id):
        self.cx.execute("DELETE FROM deadlines WHERE id=?", (id,)); self.cx.commit()

    # ── discovery sets ────────────────────────────────────────────────────────
    def disc_all(self):
        return self.cx.execute(
            "SELECT * FROM disc_sets ORDER BY rtype,set_no").fetchall()
    def disc_add(self, **r):
        self.cx.execute(
            "INSERT INTO disc_sets(rtype,set_no,direction,prop_party,resp_party,"
            "date_served,due_date,resp_date,num_reqs,status,deficiency,notes)"
            " VALUES(:rtype,:set_no,:direction,:prop_party,:resp_party,"
            ":date_served,:due_date,:resp_date,:num_reqs,:status,:deficiency,:notes)", r)
        self.cx.commit()
    def disc_upd(self, id, **r):
        r['id'] = id
        self.cx.execute(
            "UPDATE disc_sets SET rtype=:rtype,set_no=:set_no,"
            "direction=:direction,prop_party=:prop_party,resp_party=:resp_party,"
            "date_served=:date_served,due_date=:due_date,resp_date=:resp_date,"
            "num_reqs=:num_reqs,status=:status,deficiency=:deficiency,"
            "notes=:notes WHERE id=:id", r)
        self.cx.commit()
    def disc_del(self, id):
        self.cx.execute("DELETE FROM disc_sets WHERE id=?", (id,)); self.cx.commit()

    # ── depositions ───────────────────────────────────────────────────────────
    def depo_all(self):
        return self.cx.execute(
            "SELECT * FROM depos ORDER BY depo_date").fetchall()
    def depo_add(self, **r):
        self.cx.execute(
            "INSERT INTO depos(witness,role,affiliation,notice_date,depo_date,"
            "location,format,status,errata_deadline,transcript_bates,"
            "topics,issues,notes) VALUES(:witness,:role,:affiliation,"
            ":notice_date,:depo_date,:location,:format,:status,"
            ":errata_deadline,:transcript_bates,:topics,:issues,:notes)", r)
        self.cx.commit()
    def depo_upd(self, id, **r):
        r['id'] = id
        self.cx.execute(
            "UPDATE depos SET witness=:witness,role=:role,"
            "affiliation=:affiliation,notice_date=:notice_date,"
            "depo_date=:depo_date,location=:location,format=:format,"
            "status=:status,errata_deadline=:errata_deadline,"
            "transcript_bates=:transcript_bates,topics=:topics,"
            "issues=:issues,notes=:notes WHERE id=:id", r)
        self.cx.commit()
    def depo_del(self, id):
        self.cx.execute("DELETE FROM depos WHERE id=?", (id,)); self.cx.commit()

    # ── parties ───────────────────────────────────────────────────────────────
    def party_all(self):
        return self.cx.execute(
            "SELECT * FROM parties ORDER BY role,name").fetchall()
    def party_add(self, **r):
        self.cx.execute(
            "INSERT INTO parties(name,role,affiliation,contact,notes)"
            " VALUES(:name,:role,:affiliation,:contact,:notes)", r)
        self.cx.commit()
    def party_upd(self, id, **r):
        r['id'] = id
        self.cx.execute(
            "UPDATE parties SET name=:name,role=:role,affiliation=:affiliation,"
            "contact=:contact,notes=:notes WHERE id=:id", r)
        self.cx.commit()
    def party_del(self, id):
        self.cx.execute("DELETE FROM parties WHERE id=?", (id,)); self.cx.commit()

    # ── counsel ───────────────────────────────────────────────────────────────
    def counsel_all(self):
        return self.cx.execute("""
            SELECT c.*, p.name AS party_name FROM counsel c
            LEFT JOIN parties p ON c.party_id=p.id
            ORDER BY p.role,c.name""").fetchall()
    def counsel_add(self, **r):
        self.cx.execute(
            "INSERT INTO counsel(name,firm,phone,email,party_id,is_lead,"
            "term_date,notes) VALUES(:name,:firm,:phone,:email,:party_id,"
            ":is_lead,:term_date,:notes)", r)
        self.cx.commit()
    def counsel_upd(self, id, **r):
        r['id'] = id
        self.cx.execute(
            "UPDATE counsel SET name=:name,firm=:firm,phone=:phone,"
            "email=:email,party_id=:party_id,is_lead=:is_lead,"
            "term_date=:term_date,notes=:notes WHERE id=:id", r)
        self.cx.commit()
    def counsel_del(self, id):
        self.cx.execute("DELETE FROM counsel WHERE id=?", (id,)); self.cx.commit()

    # ── search ────────────────────────────────────────────────────────────────
    def search(self, q):
        w = f'%{q}%'; results = []
        for r in self.cx.execute(
                "SELECT id,entry_no,date_filed,description FROM docket "
                "WHERE description LIKE ? OR entry_no LIKE ? OR notes LIKE ?",
                (w,w,w)):
            results.append(('Docket', r['id'],
                f"#{r['entry_no']} {disp(r['date_filed'])} — {r['description'][:70]}"))
        for r in self.cx.execute(
                "SELECT id,name,due_date,status FROM deadlines "
                "WHERE name LIKE ? OR notes LIKE ?", (w,w)):
            results.append(('Deadlines', r['id'],
                f"{r['name']} ({disp(r['due_date'])}) — {r['status']}"))
        for r in self.cx.execute(
                "SELECT id,rtype,set_no,direction,prop_party FROM disc_sets "
                "WHERE prop_party LIKE ? OR resp_party LIKE ? OR notes LIKE ?",
                (w,w,w)):
            results.append(('Discovery', r['id'],
                f"{r['rtype']} Set {r['set_no']} {r['direction']} — {r['prop_party']}"))
        for r in self.cx.execute(
                "SELECT id,witness,depo_date,status FROM depos "
                "WHERE witness LIKE ? OR affiliation LIKE ? OR topics LIKE ?",
                (w,w,w)):
            results.append(('Depositions', r['id'],
                f"{r['witness']} {disp(r['depo_date'])} — {r['status']}"))
        for r in self.cx.execute(
                "SELECT id,name,role FROM parties "
                "WHERE name LIKE ? OR notes LIKE ?", (w,w)):
            results.append(('Parties', r['id'],
                f"{r['name']} ({r['role']})"))
        for r in self.cx.execute(
                "SELECT id,name,firm FROM counsel "
                "WHERE name LIKE ? OR firm LIKE ? OR email LIKE ?", (w,w,w)):
            results.append(('Parties', r['id'],
                f"{r['name']} — {r['firm']}"))
        return results

    # ── dashboard stats ───────────────────────────────────────────────────────
    def dashboard(self):
        td = TODAY().isoformat()
        sn = (TODAY() + timedelta(14)).isoformat()
        x  = self.cx
        return dict(
            dock_total   = x.execute("SELECT COUNT(*) FROM docket").fetchone()[0],
            dl_overdue   = x.execute("SELECT COUNT(*) FROM deadlines WHERE due_date<? AND status NOT IN('Completed','Vacated')",(td,)).fetchone()[0],
            dl_soon      = x.execute("SELECT COUNT(*) FROM deadlines WHERE due_date BETWEEN ? AND ? AND status NOT IN('Completed','Vacated')",(td,sn)).fetchone()[0],
            dl_total     = x.execute("SELECT COUNT(*) FROM deadlines").fetchone()[0],
            disc_total   = x.execute("SELECT COUNT(*) FROM disc_sets").fetchone()[0],
            disc_pending = x.execute("SELECT COUNT(*) FROM disc_sets WHERE status='Pending'").fetchone()[0],
            disc_deficient=x.execute("SELECT COUNT(*) FROM disc_sets WHERE deficiency NOT IN('None','Resolved')").fetchone()[0],
            depo_total   = x.execute("SELECT COUNT(*) FROM depos").fetchone()[0],
            depo_done    = x.execute("SELECT COUNT(*) FROM depos WHERE status='Completed'").fetchone()[0],
            upcoming     = x.execute(
                "SELECT name,due_date,dtype,status FROM deadlines "
                "WHERE due_date>=? AND status NOT IN('Completed','Vacated') "
                "ORDER BY due_date LIMIT 15",(td,)).fetchall(),
        )
    def close(self): self.cx.close()


# ═════════════════════════════════════════════════════════════════════════════
# PACER PARSER
# ═════════════════════════════════════════════════════════════════════════════
class PACERParser:
    TYPES = [
        ('MEMORANDUM OPINION','Opinion'),('MINUTE ORDER','Minute Order'),
        ('MINUTE ENTRY','Minute Entry'),('NOTICE OF APPEAL','Notice of Appeal'),
        ('NOTICE OF WITHDRAWAL','Notice — Withdrawal'),
        ('NOTICE OF APPEARANCE','Notice — Appearance'),
        ('MOTION TO DISMISS','Mot. to Dismiss'),
        ('MOTION FOR SUMMARY JUDGMENT','MSJ'),
        ('MOTION FOR PRELIMINARY INJUNCTION','Mot. PI'),
        ('MOTION FOR EXTENSION','Mot. Extension'),
        ('MOTION FOR LEAVE','Mot. for Leave'),
        ('EMERGENCY MOTION','Emergency Motion'),
        ('MOTION TO STAY','Mot. to Stay'),
        ('MOTION','Motion'),
        ('MEMORANDUM IN OPPOSITION','Opposition'),
        ('OPPOSITION','Opposition'),
        ('REPLY','Reply'),
        ('SUPPLEMENTAL MEMORANDUM','Supplemental'),
        ('ORDER','Order'),('COMPLAINT','Complaint'),('ANSWER','Answer'),
        ('RETURN OF SERVICE','Service/Summons'),('SUMMONS','Summons'),
        ('TRANSCRIPT','Transcript'),('STIPULATION','Stipulation'),
        ('JOINT STATUS REPORT','Status Report'),('STATUS REPORT','Status Report'),
        ('DECLARATION','Declaration'),('CERTIFICATE','Certificate'),
        ('NOTICE','Notice'),
    ]

    def parse(self, pdf_path):
        doc  = fitz.open(str(pdf_path))
        text = '\n'.join(p.get_text() for p in doc)
        doc.close()
        return dict(
            case_info = self._case_info(text),
            parties   = self._parties(text),
            counsel   = self._counsel(text),
            docket    = self._docket(text),
        )

    def _case_info(self, text):
        d = {}
        for line in text.split('\n')[:25]:
            l = line.strip()
            if ('District Court' in l or 'Court of' in l) and 'court' not in d:
                d['court'] = l
        m = re.search(r'CASE #[:\s]+([^\n]+)', text, re.I)
        if m: d['case_number'] = m.group(1).strip().replace('\u2212','-').replace('\u2013','-')
        m = re.search(r'CASE #[^\n]+\n([^\n]+)', text, re.I)
        if m: d['case_name'] = m.group(1).strip()
        m = re.search(r'Assigned to:\s*([^\n]+)', text)
        if m: d['judge'] = m.group(1).strip()
        m = re.search(r'Date Filed:\s*(\d{2}/\d{2}/\d{4})', text)
        if m:
            try:    d['date_filed'] = datetime.strptime(m.group(1),'%m/%d/%Y').strftime('%Y-%m-%d')
            except: d['date_filed'] = m.group(1)
        m = re.search(r'Nature of Suit:\s*([^\n]+)', text)
        if m: d['nature_of_suit'] = m.group(1).strip()
        return d

    def _parties(self, text):
        out, seen = [], set()
        role_re = re.compile(
            r'^\s*(Plaintiff|Defendant|Third.Party Plaintiff|Petitioner|'
            r'Respondent|Cross.Defendant|Counter.Claimant)\s*$',
            re.MULTILINE|re.I)
        for m in role_re.finditer(text):
            role    = m.group(1).strip()
            snippet = text[m.end(): m.end()+500]
            lines   = [l.strip() for l in snippet.split('\n') if l.strip()]
            for line in lines[:10]:
                if 'represented by' in line.lower():
                    name = line.split('represented by')[0].strip()
                    if name and name not in seen:
                        out.append({'name': name, 'role': role.capitalize()})
                        seen.add(name)
                    break
                if re.match(r'^[A-Z][A-Z\s\.,&]+$', line) and len(line)>3:
                    clean = line.strip()
                    if clean not in seen and clean not in ('V.','AND','THE','ET AL'):
                        out.append({'name': clean, 'role': role.capitalize()})
                        seen.add(clean)
                    break
        return out

    # Address-line patterns to reject as attorney names
    _ADDR_RE = re.compile(
        r'\b(Street|Avenue|Boulevard|Drive|Road|Lane|Court|Place|Way|'
        r'NW|NE|SW|SE|Suite|Floor|Room|Box|'
        r'AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|'
        r'MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|'
        r'RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b.*\d{5}|'
        r'\d{5}|\(\d{3}\)\s*\d{3}',
        re.IGNORECASE)

    def _is_atty_name(self, line):
        """Return True if line looks like an attorney name (Title Case), not
        an address, firm name (ALL CAPS), or label line."""
        # Must start Capital + lowercase (title case), not all-caps
        if not re.match(r'^[A-Z][a-z]', line): return False
        if len(line) > 65 or '@' in line: return False
        if self._ADDR_RE.search(line): return False
        skip = ('STREET','AVENUE','SUITE','FLOOR','EMAIL:','FAX:',
                'ATTORNEY TO BE','PRO HAC','TERMINATED','LEAD ATTORNEY',
                'SEE ABOVE','ONE CITY','CITY CENTER','TENTH STREET',
                'LOUISIANA','PENNSYLVANIA','UNION STREET','L STREET',
                'GENERAL COUNSEL','OFFICE OF','ROOM ')
        u = line.upper()
        if any(s in u for s in skip): return False
        # Reject lines that contain only initials / abbreviations
        if re.match(r'^[A-Z]\.\s', line): return False
        return True

    def _counsel(self, text):
        out, seen = [], set()
        for block in re.split(r'represented by\s*', text, flags=re.I)[1:]:
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            i = 0
            while i < len(lines):
                line = lines[i]
                if re.match(r'^(Plaintiff|Defendant|Date Filed|V\.)$', line, re.I):
                    break
                if self._is_atty_name(line):
                    a = dict(name=line, firm='', phone='', email='', is_lead=0, term_date='')
                    i += 1
                    while i < len(lines):
                        l = lines[i]
                        if self._is_atty_name(l):
                            break
                        if re.search(r'[Ee]mail:', l):
                            a['email'] = re.sub(r'[Ee]mail:\s*','',l).strip()
                        elif re.search(r'[\(\d]\d{2}[\)\-\u2212 ]\s*\d{3}[\-\u2212]\d{4}', l):
                            m = re.search(r'(\d{3})[\)\-\u2212 ]+(\d{3})[\-\u2212](\d{4})', l)
                            if m: a['phone'] = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                        elif 'LEAD ATTORNEY' in l.upper():
                            a['is_lead'] = 1
                        elif re.search(r'TERMINATED:\s*(\d{2}/\d{2}/\d{4})', l):
                            m = re.search(r'TERMINATED:\s*(\d{2}/\d{2}/\d{4})', l)
                            try: a['term_date'] = datetime.strptime(m.group(1),'%m/%d/%Y').strftime('%Y-%m-%d')
                            except: pass
                        elif not a['firm'] and any(x in l for x in [
                                'LLP','LLC','P.C.','PC','PLLC','COMMISSION',
                                'DEPARTMENT','GOVERNMENT','FEDERAL']):
                            a['firm'] = l
                        i += 1
                    if a['name'] not in seen:
                        out.append(a); seen.add(a['name'])
                else:
                    i += 1
        return out

    def _docket(self, text):
        """Parse docket entries. PACER format: Date on own line, then entry# + text."""
        # Find the docket table header — could be inline or each word on its own line
        start = -1
        # Try inline first
        for marker in ['Date Filed # Docket Text', 'Date Filed\t#\tDocket Text']:
            idx = text.find(marker)
            if idx != -1: start = idx + len(marker); break
        # Try multi-line header (each token on its own line)
        if start == -1:
            m = re.search(r'Date Filed\s*\n\s*#\s*\n\s*Docket Text', text)
            if m: start = m.end()
        if start == -1:
            m = re.search(r'Date Filed.*?Docket Text', text, re.DOTALL)
            if m: start = m.end()
        if start == -1: return []

        entries, cur = [], None
        lines = text[start:].split('\n')
        date_re    = re.compile(r'^(\d{2}/\d{2}/\d{4})$')          # date alone on line
        date_txt_re= re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.+)')    # date + text on same line
        entry_re   = re.compile(r'^(\d+)\s+(.*)')                    # entry# + text

        i = 0
        while i < len(lines):
            ls = lines[i].strip()
            i += 1
            if not ls: continue

            # Date alone on its own line (most common PACER format)
            if date_re.match(ls):
                if cur: entries.append(self._finish(cur))
                try:    isodate = datetime.strptime(ls,'%m/%d/%Y').strftime('%Y-%m-%d')
                except: isodate = ls
                # Next non-blank line is either "N text" or just text (minute orders)
                desc, entry_no = '', ''
                while i < len(lines):
                    nxt = lines[i].strip()
                    i += 1
                    if not nxt: continue
                    # Could be another date (back-to-back dates) — push back
                    if date_re.match(nxt) or date_txt_re.match(nxt):
                        i -= 1; break
                    m2 = entry_re.match(nxt)
                    if m2: entry_no = m2.group(1); desc = m2.group(2)
                    else:  desc = nxt
                    break
                cur = dict(date_filed=isodate, entry_no=entry_no,
                           description=desc, filing_party='', file_path='', notes='')

            # Date + text on same line (some PACER variants)
            elif date_txt_re.match(ls):
                if cur: entries.append(self._finish(cur))
                m2 = date_txt_re.match(ls)
                try:    isodate = datetime.strptime(m2.group(1),'%m/%d/%Y').strftime('%Y-%m-%d')
                except: isodate = m2.group(1)
                rest = m2.group(2).strip()
                m3   = entry_re.match(rest)
                if m3: entry_no, desc = m3.group(1), m3.group(2)
                else:  entry_no, desc = '', rest
                cur = dict(date_filed=isodate, entry_no=entry_no,
                           description=desc, filing_party='', file_path='', notes='')

            # Continuation line
            elif cur and ls and not re.match(r'^\d+$', ls):
                cur['description'] += ' ' + ls

        if cur: entries.append(self._finish(cur))
        return entries

    def _finish(self, e):
        d = e['description'].replace('\u2212','-').replace('\u2013','-').strip()
        e['description'] = d
        e['doc_type'] = self._classify(d)
        m = re.search(r'filed by ([^.(]+)', d, re.I)
        if m: e['filing_party'] = m.group(1).strip().rstrip(',')
        return e

    def _classify(self, text):
        """Return doc_type string for any text fragment."""
        u = (text or '').upper()
        return next((dt for kw, dt in self.TYPES if kw in u), 'Other')

    # ── Attachment helpers ────────────────────────────────────────────────────
    # These are used when a filed PDF has an ECF number like "1-3", meaning
    # attachment #3 to docket entry #1.  PACER lists them in the parent
    # entry description as "(Attachments: # 1 Exhibit A, # 2 Exhibit B, ...)"

    ATT_TYPES = [
        # (keyword_regex,              display_type)
        (r'^exhibit',                  'Exhibit'),
        (r'^declaration',              'Declaration'),
        (r'^affidavit',                'Affidavit'),
        (r'^proposed order',           'Proposed Order'),
        (r'^text of proposed order',   'Proposed Order'),
        (r'^memorandum in support',    'Memo in Support'),
        (r'^memorandum',               'Memorandum'),
        (r'^civil cover',              'Cover Sheet'),
        (r'^certificate',              'Certificate'),
        (r'^notice',                   'Notice'),
        (r'^summons',                  'Summons'),
        (r'^proof of (delivery|service)', 'Proof of Service'),
        (r'^transcript',               'Transcript'),
        (r'^stipulation',              'Stipulation'),
        (r'^order',                    'Order'),
        (r'^complaint',                'Complaint'),
    ]

    def parse_attachment_list(self, description: str) -> dict:
        """
        Parse the Attachments clause from a PACER docket description.
        Returns {attachment_number_str: label_str}, e.g. {'1':'Exhibit A', '3':'Civil Cover Sheet'}.
        """
        if not description:          # handles None (SQL NULL) and ''
            return {}
        m = re.search(r'\(Attachments?:(.*?)(?:\(Entered:|$)', description,
                      re.IGNORECASE | re.DOTALL)
        if not m:
            return {}
        labels = {}
        for am in re.finditer(r'#\s*(\d+)\s+([^,#\n\(]+)', m.group(1)):
            num   = am.group(1).strip()
            label = am.group(2).strip().rstrip(',)').strip()
            if label:
                labels[num] = label
        return labels

    def classify_attachment(self, label: str) -> str:
        """Turn a PACER attachment label like 'Exhibit A' into a doc_type."""
        l = label.lower().strip()
        for pat, dtype in self.ATT_TYPES:
            if re.match(pat, l):
                return dtype
        return 'Attachment'

    def _extract_deadlines(self, docket_entries):
        """
        Scan parsed docket entry descriptions for court-imposed deadlines.
        Returns list of dicts ready for the deadlines table.
        """
        MDY_RE = re.compile(r'\b(\d{1,2}/\d{1,2}/\d{4})\b')
        MON_RE = re.compile(
            r'\b(January|February|March|April|May|June|July|August|'
            r'September|October|November|December)'
            r'\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b', re.IGNORECASE)

        def _dates_from(text, pos):
            """Return ISO dates found in text at or after pos, earliest first."""
            hits = []
            for m in MDY_RE.finditer(text, pos):
                try:
                    hits.append((m.start(),
                                 datetime.strptime(m.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')))
                except ValueError:
                    pass
            for m in MON_RE.finditer(text, pos):
                try:
                    hits.append((m.start(),
                                 datetime.strptime(
                                     f"{m.group(1)} {int(m.group(2)):02d} {m.group(3)}",
                                     '%B %d %Y').strftime('%Y-%m-%d')))
                except ValueError:
                    pass
            hits.sort()
            return [d for _, d in hits]

        # (trigger regex, human name, dtype)
        RULES = [
            (r'\banswer due\b',                        'Answer Deadline',             'Answer'),
            (r'\bdeadline to answer\b',                'Answer / Response Deadline',  'Answer'),
            (r'\bmust answer or otherwise respond\b',  'Answer / Response Deadline',  'Answer'),
            (r'\boral arguments?\b',                   'Oral Argument',               'Hearing'),
            (r'\bappear at oral argument\b',           'Oral Argument',               'Hearing'),
            (r'\bappear (?:for|at) (?:a )?hearing\b', 'Hearing',                     'Hearing'),
            (r'\bdirects? the parties to appear\b',    'Court Appearance',            'Hearing'),
            (r'\bplaintiff to file (?:any )?reply\b',  'Plaintiff Reply Deadline',    'Motion'),
            (r'\bfile (?:a )?reply\b',                 'Reply Deadline',              'Motion'),
            (r'\bplaintiff to file\b',                 'Plaintiff Filing Deadline',   'Motion'),
            (r'\bdefendants? to file\b',               'Defense Filing Deadline',     'Motion'),
            (r'\bfile (?:a )?(?:response|opposition)\b','Response Deadline',          'Motion'),
            (r'\bfile (?:a )?joint status report\b',   'Joint Status Report Due',     'Other'),
            (r'\bfile supplemental briefing\b',        'Supplemental Briefing Due',   'Motion'),
            (r'\brespond to (?:that|the)\b',           'Response Deadline',           'Motion'),
            (r'\bfile (?:any )?reply\b',               'Reply Deadline',              'Motion'),
            (r'\bredaction request due\b',             'Redaction Request Deadline',  'Other'),
            (r'\btranscript deadline\b',               'Transcript Deadline',         'Other'),
            (r'\brelease of transcript restriction\b', 'Transcript Public Release',   'Other'),
        ]

        out, seen = [], set()
        for entry in docket_entries:
            raw  = (entry.get('description') or '')
            desc = raw.replace('\u2212', '-').replace('\u2013', '-')
            eref = entry.get('entry_no', '')
            if not desc:
                continue
            for pat, name, dtype in RULES:
                m = re.search(pat, desc, re.IGNORECASE)
                if not m:
                    continue
                dates = _dates_from(desc, m.start())
                if not dates:
                    continue
                due = dates[0]
                key = (due, name)
                if key in seen:
                    continue
                seen.add(key)
                snippet = desc[:110] + ('…' if len(desc) > 110 else '')
                out.append(dict(
                    name       = name,
                    dtype      = dtype,
                    due_date   = due,
                    docket_ref = eref,
                    notes      = snippet,
                    status     = 'Upcoming',
                    _include   = True,
                ))

        out.sort(key=lambda x: x['due_date'])
        return out


# ═════════════════════════════════════════════════════════════════════════════
# DIALOGS
# ═════════════════════════════════════════════════════════════════════════════

def _dlg(parent, title, w=520, h=460):
    top = tk.Toplevel(parent)
    top.title(title); top.configure(bg=C['bg'])
    top.resizable(False, False); top.grab_set()
    top.geometry(f'{w}x{h}+{parent.winfo_rootx()+80}+{parent.winfo_rooty()+60}')
    return top

def _dlg_btns(top, save_fn):
    f = tk.Frame(top, bg=C['bg'], pady=8)
    f.pack(fill='x', padx=16, side='bottom')
    _Btn(f,'Cancel', top.destroy).pack(side='right', padx=(4,0))
    _Btn(f,'Save', save_fn, bg=C['green'], fg=C['bg']).pack(side='right')
    return f

def first_run_dialog(parent, db):
    """
    First-run welcome screen.
    Presents two paths: import PACER PDF (recommended) or enter manually.
    Handles all DB writes. Returns True when setup is complete, False to quit.
    """
    top = tk.Toplevel(parent)
    top.title('Case Command — New Case Setup')
    top.configure(bg=C['bg'])
    top.resizable(False, False)
    top.grab_set()
    top.geometry(f'760x480+{parent.winfo_rootx()+80}+{parent.winfo_rooty()+60}')

    result = {'done': False}

    # ── Header ────────────────────────────────────────────────────────────────
    hdr = tk.Frame(top, bg=C['bg2'], pady=20)
    hdr.pack(fill='x')
    tk.Label(hdr, text='CASE COMMAND', bg=C['bg2'], fg=C['acc'],
             font=('Helvetica', 22, 'bold')).pack()
    tk.Label(hdr, text='Litigation Management  ·  New Case Setup',
             bg=C['bg2'], fg=C['fg3'], font=('Helvetica', 10)).pack()

    # ── Two-panel body ────────────────────────────────────────────────────────
    body = tk.Frame(top, bg=C['bg'], padx=20, pady=20)
    body.pack(fill='both', expand=True)
    body.columnconfigure(0, weight=3)   # PACER panel wider
    body.columnconfigure(1, weight=1)   # divider
    body.columnconfigure(2, weight=2)   # manual panel

    # ── Left: PACER panel (recommended) ──────────────────────────────────────
    lp = tk.Frame(body, bg=C['bg3'], bd=0, relief='flat', padx=20, pady=20)
    lp.grid(row=0, column=0, sticky='nsew', padx=(0, 8))

    # "Recommended" badge
    badge = tk.Frame(lp, bg=C['acc'])
    badge.pack(anchor='ne', fill='x', pady=(0, 12))
    tk.Label(badge, text='  ★  RECOMMENDED  ★  ', bg=C['acc'], fg=C['bg'],
             font=('Helvetica', 8, 'bold')).pack()

    tk.Label(lp, text='Start from PACER Docket PDF',
             bg=C['bg3'], fg=C['fg'], font=('Helvetica', 13, 'bold'),
             wraplength=280, justify='left').pack(anchor='w')

    desc = (
        'Have a docket sheet printed from PACER?  Case Command will '
        'read it automatically and populate:\n\n'
        '  •  Case name, court, number & judge\n'
        '  •  All parties and their counsel\n'
        '  •  Every docket entry\n'
        '  •  Upcoming deadlines visible on the docket\n\n'
        'Takes about 5 seconds — nothing to type.'
    )
    tk.Label(lp, text=desc, bg=C['bg3'], fg=C['fg2'],
             font=('Helvetica', 9), wraplength=280, justify='left').pack(
        anchor='w', pady=(10, 20))

    def on_pacer():
        fp = filedialog.askopenfilename(
            title='Select PACER Docket PDF',
            filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')],
            parent=top)
        if not fp:
            return
        # Parse immediately with progress feedback
        top.configure(cursor='watch'); top.update()
        try:
            parser = PACERParser()
            parsed = parser.parse(fp)
            parsed['deadlines'] = parser._extract_deadlines(parsed['docket'])
        except Exception as e:
            top.configure(cursor='')
            messagebox.showerror('Parse Error',
                f'Could not read the PDF:\n\n{e}', parent=top)
            return
        top.configure(cursor='')
        top.withdraw()
        ok = first_run_pacer_dialog(parent, db, parsed)
        if ok:
            result['done'] = True
            top.destroy()
        else:
            top.deiconify()

    tk.Button(lp, text='  Browse for PACER PDF…  ',
              command=on_pacer,
              bg=C['acc'], fg=C['bg'],
              font=('Helvetica', 11, 'bold'),
              relief='flat', padx=12, pady=8,
              cursor='hand2',
              activebackground=C['mauve'],
              activeforeground=C['bg']).pack(fill='x')

    # ── Divider ───────────────────────────────────────────────────────────────
    div = tk.Frame(body, bg=C['bg4'], width=1)
    div.grid(row=0, column=1, sticky='ns', padx=8)

    # ── Right: Manual panel ───────────────────────────────────────────────────
    rp = tk.Frame(body, bg=C['bg'], padx=8, pady=8)
    rp.grid(row=0, column=2, sticky='nsew')
    rp.columnconfigure(0, weight=1)

    tk.Label(rp, text='Enter Manually',
             bg=C['bg'], fg=C['fg2'], font=('Helvetica', 11, 'bold'),
             anchor='w').grid(row=0, column=0, sticky='w', pady=(0, 8))

    FIELDS = [('Case Name / Caption', 'case_name'),
              ('Court',               'court'),
              ('Case Number',         'case_number'),
              ('Assigned Judge',      'judge')]
    widgets = {}
    for i, (lbl, key) in enumerate(FIELDS, 1):
        tk.Label(rp, text=lbl + ':', bg=C['bg'], fg=C['fg3'],
                 font=('Helvetica', 8), anchor='w').grid(
            row=i*2-1, column=0, sticky='w')
        e = _E(rp)
        e.grid(row=i*2, column=0, sticky='ew', pady=(0, 6))
        widgets[key] = e

    def on_manual():
        vals = {k: w.get().strip() for k, w in widgets.items()}
        if not vals['case_name']:
            messagebox.showwarning('Required', 'Case name is required.', parent=top)
            return
        db.set_meta(vals)
        result['done'] = True
        top.destroy()

    _Btn(rp, 'Start with This Info', on_manual,
         bg=C['bg3']).grid(row=len(FIELDS)*2+1, column=0,
                           sticky='ew', pady=(12, 0))

    # ── Footer ────────────────────────────────────────────────────────────────
    ft = tk.Frame(top, bg=C['bg2'], pady=6)
    ft.pack(fill='x', side='bottom')
    tk.Label(ft, text='You can update all case information at any time from within the app.',
             bg=C['bg2'], fg=C['fg3'], font=('Helvetica', 8)).pack()

    def on_close():
        result['done'] = False
        top.destroy()
    top.protocol('WM_DELETE_WINDOW', on_close)

    top.wait_window()
    return result['done']


def first_run_pacer_dialog(parent, db, parsed):
    """
    4-tab confirmation dialog shown after parsing a PACER PDF on first run.
    Writes everything to db on confirm. Returns True on success, False on cancel.
    """
    ci        = parsed['case_info']
    parties   = parsed['parties']
    counsel   = parsed['counsel']
    docket    = parsed['docket']
    deadlines = parsed.get('deadlines', [])

    top = tk.Toplevel(parent)
    top.title('Case Command — Review & Confirm Setup')
    top.configure(bg=C['bg'])
    top.resizable(True, True)
    top.grab_set()
    top.geometry(f'920x660+{parent.winfo_rootx()+40}+{parent.winfo_rooty()+20}')

    result = {'done': False}

    # ── Instruction bar ───────────────────────────────────────────────────────
    ib = tk.Frame(top, bg=C['bg2'], pady=8)
    ib.pack(fill='x')
    tk.Label(ib,
             text='Review what was found in the docket.  '
                  'Edit anything before importing.  '
                  'Uncheck rows in the Docket and Deadlines tabs to skip individual items.',
             bg=C['bg2'], fg=C['fg2'], font=('Helvetica', 9)).pack(padx=16)

    # ── Notebook ──────────────────────────────────────────────────────────────
    nb = ttk.Notebook(top)
    nb.pack(fill='both', expand=True, padx=8, pady=(6, 0))

    # ═══ TAB 1: Case Info ════════════════════════════════════════════════════
    t1 = tk.Frame(nb, bg=C['bg']); nb.add(t1, text='  Case Info  ')
    t1.columnconfigure(1, weight=1)

    CI_FIELDS = [('Case Name / Caption', 'case_name'),
                 ('Court',               'court'),
                 ('Case Number',         'case_number'),
                 ('Assigned Judge',      'judge'),
                 ('Date Filed',          'date_filed'),
                 ('Nature of Suit',      'nature_of_suit')]
    ci_widgets = {}
    for i, (lbl, key) in enumerate(CI_FIELDS):
        tk.Label(t1, text=lbl + ':', bg=C['bg'], fg=C['fg2'],
                 font=('Helvetica', 9), anchor='e').grid(
            row=i, column=0, sticky='e', padx=(16, 8), pady=6)
        e = _E(t1, ci.get(key, ''))
        e.grid(row=i, column=1, sticky='ew', padx=(0, 16), pady=6)
        ci_widgets[key] = e

    tk.Label(t1, text='All fields are editable — correct anything the parser got wrong.',
             bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8)
             ).grid(row=len(CI_FIELDS), column=0, columnspan=2,
                    sticky='w', padx=16, pady=(4, 0))

    # ═══ TAB 2: Parties & Counsel ════════════════════════════════════════════
    t2  = tk.Frame(nb, bg=C['bg']); nb.add(t2, text=f'  Parties & Counsel ({len(parties)}p · {len(counsel)}c)  ')

    tk.Label(t2, text=f'  Found {len(parties)} parties and {len(counsel)} counsel — all will be imported.',
             bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8), anchor='w'
             ).pack(fill='x', pady=(6, 2))

    pc_cols = ('Type', 'Name', 'Firm / Role', 'Lead', 'Email')
    pc_tv   = ttk.Treeview(t2, columns=pc_cols, show='headings',
                            style=_tv_style('frpc'))
    for col, w in [('Type', 75), ('Name', 190), ('Firm / Role', 270),
                   ('Lead', 40), ('Email', 210)]:
        pc_tv.heading(col, text=col); pc_tv.column(col, width=w, minwidth=30)
    pc_tv.tag_configure('party',   foreground=C['acc'])
    pc_tv.tag_configure('counsel', foreground=C['fg'])
    for p in parties:
        pc_tv.insert('', 'end',
                     values=('Party', p['name'], p['role'], '', ''),
                     tags=('party',))
    for c in counsel:
        pc_tv.insert('', 'end',
                     values=('Counsel', c['name'], c.get('firm', ''),
                              '★' if c.get('is_lead') else '',
                              c.get('email', '')),
                     tags=('counsel',))
    pc_vsb = ttk.Scrollbar(t2, orient='vertical', command=pc_tv.yview)
    pc_tv.configure(yscrollcommand=pc_vsb.set)
    pc_vsb.pack(side='right', fill='y', padx=(0, 4))
    pc_tv.pack(fill='both', expand=True, padx=(8, 0), pady=(0, 8))

    # ═══ TAB 3: Docket ═══════════════════════════════════════════════════════
    t3 = tk.Frame(nb, bg=C['bg'])
    nb.add(t3, text=f'  Docket ({len(docket)} entries)  ')

    tk.Label(t3,
             text='  Click the ✓ column to include/exclude individual entries.',
             bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8), anchor='w'
             ).pack(fill='x', pady=(6, 2))

    dk_cols = ('✓', '#', 'Date', 'Type', 'Description')
    dk_tv   = ttk.Treeview(t3, columns=dk_cols, show='headings',
                            style=_tv_style('frdk'))
    for col, w in [('✓', 28), ('#', 44), ('Date', 88),
                   ('Type', 120), ('Description', 490)]:
        dk_tv.heading(col, text=col); dk_tv.column(col, width=w, minwidth=24)
    dk_tv.tag_configure('on',  foreground=C['fg'])
    dk_tv.tag_configure('off', foreground=C['fg3'])
    dk_vsb = ttk.Scrollbar(t3, orient='vertical', command=dk_tv.yview)
    dk_tv.configure(yscrollcommand=dk_vsb.set)
    dk_vsb.pack(side='right', fill='y', padx=(0, 4))
    dk_tv.pack(fill='both', expand=True, padx=(8, 0), pady=(0, 8))

    dk_include = {}   # iid → BooleanVar
    for e in docket:
        iid = dk_tv.insert('', 'end', values=(
            '✓', e['entry_no'], disp(e['date_filed']),
            e['doc_type'], (e['description'] or '')[:88]),
            tags=('on',))
        dk_include[iid] = tk.BooleanVar(value=True)

    def dk_toggle(event):
        if dk_tv.identify_column(event.x) != '#1': return
        row = dk_tv.identify_row(event.y)
        if row not in dk_include: return
        v = not dk_include[row].get()
        dk_include[row].set(v)
        vals = list(dk_tv.item(row, 'values'))
        vals[0] = '✓' if v else ''
        dk_tv.item(row, values=vals, tags=('on' if v else 'off',))
    dk_tv.bind('<Button-1>', dk_toggle)

    # ═══ TAB 4: Detected Deadlines ═══════════════════════════════════════════
    t4 = tk.Frame(nb, bg=C['bg'])
    n_dl = len(deadlines)
    nb.add(t4, text=f'  Deadlines Detected ({n_dl})  ')

    if n_dl:
        tk.Label(t4,
                 text='  These deadlines were extracted from the docket text.  '
                      'Click ✓ to include/exclude.  You can edit all deadlines after setup.',
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8), anchor='w'
                 ).pack(fill='x', pady=(6, 2))
    else:
        tk.Label(t4, text='  No upcoming deadlines were detected in the docket text.',
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 9), anchor='w'
                 ).pack(fill='x', pady=20)

    dl_cols = ('✓', 'Deadline', 'Type', 'Due Date', 'From Entry', 'Source Text')
    dl_tv   = ttk.Treeview(t4, columns=dl_cols, show='headings',
                            style=_tv_style('frdl'))
    for col, w in [('✓', 28), ('Deadline', 185), ('Type', 85),
                   ('Due Date', 90), ('From Entry', 75), ('Source Text', 280)]:
        dl_tv.heading(col, text=col); dl_tv.column(col, width=w, minwidth=24)
    dl_tv.tag_configure('on',      foreground=C['fg'])
    dl_tv.tag_configure('off',     foreground=C['fg3'])
    dl_tv.tag_configure('overdue', foreground=C['fg3'])
    dl_vsb = ttk.Scrollbar(t4, orient='vertical', command=dl_tv.yview)
    dl_tv.configure(yscrollcommand=dl_vsb.set)
    dl_vsb.pack(side='right', fill='y', padx=(0, 4))
    dl_tv.pack(fill='both', expand=True, padx=(8, 0), pady=(0, 8))

    dl_include = {}   # iid → BooleanVar
    for d in deadlines:
        tag = 'on'
        # Auto-exclude clearly past deadlines (more than 30 days ago)
        try:
            dt = datetime.strptime(d['due_date'], '%Y-%m-%d').date()
            if dt < TODAY() - timedelta(30):
                d['_include'] = False
                tag = 'off'
        except Exception:
            pass
        iid = dl_tv.insert('', 'end', values=(
            '✓' if d['_include'] else '',
            d['name'], d['dtype'],
            disp(d['due_date']),
            d['docket_ref'],
            (d['notes'] or '')[:60]),
            tags=(tag,))
        dl_include[iid] = tk.BooleanVar(value=d['_include'])

    def dl_toggle(event):
        if dl_tv.identify_column(event.x) != '#1': return
        row = dl_tv.identify_row(event.y)
        if row not in dl_include: return
        v = not dl_include[row].get()
        dl_include[row].set(v)
        vals = list(dl_tv.item(row, 'values'))
        vals[0] = '✓' if v else ''
        dl_tv.item(row, values=vals, tags=('on' if v else 'off',))
    dl_tv.bind('<Button-1>', dl_toggle)

    # ── Bottom bar ────────────────────────────────────────────────────────────
    bf = tk.Frame(top, bg=C['bg'], pady=8)
    bf.pack(fill='x', padx=16, side='bottom')

    stats_lbl = tk.Label(bf, text='', bg=C['bg'], fg=C['fg3'],
                         font=('Helvetica', 8))
    stats_lbl.pack(side='left')

    def _update_stats(*_):
        nd = sum(1 for v in dk_include.values() if v.get())
        nl = sum(1 for v in dl_include.values() if v.get())
        stats_lbl.config(
            text=f'{len(parties)} parties  ·  {len(counsel)} counsel  ·  '
                 f'{nd} docket entries  ·  {nl} deadlines')
    _update_stats()

    _Btn(bf, 'Back', top.destroy).pack(side='right', padx=(4, 0))

    def do_import():
        # 1. Case meta
        meta = {k: w.get().strip() for k, w in ci_widgets.items()}
        if not meta.get('case_name'):
            messagebox.showwarning('Required', 'Case name is required.', parent=top)
            nb.select(0); return
        db.set_meta(meta)

        # 2. Parties
        existing_parties = {p['name'] for p in db.party_all()}
        pid_map = {}
        for p in parties:
            if p['name'] not in existing_parties:
                db.party_add(name=p['name'], role=p['role'],
                             affiliation='', contact='', notes='')
        for p in db.party_all():
            pid_map[p['name']] = p['id']

        # 3. Counsel
        existing_counsel = {c['name'] for c in db.counsel_all()}
        for c in counsel:
            if c['name'] not in existing_counsel:
                # Best-effort party match via firm name
                pid = None
                for pname, pid_v in pid_map.items():
                    if pname.lower() in c.get('firm', '').lower():
                        pid = pid_v; break
                db.counsel_add(
                    name=c['name'], firm=c.get('firm', ''),
                    phone=c.get('phone', ''), email=c.get('email', ''),
                    party_id=pid, is_lead=c.get('is_lead', 0),
                    term_date=c.get('term_date', ''), notes='')

        # 4. Docket entries (checked only)
        for iid, var in dk_include.items():
            if not var.get(): continue
            vals   = dk_tv.item(iid, 'values')
            # vals = (✓, entry_no, date_disp, doc_type, desc_truncated)
            # We need the full entry — match by entry_no + date
            entry_no = vals[1]
            match    = next((e for e in docket
                             if e['entry_no'] == entry_no
                             and disp(e['date_filed']) == vals[2]), None)
            if match:
                db.docket_add(**match)

        # 5. Deadlines (checked only)
        for iid, var in dl_include.items():
            if not var.get(): continue
            vals = dl_tv.item(iid, 'values')
            # vals = (✓, name, dtype, due_date_disp, docket_ref, notes)
            dl   = next((d for d in deadlines
                         if d['name'] == vals[1]
                         and disp(d['due_date']) == vals[3]), None)
            if dl:
                db.dl_add(name=dl['name'], dtype=dl['dtype'],
                          due_date=dl['due_date'],
                          status='Upcoming',
                          docket_ref=dl['docket_ref'],
                          notes=dl['notes'])

        result['done'] = True
        top.destroy()

    _Btn(bf, 'Set Up Case  →', do_import,
         bg=C['green'], fg=C['bg']).pack(side='right')

    top.wait_window()
    return result['done']

def docket_dialog(parent, row=None, db=None):
    """
    Add or edit a docket entry.  If db is provided the dialog offers a
    Parent Entry picker so the user can create attachment sub-entries manually.
    """
    # Convert sqlite3.Row → dict first; Row has no .get() method
    r       = dict(row) if row else {}
    is_att  = bool(r.get('parent_id'))
    h       = 520 if db else 460
    top     = _dlg(parent, 'Edit Docket Entry' if row else 'Add Docket Entry', 580, h)
    f       = tk.Frame(top, bg=C['bg'], padx=16, pady=12)
    f.pack(fill='both', expand=True); f.columnconfigure(1, weight=1)

    DOC_TYPES = [
        'Complaint','Answer','Motion','Mot. to Dismiss','MSJ','Mot. PI',
        'Opposition','Reply','Order','Minute Order','Opinion',
        'Notice of Appeal','Notice — Appearance','Notice — Withdrawal','Notice',
        'Service/Summons','Transcript','Stipulation','Status Report','Supplemental',
        # Attachment-specific types
        'Exhibit','Declaration','Affidavit','Proposed Order','Memo in Support',
        'Cover Sheet','Certificate','Proof of Service','Attachment','Other',
    ]

    rows_data = [
        ('Entry # (ECF)',          _E(f, r.get('entry_no', ''),    w=12)),
        ('Date Filed (YYYY-MM-DD)',_E(f, r.get('date_filed', ''))),
        ('Doc Type',               _CB(f, DOC_TYPES, r.get('doc_type', 'Other'))),
        ('Filing Party',           _E(f, r.get('filing_party', ''))),
    ]
    for i, (lbl, w) in enumerate(rows_data):
        _L(f, lbl + ':').grid(row=i, column=0, sticky='w', pady=3, padx=(0, 10))
        w.grid(row=i, column=1, sticky='ew', pady=3)

    base_row = len(rows_data)

    _L(f, 'Description:').grid(row=base_row, column=0, sticky='nw',
                                pady=(6, 3), padx=(0, 10))
    desc_w = _T(f, r.get('description', ''), h=4)
    desc_w.grid(row=base_row, column=1, sticky='ew', pady=3)

    # ── Attachment section (only shown when db supplied) ──────────────────────
    parent_id_var   = tk.IntVar(value=r.get('parent_id') or 0)
    att_no_var      = tk.StringVar(value=r.get('attachment_no') or '')
    parent_entries  = []   # list of docket rows for combobox
    parent_labels   = []

    if db:
        sep = tk.Frame(f, bg=C['bg4'], height=1)
        sep.grid(row=base_row+1, column=0, columnspan=2, sticky='ew',
                 pady=(10, 4))
        tk.Label(f, text='ATTACHMENT  (leave blank if this is a top-level filing)',
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8, 'italic'),
                 anchor='w').grid(row=base_row+2, column=0, columnspan=2, sticky='w')

        # Parent entry picker
        all_top = [e for e in db.docket_all() if not e['parent_id']]
        parent_labels = ['(none — top-level entry)'] + [
            f"#{e['entry_no']}  {disp(e['date_filed'])}  —  {(e['description'] or '')[:45]}"
            for e in all_top]
        parent_entries = [None] + list(all_top)

        cur_parent_lbl = '(none — top-level entry)'
        if r.get('parent_id'):
            match = next((e for e in all_top if e['id'] == r['parent_id']), None)
            if match:
                cur_parent_lbl = (
                    f"#{match['entry_no']}  {disp(match['date_filed'])}  —  "
                    f"{(match['description'] or '')[:45]}")

        _L(f, 'Parent Entry:').grid(row=base_row+3, column=0, sticky='w',
                                     pady=3, padx=(0, 10))
        parent_cb = _CB(f, parent_labels, cur_parent_lbl)
        parent_cb.grid(row=base_row+3, column=1, sticky='ew', pady=3)

        _L(f, 'Attachment # (e.g. 3\nfor ECF 1-3):').grid(
            row=base_row+4, column=0, sticky='w', pady=3, padx=(0, 10))
        att_e = _E(f, att_no_var.get(), w=8)
        att_e.grid(row=base_row+4, column=1, sticky='w', pady=3)

        def _on_parent_select(event=None):
            """Auto-fill Entry # when a parent is chosen."""
            idx = parent_labels.index(parent_cb.get()) if parent_cb.get() in parent_labels else 0
            pe  = parent_entries[idx]
            if pe:
                parent_id_var.set(pe['id'])
                # Suggest next attachment number
                existing_atts = [e for e in db.docket_all()
                                 if e['parent_id'] == pe['id']]
                next_no = str(len(existing_atts) + 1)
                if not att_e.get().strip():
                    att_e.delete(0, 'end'); att_e.insert(0, next_no)
                # Auto-fill entry_no field as "parent-att"
                en_w = rows_data[0][1]
                if not en_w.get().strip():
                    an = att_e.get().strip() or next_no
                    en_w.delete(0, 'end')
                    en_w.insert(0, f"{pe['entry_no']}-{an}")
            else:
                parent_id_var.set(0)
        parent_cb.bind('<<ComboboxSelected>>', _on_parent_select)

        # Attachment # change auto-updates Entry #
        def _on_att_change(event=None):
            idx = parent_labels.index(parent_cb.get()) if parent_cb.get() in parent_labels else 0
            pe  = parent_entries[idx]
            an  = att_e.get().strip()
            if pe and an:
                en_w = rows_data[0][1]
                en_w.delete(0, 'end')
                en_w.insert(0, f"{pe['entry_no']}-{an}")
        att_e.bind('<KeyRelease>', _on_att_change)

    # ── File attachment ───────────────────────────────────────────────────────
    file_row = base_row + (5 if db else 1)
    fp_var   = tk.StringVar(value=r.get('file_path', ''))

    def browse():
        p = filedialog.askopenfilename(
            title='Attach PDF', filetypes=[('PDF', '*.pdf'), ('All', '*.*')])
        if p: fp_var.set(p)

    _L(f, 'Attached File:').grid(row=file_row, column=0, sticky='w',
                                   padx=(0, 10), pady=3)
    att_frame = tk.Frame(f, bg=C['bg'])
    att_frame.grid(row=file_row, column=1, sticky='ew', pady=3)
    tk.Label(att_frame, textvariable=fp_var, bg=C['bg3'], fg=C['fg'],
             font=('Helvetica', 8), anchor='w', relief='flat'
             ).pack(side='left', fill='x', expand=True)
    _Btn(att_frame, 'Browse', browse).pack(side='right', padx=(4, 0))

    notes_row = file_row + 1
    _L(f, 'Notes:').grid(row=notes_row, column=0, sticky='nw',
                          pady=(6, 3), padx=(0, 10))
    notes_w = _T(f, r.get('notes', ''), h=2)
    notes_w.grid(row=notes_row, column=1, sticky='ew', pady=3)

    result = {}
    def save():
        pid = None
        an  = ''
        if db:
            idx = (parent_labels.index(parent_cb.get())
                   if parent_cb.get() in parent_labels else 0)
            pe  = parent_entries[idx]
            if pe:
                pid = pe['id']
                an  = att_e.get().strip()
        result.update(dict(
            entry_no      = rows_data[0][1].get().strip(),
            date_filed    = iso(rows_data[1][1].get().strip()),
            doc_type      = rows_data[2][1].get(),
            filing_party  = rows_data[3][1].get().strip(),
            description   = desc_w.get('1.0', 'end').strip(),
            file_path     = fp_var.get().strip(),
            notes         = notes_w.get('1.0', 'end').strip(),
            parent_id     = pid,
            attachment_no = an,
        ))
        top.destroy()

    _dlg_btns(top, save)
    top.wait_window()
    return result or None

def deadline_dialog(parent, row=None):
    top = _dlg(parent,'Edit Deadline' if row else 'Add Deadline',500,360)
    f   = tk.Frame(top,bg=C['bg'],padx=16,pady=12)
    f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1)
    r   = dict(row) if row else {}

    DTYPES  = ['Hearing','Motion','Trial','Pretrial','Expert','Scheduling','Discovery','Other']
    STATUSES= ['Upcoming','In Progress','Filed','Completed','Continued','Vacated']

    fields = [
        ('Deadline Name',     _E(f, r.get('name',''))),
        ('Type',              _CB(f, DTYPES,  r.get('dtype','Hearing'))),
        ('Due Date (YYYY-MM-DD)', _E(f, r.get('due_date',''))),
        ('Status',            _CB(f, STATUSES, r.get('status','Upcoming'))),
        ('Related Docket #',  _E(f, r.get('docket_ref',''), w=10)),
    ]
    for i,(lbl,w) in enumerate(fields):
        _L(f,lbl+':').grid(row=i,column=0,sticky='w',pady=4,padx=(0,10))
        w.grid(row=i,column=1,sticky='ew',pady=4)

    _L(f,'Notes:').grid(row=len(fields),column=0,sticky='nw',pady=(6,4),padx=(0,10))
    notes_w = _T(f, r.get('notes',''), h=3)
    notes_w.grid(row=len(fields),column=1,sticky='ew',pady=4)

    result = {}
    def save():
        if not fields[0][1].get().strip():
            messagebox.showwarning('Required','Deadline name required.',parent=top); return
        result.update(dict(
            name       = fields[0][1].get().strip(),
            dtype      = fields[1][1].get(),
            due_date   = iso(fields[2][1].get().strip()),
            status     = fields[3][1].get(),
            docket_ref = fields[4][1].get().strip(),
            notes      = notes_w.get('1.0','end').strip(),
        ))
        top.destroy()
    _dlg_btns(top,save); top.wait_window()
    return result or None

def disc_set_dialog(parent, row=None, default_days=30):
    top = _dlg(parent,'Edit Discovery Set' if row else 'Add Discovery Set',560,500)
    f   = tk.Frame(top,bg=C['bg'],padx=16,pady=12)
    f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1)
    r   = dict(row) if row else {}

    RTYPES   = ['RFP','ROG','RFA']
    DIRS     = ['Received','Propounded']
    STATUSES = ['Pending','Objections Only','Substantive','Supplemented','Complete']
    DEFS     = ['None','Letter Sent','Meet & Confer','Motion Filed','Resolved']

    served_var = tk.StringVar(value=r.get('date_served',''))
    due_var    = tk.StringVar(value=r.get('due_date',''))

    def auto_due(*_):
        s = served_var.get().strip()
        if s and not due_var.get().strip():
            try:
                dt = datetime.strptime(iso(s),'%Y-%m-%d') + timedelta(default_days)
                due_var.set(dt.strftime('%Y-%m-%d'))
            except: pass

    fields = [
        ('Request Type',    _CB(f, RTYPES,  r.get('rtype','RFP'))),
        ('Set Number',      _E(f, r.get('set_no','1'), w=6)),
        ('Direction',       _CB(f, DIRS,    r.get('direction','Received'))),
        ('Propounding Party', _E(f, r.get('prop_party',''))),
        ('Responding Party',  _E(f, r.get('resp_party',''))),
        ('# of Requests',   _E(f, r.get('num_reqs',''), w=6)),
        ('Status',          _CB(f, STATUSES, r.get('status','Pending'))),
        ('Deficiency',      _CB(f, DEFS,    r.get('deficiency','None'))),
    ]
    for i,(lbl,w) in enumerate(fields):
        _L(f,lbl+':').grid(row=i,column=0,sticky='w',pady=3,padx=(0,10))
        w.grid(row=i,column=1,sticky='ew',pady=3)

    # Date served + due (linked)
    _L(f,'Date Served (YYYY-MM-DD):').grid(row=len(fields),column=0,sticky='w',pady=3,padx=(0,10))
    served_e = _E(f, served_var.get())
    served_e.grid(row=len(fields),column=1,sticky='ew',pady=3)
    served_e.bind('<FocusOut>', auto_due)

    _L(f,f'Due Date (YYYY-MM-DD)\n(default +{default_days}d):').grid(
        row=len(fields)+1,column=0,sticky='w',pady=3,padx=(0,10))
    due_e = _E(f, due_var.get())
    due_e.grid(row=len(fields)+1,column=1,sticky='ew',pady=3)

    _L(f,'Response Date:').grid(row=len(fields)+2,column=0,sticky='w',pady=3,padx=(0,10))
    resp_e = _E(f, r.get('resp_date',''))
    resp_e.grid(row=len(fields)+2,column=1,sticky='ew',pady=3)

    _L(f,'Notes:').grid(row=len(fields)+3,column=0,sticky='nw',pady=(6,3),padx=(0,10))
    notes_w = _T(f, r.get('notes',''), h=2)
    notes_w.grid(row=len(fields)+3,column=1,sticky='ew',pady=3)

    result = {}
    def save():
        try: nr = int(fields[1][1].get().strip() or 0)
        except: nr = 0
        result.update(dict(
            rtype      = fields[0][1].get(),
            set_no     = nr,
            direction  = fields[2][1].get(),
            prop_party = fields[3][1].get().strip(),
            resp_party = fields[4][1].get().strip(),
            num_reqs   = int(fields[5][1].get().strip() or 0),
            status     = fields[6][1].get(),
            deficiency = fields[7][1].get(),
            date_served= iso(served_e.get().strip()),
            due_date   = iso(due_e.get().strip()),
            resp_date  = iso(resp_e.get().strip()),
            notes      = notes_w.get('1.0','end').strip(),
        ))
        top.destroy()
    _dlg_btns(top,save); top.wait_window()
    return result or None

def depo_dialog(parent, row=None):
    top = _dlg(parent,'Edit Deposition' if row else 'Add Deposition',560,560)
    f   = tk.Frame(top,bg=C['bg'],padx=16,pady=10)
    f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1)
    r   = dict(row) if row else {}

    ROLES    = ['Party','Fact Witness','Expert','30(b)(6)','Corporate Rep']
    FORMATS  = ['In-Person','Zoom','Phone','Other']
    STATUSES = ['Not Noticed','Noticed','Completed','Transcript Received','Errata Deadline Passed']

    fields = [
        ('Witness Name',      _E(f, r.get('witness',''))),
        ('Role',              _CB(f, ROLES,    r.get('role','Fact Witness'))),
        ('Affiliation',       _E(f, r.get('affiliation',''))),
        ('Notice Date (YYYY-MM-DD)', _E(f, r.get('notice_date',''))),
        ('Depo Date (YYYY-MM-DD)',   _E(f, r.get('depo_date',''))),
        ('Location',          _E(f, r.get('location',''))),
        ('Format',            _CB(f, FORMATS,  r.get('format','In-Person'))),
        ('Status',            _CB(f, STATUSES, r.get('status','Not Noticed'))),
        ('Errata Deadline',   _E(f, r.get('errata_deadline',''))),
        ('Transcript Bates',  _E(f, r.get('transcript_bates',''))),
    ]
    for i,(lbl,w) in enumerate(fields):
        _L(f,lbl+':').grid(row=i,column=0,sticky='w',pady=2,padx=(0,10))
        w.grid(row=i,column=1,sticky='ew',pady=2)

    _L(f,'Key Topics / Notes:').grid(row=len(fields),column=0,sticky='nw',pady=(4,2),padx=(0,10))
    topics_w = _T(f, r.get('topics',''), h=2)
    topics_w.grid(row=len(fields),column=1,sticky='ew',pady=2)

    _L(f,'Outstanding Issues:').grid(row=len(fields)+1,column=0,sticky='nw',pady=(4,2),padx=(0,10))
    issues_w = _T(f, r.get('issues',''), h=2)
    issues_w.grid(row=len(fields)+1,column=1,sticky='ew',pady=2)

    result = {}
    def save():
        if not fields[0][1].get().strip():
            messagebox.showwarning('Required','Witness name required.',parent=top); return
        result.update(dict(
            witness          = fields[0][1].get().strip(),
            role             = fields[1][1].get(),
            affiliation      = fields[2][1].get().strip(),
            notice_date      = iso(fields[3][1].get().strip()),
            depo_date        = iso(fields[4][1].get().strip()),
            location         = fields[5][1].get().strip(),
            format           = fields[6][1].get(),
            status           = fields[7][1].get(),
            errata_deadline  = iso(fields[8][1].get().strip()),
            transcript_bates = fields[9][1].get().strip(),
            topics           = topics_w.get('1.0','end').strip(),
            issues           = issues_w.get('1.0','end').strip(),
            notes            = '',
        ))
        top.destroy()
    _dlg_btns(top,save); top.wait_window()
    return result or None

def party_dialog(parent, row=None):
    top = _dlg(parent,'Edit Party' if row else 'Add Party',480,360)
    f   = tk.Frame(top,bg=C['bg'],padx=16,pady=12)
    f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1)
    r   = dict(row) if row else {}

    ROLES = ['Plaintiff','Defendant','Third Party','Expert','Fact Witness','Other']
    fields = [
        ('Name',        _E(f, r.get('name',''))),
        ('Role',        _CB(f, ROLES, r.get('role','Plaintiff'))),
        ('Affiliation', _E(f, r.get('affiliation',''))),
        ('Contact',     _E(f, r.get('contact',''))),
    ]
    for i,(lbl,w) in enumerate(fields):
        _L(f,lbl+':').grid(row=i,column=0,sticky='w',pady=4,padx=(0,10))
        w.grid(row=i,column=1,sticky='ew',pady=4)
    _L(f,'Notes:').grid(row=len(fields),column=0,sticky='nw',pady=(6,4),padx=(0,10))
    notes_w = _T(f, r.get('notes',''), h=3)
    notes_w.grid(row=len(fields),column=1,sticky='ew',pady=4)

    result = {}
    def save():
        if not fields[0][1].get().strip():
            messagebox.showwarning('Required','Name required.',parent=top); return
        result.update(name=fields[0][1].get().strip(),role=fields[1][1].get(),
                      affiliation=fields[2][1].get().strip(),
                      contact=fields[3][1].get().strip(),
                      notes=notes_w.get('1.0','end').strip())
        top.destroy()
    _dlg_btns(top,save); top.wait_window()
    return result or None

def counsel_dialog(parent, db, row=None):
    top = _dlg(parent,'Edit Counsel' if row else 'Add Counsel',520,400)
    f   = tk.Frame(top,bg=C['bg'],padx=16,pady=12)
    f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1)
    r   = dict(row) if row else {}

    parties      = db.party_all()
    party_names  = [p['name'] for p in parties]
    party_ids    = {p['name']: p['id'] for p in parties}
    cur_party    = next((p['name'] for p in parties if p['id']==r.get('party_id')), '')

    fields = [
        ('Name',         _E(f, r.get('name',''))),
        ('Firm',         _E(f, r.get('firm',''))),
        ('Phone',        _E(f, r.get('phone',''))),
        ('Email',        _E(f, r.get('email',''))),
        ('Party',        _CB(f, party_names, cur_party)),
        ('Terminated (YYYY-MM-DD)', _E(f, r.get('term_date',''))),
    ]
    for i,(lbl,w) in enumerate(fields):
        _L(f,lbl+':').grid(row=i,column=0,sticky='w',pady=4,padx=(0,10))
        w.grid(row=i,column=1,sticky='ew',pady=4)

    lead_var = tk.IntVar(value=r.get('is_lead',0))
    tk.Checkbutton(f,text='Lead Attorney',variable=lead_var,
                   bg=C['bg'],fg=C['fg'],selectcolor=C['bg3'],
                   activebackground=C['bg'],font=('Helvetica',10)).grid(
        row=len(fields),column=1,sticky='w',pady=4)

    result = {}
    def save():
        if not fields[0][1].get().strip():
            messagebox.showwarning('Required','Name required.',parent=top); return
        pname = fields[4][1].get()
        result.update(
            name     = fields[0][1].get().strip(),
            firm     = fields[1][1].get().strip(),
            phone    = fields[2][1].get().strip(),
            email    = fields[3][1].get().strip(),
            party_id = party_ids.get(pname),
            is_lead  = lead_var.get(),
            term_date= iso(fields[5][1].get().strip()),
            notes    = '',
        )
        top.destroy()
    _dlg_btns(top,save); top.wait_window()
    return result or None

def pacer_import_dialog(parent, db, pdf_path):
    """Parse PACER PDF, show preview, let user confirm import."""
    parser = PACERParser()
    try:
        parsed = parser.parse(pdf_path)
    except Exception as e:
        messagebox.showerror('Parse Error', f'Could not parse PDF:\n{e}', parent=parent)
        return

    top = tk.Toplevel(parent)
    top.title('PACER Import — Review & Confirm')
    top.configure(bg=C['bg'])
    top.geometry(f'860x660+{parent.winfo_rootx()+40}+{parent.winfo_rooty()+30}')
    top.grab_set()

    nb = ttk.Notebook(top)
    nb.pack(fill='both', expand=True, padx=8, pady=8)

    # ── Tab 1: Case Info ──────────────────────────────────────────────────────
    t1 = tk.Frame(nb, bg=C['bg']); nb.add(t1, text='  Case Info  ')
    t1.columnconfigure(1, weight=1)
    ci    = parsed['case_info']
    ci_e  = {}
    flds  = [('Case Name','case_name'),('Court','court'),
             ('Case Number','case_number'),('Judge','judge'),
             ('Date Filed','date_filed'),('Nature of Suit','nature_of_suit')]
    for i,(lbl,key) in enumerate(flds):
        _L(t1,lbl+':').grid(row=i,column=0,sticky='w',padx=(16,8),pady=5)
        e = _E(t1, ci.get(key,''))
        e.grid(row=i,column=1,sticky='ew',padx=(0,16),pady=5)
        ci_e[key] = e

    # ── Tab 2: Parties & Counsel ──────────────────────────────────────────────
    t2  = tk.Frame(nb,bg=C['bg']); nb.add(t2,text='  Parties & Counsel  ')
    tv2 = ttk.Treeview(t2,columns=('Type','Name','Firm/Role'),show='headings',
                       style=_tv_style('pi'))
    for col,w in [('Type',90),('Name',200),('Firm/Role',340)]:
        tv2.heading(col,text=col); tv2.column(col,width=w)
    for p in parsed['parties']:
        tv2.insert('','end',values=('Party',p['name'],p['role']))
    for c in parsed['counsel']:
        tv2.insert('','end',values=('Counsel',c['name'],c['firm']))
    tv2.pack(fill='both',expand=True,padx=8,pady=8)
    tk.Label(t2,text='All parties and counsel above will be imported.',
             bg=C['bg'],fg=C['fg3'],font=('Helvetica',8)).pack(pady=(0,6))

    # ── Tab 3: Docket Entries ─────────────────────────────────────────────────
    t3   = tk.Frame(nb,bg=C['bg']); nb.add(t3,text=f"  Docket ({len(parsed['docket'])} entries)  ")
    existing = db.docket_existing_keys()

    tk.Label(t3,text='Uncheck entries you do not want to import. '
             'Greyed entries already exist in the database.',
             bg=C['bg'],fg=C['fg3'],font=('Helvetica',8)).pack(anchor='w',padx=8,pady=(8,2))

    cols3 = ('Import','#','Date','Type','Description')
    tv3   = ttk.Treeview(t3,columns=cols3,show='headings',style=_tv_style('pi2'))
    for col,w in [('Import',50),('#',40),('Date',90),('Type',110),('Description',420)]:
        tv3.heading(col,text=col); tv3.column(col,width=w,minwidth=30)

    check_vars = {}
    for e in parsed['docket']:
        key    = (e['date_filed'], e['entry_no'])
        exists = key in existing
        var    = tk.BooleanVar(value=not exists)
        iid    = tv3.insert('','end',values=(
            '' if exists else 'Yes',
            e['entry_no'],
            disp(e['date_filed']),
            e['doc_type'],
            e['description'][:80]
        ))
        check_vars[iid] = (var, e)
        if exists:
            tv3.item(iid, tags=('exists',))
    tv3.tag_configure('exists', foreground=C['fg3'])

    def toggle(event):
        iid = tv3.identify_row(event.y)
        if iid and iid in check_vars:
            var, e = check_vars[iid]
            if (e['date_filed'], e['entry_no']) not in existing:
                var.set(not var.get())
                tv3.item(iid, values=('Yes' if var.get() else '', *tv3.item(iid,'values')[1:]))
    tv3.bind('<Button-1>', toggle)

    vsb3 = ttk.Scrollbar(t3,orient='vertical',command=tv3.yview)
    tv3.configure(yscrollcommand=vsb3.set)
    vsb3.pack(side='right',fill='y',padx=(0,4))
    tv3.pack(fill='both',expand=True,padx=(8,0),pady=(0,8))

    # ── Bottom buttons ────────────────────────────────────────────────────────
    bf = tk.Frame(top,bg=C['bg'],pady=8)
    bf.pack(fill='x',padx=16,side='bottom')
    status_lbl = tk.Label(bf,text='',bg=C['bg'],fg=C['acc'],font=('Helvetica',9))
    status_lbl.pack(side='left')
    _Btn(bf,'Cancel',top.destroy).pack(side='right',padx=(4,0))

    def do_import():
        # 1. Update case meta
        for k,e in ci_e.items():
            v = e.get().strip()
            if v: db.set_meta({k: v})
        # 2. Import parties
        existing_party_names = {p['name'] for p in db.party_all()}
        party_id_map = {}
        for p in parsed['parties']:
            if p['name'] not in existing_party_names:
                db.party_add(name=p['name'],role=p['role'],
                             affiliation='',contact='',notes='')
        for p in db.party_all():
            party_id_map[p['name']] = p['id']
        # 3. Import counsel
        existing_counsel = {c['name'] for c in db.counsel_all()}
        for c in parsed['counsel']:
            if c['name'] not in existing_counsel:
                # Try to match party
                pid = None
                for pname, pid_v in party_id_map.items():
                    if pname.lower() in c.get('firm','').lower():
                        pid = pid_v; break
                db.counsel_add(name=c['name'],firm=c.get('firm',''),
                               phone=c.get('phone',''),email=c.get('email',''),
                               party_id=pid,is_lead=c.get('is_lead',0),
                               term_date=c.get('term_date',''),notes='')
        # 4. Import selected docket entries
        imported = 0
        for iid,(var,e) in check_vars.items():
            if var.get():
                db.docket_add(**e); imported += 1
        status_lbl.config(text=f'Imported: {imported} docket entries, '
                               f'{len(parsed["parties"])} parties, '
                               f'{len(parsed["counsel"])} counsel.')
        messagebox.showinfo('Import Complete',
            f'Successfully imported:\n'
            f'  {imported} docket entries\n'
            f'  {len(parsed["parties"])} parties\n'
            f'  {len(parsed["counsel"])} counsel\n\n'
            f'Case info updated. Click OK to refresh.',
            parent=top)
        top.destroy()

    _Btn(bf,'Import Selected',do_import,bg=C['green'],fg=C['bg']).pack(side='right')
    top.wait_window()


# ═════════════════════════════════════════════════════════════════════════════
# PDF COURT-HEADER SCANNER
# ═════════════════════════════════════════════════════════════════════════════
def scan_pdf_header(pdf_path: Path) -> dict:
    """
    Read the first page of a court-filed PDF and extract the ECF header stamp.

    Standard entry:       Document 22     Filed 07/14/25
    Attachment:           Document 22-3   Filed 07/14/25   (attachment #3 to entry 22)

    Returns dict with keys:
        entry_no       – parent entry number string (e.g. '22')
        attachment_no  – attachment number string (e.g. '3'), empty if not an attachment
        is_attachment  – True if this is an attachment to another entry
        date_filed     – ISO YYYY-MM-DD
        description    – first substantive heading found in the document body
        confidence     – 'high' | 'low' | 'none'
        raw_header     – the matched header text for display
    """
    out = dict(entry_no='', attachment_no='', is_attachment=False,
               date_filed='', description='', confidence='none', raw_header='')
    try:
        doc  = fitz.open(str(pdf_path))
        text = doc[0].get_text()
        if len(text.strip()) < 60 and len(doc) > 1:
            text += '\n' + doc[1].get_text()
        doc.close()
    except Exception:
        return out

    def _parse_date(s):
        for fmt in ('%m/%d/%Y', '%m/%d/%y'):
            try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
            except ValueError: pass
        return ''

    # ── Pattern A: "Document N-M  Filed ..." (attachment) ────────────────────
    m = re.search(
        r'Case\s+[\w:\.\-]+\s+Document\s+(\d+)-(\d+)\s+Filed\s+(\d{1,2}/\d{1,2}/\d{2,4})',
        text, re.IGNORECASE)
    if m:
        out['raw_header']    = m.group(0).replace('\n', ' ').strip()
        out['entry_no']      = m.group(1)
        out['attachment_no'] = m.group(2)
        out['is_attachment'] = True
        out['confidence']    = 'high'
        out['date_filed']    = _parse_date(m.group(3))

    # ── Pattern A2: "Document N  Filed ..." (main entry) ─────────────────────
    if not out['entry_no']:
        m = re.search(
            r'Case\s+[\w:\.\-]+\s+Document\s+(\d+)\s+Filed\s+(\d{1,2}/\d{1,2}/\d{2,4})',
            text, re.IGNORECASE)
        if m:
            out['raw_header'] = m.group(0).replace('\n', ' ').strip()
            out['entry_no']   = m.group(1)
            out['confidence'] = 'high'
            out['date_filed'] = _parse_date(m.group(2))

    # ── Pattern B: "Doc #: N-M  Filed: ..." variant ──────────────────────────
    if not out['entry_no']:
        m = re.search(
            r'Doc(?:ument)?\s*#?:?\s*(\d+)(?:-(\d+))?\s+Filed:?\s+(\d{1,2}/\d{1,2}/\d{2,4})',
            text, re.IGNORECASE)
        if m:
            out['raw_header']    = m.group(0).strip()
            out['entry_no']      = m.group(1)
            out['attachment_no'] = m.group(2) or ''
            out['is_attachment'] = bool(m.group(2))
            out['confidence']    = 'high'
            out['date_filed']    = _parse_date(m.group(3))

    # ── Pattern C: "Document N-M" or "Document N" near top (no date) ─────────
    if not out['entry_no']:
        m = re.search(r'\bDocument\s+(\d+)(?:-(\d+))?\b', text[:600], re.IGNORECASE)
        if m:
            out['entry_no']      = m.group(1)
            out['attachment_no'] = m.group(2) or ''
            out['is_attachment'] = bool(m.group(2))
            out['confidence']    = 'low'
            out['raw_header']    = m.group(0)

    # ── Pattern D: Filename fallback "N-M_..." or "NNN_..." ──────────────────
    if not out['entry_no']:
        fname = pdf_path.stem
        m = re.match(r'^0*(\d{1,5})-(\d{1,3})[_\-\s]', fname)
        if m and int(m.group(1)) > 0:
            out['entry_no']      = str(int(m.group(1)))
            out['attachment_no'] = str(int(m.group(2)))
            out['is_attachment'] = True
            out['confidence']    = 'low'
            out['raw_header']    = f'Filename: {pdf_path.name}'
        else:
            m2 = re.match(r'^0*(\d{1,5})[_\-\s]', fname)
            if m2 and int(m2.group(1)) > 0:
                out['entry_no']   = str(int(m2.group(1)))
                out['confidence'] = 'low'
                out['raw_header'] = f'Filename: {pdf_path.name}'

    # ── Extract document title from page body ─────────────────────────────────
    CAPTION_SKIP = {
        'UNITED STATES', 'DISTRICT COURT', 'DISTRICT OF', 'COURT FOR',
        'IN THE', 'FOR THE', 'CIVIL ACTION', 'CIVIL CASE', 'CASE NO',
        'PLAINTIFF', 'DEFENDANT', 'RESPONDENT', 'PETITIONER',
        'ET AL', 'V.', 'VS.', 'NO.', 'CIV.', 'MDL',
    }
    lines       = [l.strip() for l in text.split('\n') if l.strip()]
    past_header = bool(out['entry_no'])

    for line in lines:
        if re.search(r'Page\s+\d+\s+of\s+\d+', line, re.IGNORECASE):
            past_header = True
            continue
        if not past_header: continue
        if len(line) < 8 or re.match(r'^[\d\s\W]+$', line): continue
        u = line.upper()
        if any(s in u for s in CAPTION_SKIP): continue
        if re.match(r'^Case\s+[\d:\-]+', line, re.IGNORECASE): continue
        out['description'] = line[:120]
        break

    if not out['description']:
        for line in lines[:50]:
            if (len(line) > 12
                    and re.match(r'^[A-Z]', line)
                    and not any(s in line.upper() for s in CAPTION_SKIP)
                    and not re.search(r'\d{5}', line)
                    and not re.match(r'^Case\s', line, re.IGNORECASE)):
                out['description'] = line[:120]
                break

    return out


# ═════════════════════════════════════════════════════════════════════════════
# LINK DOCKET FOLDER DIALOG
# ═════════════════════════════════════════════════════════════════════════════
def link_folder_dialog(parent, db):
    """
    Scan a folder of docket-entry PDFs, match each one to existing docket
    entries via the ECF court-header stamp, then link or create entries.
    """
    folder = filedialog.askdirectory(
        title='Select Folder Containing Filed Document PDFs', parent=parent)
    if not folder:
        return
    folder = Path(folder)

    pdfs = sorted(folder.glob('*.pdf'), key=lambda p: p.name.lower())
    if not pdfs:
        messagebox.showinfo('No PDFs Found',
                            f'No PDF files were found in:\n{folder}', parent=parent)
        return

    # ── Scan every PDF ────────────────────────────────────────────────────────
    # Show a brief progress indicator (simple label in a small window)
    prog = tk.Toplevel(parent)
    prog.title('Scanning…')
    prog.configure(bg=C['bg'])
    prog.geometry(f'340x80+{parent.winfo_rootx()+200}+{parent.winfo_rooty()+200}')
    prog.grab_set()
    prog_lbl = tk.Label(prog, text=f'Scanning 0 / {len(pdfs)} …',
                        bg=C['bg'], fg=C['acc'], font=('Helvetica', 10))
    prog_lbl.pack(expand=True)
    prog.update()

    scan_results = []
    for i, pdf in enumerate(pdfs):
        prog_lbl.config(text=f'Scanning {i+1} / {len(pdfs)} …  {pdf.name[:30]}')
        prog.update()
        info = scan_pdf_header(pdf)
        scan_results.append({'path': pdf, **info})
    prog.destroy()

    # ── Build lookup maps from existing DB entries ────────────────────────────
    all_rows  = db.docket_all()
    by_entry  = {r['entry_no']: dict(r) for r in all_rows if r['entry_no']}
    by_id     = {r['id']: dict(r) for r in all_rows}
    by_date   = {}
    for r in all_rows:
        if r['date_filed']:
            by_date.setdefault(r['date_filed'], []).append(dict(r))

    # ── Determine initial action for each scan result ─────────────────────────
    _parser = PACERParser()
    for s in scan_results:
        en  = s['entry_no']       # parent entry number  (e.g. '22')
        an  = s.get('attachment_no', '')  # attachment number (e.g. '3'), '' if main entry
        df  = s['date_filed']
        ecf = f"{en}-{an}" if an else en   # full ECF reference for display

        if an:
            # ── Attachment: look up parent entry, build sub-entry ─────────────
            # 'par' not 'parent' — avoids shadowing the tkinter widget parameter
            par = by_entry.get(en)
            if par:
                # Use `or ''` — dict.get returns None for SQL NULL values
                att_labels = _parser.parse_attachment_list(par.get('description') or '')
                att_label  = att_labels.get(an, '')
                att_type   = (_parser.classify_attachment(att_label) if att_label
                              else _parser._classify(s['description']))
                # Use description from PDF heading if available and better
                if s['description'] and not att_label:
                    att_label = s['description']
                elif s['description'] and len(s['description']) > len(att_label):
                    att_label = s['description']

                s['action']         = 'new'   # always create a sub-entry
                s['parent_id']      = par['id']
                s['att_label']      = att_label or s['description'] or s['path'].stem
                s['att_type']       = att_type
                s['match_id']       = None
                s['match_label']    = (f"Sub-entry of ECF #{en} — "
                                       f"{(par.get('description') or '')[:40]}")
            else:
                # Parent not in DB yet — can still create entry, but no parent linkage
                s['action']         = 'new'
                s['parent_id']      = None
                s['att_label']      = s['description'] or s['path'].stem
                s['att_type']       = _parser._classify(s['description'])
                s['match_id']       = None
                s['match_label']    = f"Parent ECF #{en} not in docket — will create standalone"
                s['confidence']     = 'low'
        elif en and en in by_entry:
            # ── Main entry: link file to existing row ─────────────────────────
            row = by_entry[en]
            s['action']         = 'link'
            s['parent_id']      = None
            s['att_label']      = ''
            s['att_type']       = ''
            s['match_id']       = row['id']
            s['match_label']    = f"ECF #{en} — {(row['description'] or '')[:50]}"
        elif df and df in by_date:
            candidates = [r for r in by_date[df] if not r.get('file_path')]
            if candidates:
                c = candidates[0]
                s['action']      = 'link'
                s['parent_id']   = None
                s['att_label']   = ''
                s['att_type']    = ''
                s['match_id']    = c['id']
                s['match_label'] = f"Date match #{c['entry_no']} — {(c['description'] or '')[:45]}"
                s['confidence']  = 'low'
            else:
                s['action']      = 'new'
                s['parent_id']   = None
                s['att_label']   = s['description'] or s['path'].stem
                s['att_type']    = _parser._classify(s['description'])
                s['match_id']    = None
                s['match_label'] = '— Create new entry —'
        else:
            s['action']      = 'new' if s['confidence'] != 'none' else 'skip'
            s['parent_id']   = None
            s['att_label']   = s['description'] or s['path'].stem
            s['att_type']    = _parser._classify(s['description'])
            s['match_id']    = None
            s['match_label'] = '— Create new entry —' if s['action'] == 'new' else '— Skip —'

    # ── Dialog ────────────────────────────────────────────────────────────────
    top = tk.Toplevel(parent)
    top.title(f'Link Docket Folder — {len(pdfs)} PDFs')
    top.configure(bg=C['bg'])
    top.geometry(f'1100x580+{parent.winfo_rootx()+20}+{parent.winfo_rooty()+20}')
    top.grab_set()

    # Instructions
    tk.Label(top, text=f'Folder: {folder}',
             bg=C['bg'], fg=C['acc'], font=('Helvetica', 9, 'bold'),
             anchor='w').pack(fill='x', padx=12, pady=(10, 0))
    tk.Label(top,
             text='  Green = ECF header matched   Yellow = low-confidence / date-only   '
                  'Grey = no header (skip)     Double-click any row to override.',
             bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8), anchor='w'
             ).pack(fill='x', padx=12)

    # Treeview
    tvf  = tk.Frame(top, bg=C['bg2'])
    tvf.pack(fill='both', expand=True, padx=12, pady=(4, 0))
    cols = ('✓', 'Filename', 'ECF #', 'Date', 'Doc Type Detected', 'Action', 'Matched Docket Entry')
    tv   = ttk.Treeview(tvf, columns=cols, show='headings', style=_tv_style('lnk'))
    for col, w in [('✓', 28), ('Filename', 190), ('ECF #', 52), ('Date', 88),
                   ('Doc Type Detected', 165), ('Action', 62),
                   ('Matched Docket Entry', 310)]:
        tv.heading(col, text=col)
        tv.column(col, width=w, minwidth=24)
    tv.tag_configure('high', foreground=C['green'])
    tv.tag_configure('low',  foreground=C['yellow'])
    tv.tag_configure('none', foreground=C['fg3'])
    tv.tag_configure('skip', foreground=C['fg3'])
    vsb = ttk.Scrollbar(tvf, orient='vertical', command=tv.yview)
    tv.configure(yscrollcommand=vsb.set)
    vsb.pack(side='right', fill='y')
    tv.pack(fill='both', expand=True)

    iid_map = {}   # treeview iid → scan-result dict

    def _row_tag(s):
        if s['action'] == 'skip': return 'skip'
        return s['confidence']   # 'high' | 'low' | 'none'

    def _action_lbl(s):
        return {'link': 'Link', 'new': 'New', 'skip': 'Skip'}.get(s['action'], '?')

    def populate():
        tv.delete(*tv.get_children())
        iid_map.clear()
        for s in scan_results:
            desc   = s['description'] or '—'
            dtype  = _parser._classify(desc) if desc != '—' else '—'
            iid    = tv.insert('', 'end', values=(
                '✓' if s['action'] != 'skip' else '',
                s['path'].name[:32],
                s['entry_no'] or '—',
                disp(s['date_filed']),
                dtype,
                _action_lbl(s),
                s['match_label'],
            ), tags=(_row_tag(s),))
            iid_map[iid] = s
        update_stats()

    # ── Override dialog (double-click) ────────────────────────────────────────
    all_docket = list(db.docket_all())
    entry_labels = [
        f"#{r['entry_no']} {disp(r['date_filed'])} — {(r['description'] or '')[:45]}"
        for r in all_docket
    ]
    entry_ids = [r['id'] for r in all_docket]

    def on_dbl(event):
        col = tv.identify_column(event.x)
        if col == '#1': return   # leave toggle column for single-click
        sel = tv.selection()
        if not sel: return
        s = iid_map[sel[0]]

        ov = tk.Toplevel(top)
        ov.title('Override — ' + s['path'].name)
        ov.configure(bg=C['bg'])
        ov.geometry(f'520x270+{top.winfo_rootx()+100}+{top.winfo_rooty()+80}')
        ov.grab_set()

        tk.Label(ov, text=f"File: {s['path'].name}",
                 bg=C['bg'], fg=C['acc'], font=('Helvetica', 9, 'bold'),
                 anchor='w').pack(fill='x', padx=16, pady=(12, 2))
        tk.Label(ov, text=f"Detected header: {s['raw_header'] or '(none)'}",
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8),
                 anchor='w').pack(fill='x', padx=16, pady=(0, 8))

        act_var = tk.StringVar(value=s['action'])
        af = tk.Frame(ov, bg=C['bg']); af.pack(fill='x', padx=16)
        for v, lbl in [('link', 'Link to an existing docket entry'),
                       ('new',  'Create a new docket entry for this file'),
                       ('skip', 'Skip this file')]:
            tk.Radiobutton(af, text=lbl, variable=act_var, value=v,
                           bg=C['bg'], fg=C['fg'], selectcolor=C['bg3'],
                           activebackground=C['bg'],
                           font=('Helvetica', 9)).pack(anchor='w', pady=1)

        tk.Label(ov, text='Link to entry:', bg=C['bg'], fg=C['fg2'],
                 font=('Helvetica', 9)).pack(anchor='w', padx=16, pady=(8, 2))
        cur_lbl = next((entry_labels[i] for i, eid in enumerate(entry_ids)
                        if eid == s.get('match_id')), '')
        match_cb = _CB(ov, entry_labels, cur_lbl)
        match_cb.pack(fill='x', padx=16)

        def apply_ov():
            act = act_var.get()
            s['action'] = act
            if act == 'link':
                ml = match_cb.get()
                if ml in entry_labels:
                    idx = entry_labels.index(ml)
                    s['match_id']    = entry_ids[idx]
                    s['match_label'] = ml
                else:
                    s['action'] = 'skip'
            elif act == 'new':
                s['match_id']    = None
                s['match_label'] = '— Create new entry —'
            else:
                s['match_label'] = '— Skip —'
            populate()
            ov.destroy()

        bf2 = tk.Frame(ov, bg=C['bg']); bf2.pack(fill='x', padx=16, side='bottom', pady=10)
        _Btn(bf2, 'Cancel', ov.destroy).pack(side='right', padx=(4, 0))
        _Btn(bf2, 'Apply', apply_ov, bg=C['green'], fg=C['bg']).pack(side='right')
        ov.wait_window()

    tv.bind('<Double-1>', on_dbl)

    # ── Single-click col-1 toggles include/exclude ────────────────────────────
    def on_click(event):
        if tv.identify_column(event.x) != '#1': return
        row = tv.identify_row(event.y)
        if not row or row not in iid_map: return
        s = iid_map[row]
        if s['action'] == 'skip':
            en = s['entry_no']
            an = s.get('attachment_no', '')
            if an:
                s['action']      = 'new'
                par = by_entry.get(en)
                s['match_label'] = (f"Sub-entry of ECF #{en} — "
                                    f"{(par.get('description') or '')[:40]}"
                                    if par else f"Parent ECF #{en} not found")
            elif en and en in by_entry:
                s['action']      = 'link'
                s['match_id']    = by_entry[en]['id']
                s['match_label'] = (f"ECF #{en} — "
                                    f"{(by_entry[en].get('description') or '')[:50]}")
            else:
                s['action']      = 'new' if s['confidence'] != 'none' else 'skip'
                s['match_label'] = ('— Create new entry —'
                                    if s['action'] == 'new' else '— Skip —')
        else:
            s['action']      = 'skip'
            s['match_label'] = '— Skip —'
        populate()

    tv.bind('<Button-1>', on_click)

    # ── Stats bar ─────────────────────────────────────────────────────────────
    sf = tk.Frame(top, bg=C['bg']); sf.pack(fill='x', padx=12, pady=(4, 0))
    stats_lbl = tk.Label(sf, text='', bg=C['bg'], fg=C['fg2'], font=('Helvetica', 8))
    stats_lbl.pack(side='left')
    hint_lbl  = tk.Label(sf, text='Click ✓ column to include/exclude  ·  Double-click row to override',
                         bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8))
    hint_lbl.pack(side='right')

    def update_stats():
        n_link = sum(1 for s in scan_results if s['action'] == 'link')
        n_new  = sum(1 for s in scan_results if s['action'] == 'new')
        n_skip = sum(1 for s in scan_results if s['action'] == 'skip')
        stats_lbl.config(
            text=f'Will link: {n_link}   ·   New entries: {n_new}   ·   Skip: {n_skip}')

    # ── Apply ─────────────────────────────────────────────────────────────────
    def do_apply():
        linked = created = skipped = 0
        # Two passes: parents first so their IDs exist when children are inserted
        # Pass 1: link existing entries to files
        for s in scan_results:
            if s['action'] == 'link' and s.get('match_id'):
                db.docket_set_path(s['match_id'], s['path'])
                linked += 1
        # Pass 2: create new entries — parents before attachments
        parents_first = sorted(
            [s for s in scan_results if s['action'] == 'new'],
            key=lambda s: (bool(s.get('attachment_no')), s['entry_no']))
        new_entry_id_map = {}   # entry_no → newly-created db id
        for s in parents_first:
            if s['action'] != 'new':
                continue
            an        = s.get('attachment_no', '')
            parent_id = s.get('parent_id')     # id from existing DB entries
            # If parent was just created in this batch, use that id
            if an and not parent_id and s['entry_no'] in new_entry_id_map:
                parent_id = new_entry_id_map[s['entry_no']]
            desc  = s.get('att_label') or s['description'] or s['path'].stem
            dtype = s.get('att_type')  or _parser._classify(s['description'])
            ecf_no = f"{s['entry_no']}-{an}" if an else s['entry_no']
            new_id = db.docket_add(
                entry_no     = ecf_no,
                date_filed   = s['date_filed'],
                description  = desc,
                doc_type     = dtype,
                filing_party = '',
                file_path    = str(s['path']),
                notes        = (f"Attachment #{an} to ECF #{s['entry_no']}"
                                if an else f"Created from folder scan · {s['path'].name}"),
                parent_id    = parent_id,
                attachment_no= an,
            )
            if not an:
                new_entry_id_map[s['entry_no']] = new_id
            created += 1
        for s in scan_results:
            if s['action'] == 'skip':
                skipped += 1

        messagebox.showinfo('Folder Link Complete',
            f'Done:\n'
            f'  {linked} {"entry" if linked==1 else "entries"} linked to files\n'
            f'  {created} new {"entry" if created==1 else "entries"} created\n'
            f'  {skipped} {"file" if skipped==1 else "files"} skipped',
            parent=top)
        top.destroy()

    bf = tk.Frame(top, bg=C['bg'], pady=6)
    bf.pack(fill='x', padx=12, side='bottom')
    _Btn(bf, 'Cancel', top.destroy).pack(side='right', padx=(4, 0))
    _Btn(bf, 'Apply Links', do_apply, bg=C['green'], fg=C['bg']).pack(side='right')

    populate()
    top.wait_window()


# ═════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION  —  multi-case
# ═════════════════════════════════════════════════════════════════════════════
class CaseCommand(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Case Command')
        self.configure(bg=C['bg'])
        self.geometry('1420x880')
        self.minsize(1080, 660)
        self.resizable(True, True)

        self.master_db  = CaseMasterDB()
        self.active_db  = None    # CaseDB for currently open case
        self.active_id  = None    # master case id
        self._img_ref   = None
        self._prev_doc  = None

        self._build_ui()
        self._refresh_master()

    # ── helpers ───────────────────────────────────────────────────────────────
    @property
    def db(self):
        """Alias so all existing methods still work."""
        return self.active_db

    def _require_case(self):
        """Return True if a case is open, otherwise flash message."""
        if self.active_db:
            return True
        messagebox.showinfo('No Case Open',
            'Please open or add a case from the Case Master tab first.')
        self.nb.select(0)
        return False

    def _cases_dir(self):
        d = Path(self.master_db.get_setting('cases_dir', str(DEFAULT_CASES_DIR)))
        d.mkdir(parents=True, exist_ok=True)
        return d

    # ── title bar ─────────────────────────────────────────────────────────────
    def _set_title(self, m=None):
        if self.active_db is None:
            self.title('Case Command  |  No case open')
            if hasattr(self, 'case_lbl'):
                self.case_lbl.config(text='No case open — select or add a case in the Case Master tab',
                                     fg=C['fg3'])
            return
        if m is None: m = self.active_db.meta()
        cn = m.get('case_name', '—')
        ct = m.get('court', '')
        no = m.get('case_number', '')
        self.title(f"Case Command  |  {cn}  |  {ct}  |  {no}")
        if hasattr(self, 'case_lbl'):
            self.case_lbl.config(text=f"Active Case:  {cn}  ·  {ct}  ·  {no}",
                                 fg=C['acc'])

    # ── UI shell ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self, bg=C['bg2'], pady=5)
        hdr.pack(fill='x')
        self.case_lbl = tk.Label(hdr, text='', bg=C['bg2'],
                                 fg=C['fg3'], font=('Helvetica', 10, 'bold'), anchor='w')
        self.case_lbl.pack(side='left', padx=14)

        sf = tk.Frame(hdr, bg=C['bg2'])
        sf.pack(side='right', padx=14)
        tk.Label(sf, text='Search:', bg=C['bg2'], fg=C['fg2'],
                 font=('Helvetica', 9)).pack(side='left', padx=(0, 4))
        self.search_var = tk.StringVar()
        se = tk.Entry(sf, textvariable=self.search_var,
                      bg=C['bg3'], fg=C['fg'], insertbackground=C['fg'],
                      relief='flat', font=('Helvetica', 10), width=28)
        se.pack(side='left')
        se.bind('<Return>', lambda _: self._do_search())
        _Btn(sf, 'Go', self._do_search, bg=C['acc'], fg=C['bg']).pack(side='left', padx=(4, 0))

        # Notebook
        style = ttk.Style(); style.theme_use('clam')
        style.configure('TNotebook', background=C['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', background=C['bg3'], foreground=C['fg2'],
                        padding=(12, 5), font=('Helvetica', 9, 'bold'))
        style.map('TNotebook.Tab', background=[('selected', C['bg'])],
                  foreground=[('selected', C['acc'])])

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=6, pady=(4, 6))

        self.tab_master = tk.Frame(self.nb, bg=C['bg'])
        self.tab_dash   = tk.Frame(self.nb, bg=C['bg'])
        self.tab_docket = tk.Frame(self.nb, bg=C['bg'])
        self.tab_dl     = tk.Frame(self.nb, bg=C['bg'])
        self.tab_disc   = tk.Frame(self.nb, bg=C['bg'])
        self.tab_depo   = tk.Frame(self.nb, bg=C['bg'])
        self.tab_party  = tk.Frame(self.nb, bg=C['bg'])

        self.nb.add(self.tab_master, text='  Case Master  ')
        self.nb.add(self.tab_dash,   text='  Dashboard  ')
        self.nb.add(self.tab_docket, text='  Docket  ')
        self.nb.add(self.tab_dl,     text='  Court Hearings & Deadlines  ')
        self.nb.add(self.tab_disc,   text='  Written Discovery  ')
        self.nb.add(self.tab_depo,   text='  Depositions  ')
        self.nb.add(self.tab_party,  text='  Parties & Counsel  ')

        self._build_master_tab()
        self._build_dashboard()
        self._build_docket_tab()
        self._build_dl_tab()
        self._build_disc_tab()
        self._build_depo_tab()
        self._build_party_tab()

        self._set_title()

    # ═══════════════════════════════════════════════════════════════════════
    # CASE MASTER TAB
    # ═══════════════════════════════════════════════════════════════════════
    def _build_master_tab(self):
        p = self.tab_master

        # Cases folder bar
        ff = tk.Frame(p, bg=C['bg2'], pady=5); ff.pack(fill='x')
        tk.Label(ff, text='Cases folder:', bg=C['bg2'], fg=C['fg3'],
                 font=('Helvetica', 8)).pack(side='left', padx=(12, 4))
        self.cases_dir_lbl = tk.Label(ff, text='', bg=C['bg2'], fg=C['fg2'],
                                      font=('Helvetica', 8), anchor='w')
        self.cases_dir_lbl.pack(side='left', padx=(0, 8))
        _Btn(ff, 'Change Folder', self._change_cases_dir,
             bg=C['bg4']).pack(side='left')

        # Toolbar
        tb = tk.Frame(p, bg=C['bg'], pady=6); tb.pack(fill='x', padx=8)
        _Btn(tb, '+ Add Case',    self._add_case,
             bg=C['acc'], fg=C['bg']).pack(side='left', padx=(0, 4))
        _Btn(tb, 'Open Case',     self._open_selected_case).pack(side='left', padx=2)
        _Btn(tb, 'Archive',       lambda: self._set_case_status('Archived')).pack(side='left', padx=2)
        _Btn(tb, 'Reactivate',    lambda: self._set_case_status('Active')).pack(side='left', padx=2)
        _Btn(tb, 'Remove',        self._remove_case,
             bg=C['bg3']).pack(side='left', padx=(12, 2))
        _Btn(tb, '💾 Backup Case',  self._backup_case,
             bg=C['green'], fg=C['bg']).pack(side='right', padx=(4, 0))
        _Btn(tb, '📂 Restore Backup', self._restore_case,
             bg=C['bg3']).pack(side='right', padx=2)

        # Filter
        flt = tk.Frame(p, bg=C['bg2'], pady=3); flt.pack(fill='x', padx=8)
        tk.Label(flt, text='Show:', bg=C['bg2'], fg=C['fg2'],
                 font=('Helvetica', 9)).pack(side='left', padx=(4, 6))
        self.master_filter = tk.StringVar(value='Active')
        for v in ['All', 'Active', 'Archived']:
            tk.Radiobutton(flt, text=v, variable=self.master_filter, value=v,
                           bg=C['bg2'], fg=C['fg'], selectcolor=C['bg3'],
                           activebackground=C['bg2'], font=('Helvetica', 9),
                           command=self._refresh_master).pack(side='left', padx=4)

        # Case list treeview
        tf = tk.Frame(p, bg=C['bg2']); tf.pack(fill='both', expand=True, padx=8, pady=(4, 0))
        cols = ('Case Name', 'Court', 'Case #', 'Judge', 'Status', 'Last Opened')
        self.master_tv = ttk.Treeview(tf, columns=cols, show='headings',
                                      style=_tv_style('mstr'))
        for col, w in [('Case Name', 260), ('Court', 180), ('Case #', 130),
                       ('Judge', 170), ('Status', 80), ('Last Opened', 120)]:
            self.master_tv.heading(col, text=col)
            self.master_tv.column(col, width=w, minwidth=60)
        self.master_tv.tag_configure('active',   foreground=C['fg'])
        self.master_tv.tag_configure('archived', foreground=C['fg3'])
        self.master_tv.tag_configure('open',     foreground=C['acc'])
        vsb = ttk.Scrollbar(tf, orient='vertical', command=self.master_tv.yview)
        self.master_tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.master_tv.pack(fill='both', expand=True)
        self.master_tv.bind('<Double-1>', lambda _: self._open_selected_case())
        self._master_iid_map = {}   # iid → case id
        _make_sortable(self.master_tv)

        # Status bar
        sb = tk.Frame(p, bg=C['bg2'], pady=4); sb.pack(fill='x', padx=8)
        self.master_status = tk.Label(sb, text='', bg=C['bg2'], fg=C['fg3'],
                                      font=('Helvetica', 8), anchor='w')
        self.master_status.pack(side='left', padx=4)
        tk.Label(sb, text='Double-click to open a case',
                 bg=C['bg2'], fg=C['fg3'], font=('Helvetica', 8)).pack(side='right', padx=4)

    def _refresh_master(self):
        d = self._cases_dir()
        self.cases_dir_lbl.config(text=str(d))
        self.master_tv.delete(*self.master_tv.get_children())
        self._master_iid_map.clear()
        flt = self.master_filter.get()
        rows = [r for r in self.master_db.all_cases()
                if flt == 'All' or r['status'] == flt]
        for r in rows:
            tag = 'open' if r['id'] == self.active_id else \
                  ('archived' if r['status'] == 'Archived' else 'active')
            lo  = r['last_opened'] or ''
            try: lo = datetime.fromisoformat(lo).strftime('%m/%d/%Y %H:%M')
            except: pass
            iid = self.master_tv.insert('', 'end', values=(
                r['case_name'], r['court'], r['case_number'],
                r['judge'], r['status'], lo), tags=(tag,))
            self._master_iid_map[iid] = r['id']
        _apply_sort(self.master_tv)
        self.master_status.config(
            text=f'{len(rows)} case{"s" if len(rows)!=1 else ""} shown')

    def _selected_master_id(self):
        sel = self.master_tv.selection()
        if not sel:
            messagebox.showinfo('Select', 'Select a case first.')
            return None
        return self._master_iid_map.get(sel[0])

    def _change_cases_dir(self):
        d = filedialog.askdirectory(title='Select Cases Folder')
        if d:
            self.master_db.set_setting('cases_dir', d)
            self._refresh_master()

    # ── Add case ──────────────────────────────────────────────────────────────
    def _add_case(self):
        cases_dir = self._cases_dir()
        ts        = datetime.now().strftime('%Y%m%d_%H%M%S')
        db_path   = cases_dir / f'case_{ts}.db'
        new_db    = CaseDB(db_path)

        ok = first_run_dialog(self, new_db)
        if not ok:
            new_db.close()
            try: db_path.unlink()
            except: pass
            return

        meta = new_db.meta()
        cid  = self.master_db.add_case(
            case_name   = meta.get('case_name', ''),
            case_number = meta.get('case_number', ''),
            court       = meta.get('court', ''),
            judge       = meta.get('judge', ''),
            db_path     = db_path)
        new_db.close()
        self._refresh_master()

        # Automatically open the new case
        self._open_case(cid)

    # ── Open case ─────────────────────────────────────────────────────────────
    def _open_selected_case(self):
        cid = self._selected_master_id()
        if cid: self._open_case(cid)

    def _open_case(self, case_id):
        row = self.master_db.get_case(case_id)
        if not row:
            messagebox.showerror('Error', 'Case record not found.'); return
        db_path = Path(row['db_path'])
        if not db_path.exists():
            messagebox.showerror('File Not Found',
                f'The database for this case was not found:\n{db_path}\n\n'
                f'The file may have been moved or deleted.')
            return
        # Close previous
        self._close_active_case()
        self.active_db = CaseDB(db_path)
        self.active_id = case_id
        self.master_db.touch(case_id)
        self._set_title()
        self._refresh_master()
        self._refresh_all()
        self.nb.select(1)   # jump to Dashboard

    def _close_active_case(self):
        if self.active_db:
            self._clear_viewer()
            self.active_db.close()
            self.active_db = None
            self.active_id = None

    # ── Archive / Remove ──────────────────────────────────────────────────────
    def _set_case_status(self, status):
        cid = self._selected_master_id()
        if cid is None: return
        self.master_db.set_status(cid, status)
        self._refresh_master()

    def _remove_case(self):
        cid = self._selected_master_id()
        if cid is None: return
        row = self.master_db.get_case(cid)
        if not messagebox.askyesno('Remove Case',
            f'Remove "{row["case_name"]}" from Case Master?\n\n'
            f'The database file will NOT be deleted — only the registry entry.'):
            return
        if cid == self.active_id:
            self._close_active_case()
            self._set_title()
        self.master_db.remove(cid)
        self._refresh_master()

    # ── Backup ────────────────────────────────────────────────────────────────
    def _backup_case(self):
        # Use selected case from master list, or active case
        cid = None
        sel = self.master_tv.selection()
        if sel:
            cid = self._master_iid_map.get(sel[0])
        if cid is None:
            cid = self.active_id
        if cid is None:
            messagebox.showinfo('Select', 'Select a case to back up.'); return

        row = self.master_db.get_case(cid)
        if not row: return
        db_path = Path(row['db_path'])
        if not db_path.exists():
            messagebox.showerror('File Not Found',
                f'Database not found:\n{db_path}'); return

        safe = re.sub(r'[^\w\-]', '_',
                      row['case_number'] or row['case_name'] or 'case')[:40]
        default_fn = f"{safe}_{date.today()}.zip"
        save_path  = filedialog.asksaveasfilename(
            title='Save Case Backup',
            defaultextension='.zip',
            filetypes=[('ZIP archive', '*.zip')],
            initialfile=default_fn)
        if not save_path: return

        # Open a read connection (use active db if it's the right case)
        if cid == self.active_id and self.active_db:
            cdb, close_after = self.active_db, False
        else:
            cdb, close_after = CaseDB(db_path), True

        try:
            # Flush WAL to main db file before zipping
            cdb.cx.execute("PRAGMA wal_checkpoint(FULL)")
            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                # 1. Raw database file
                zf.write(str(db_path), 'case_data.db')

                # 2. Case info JSON
                meta = cdb.meta()
                zf.writestr('case_info.json', json.dumps(meta, indent=2))

                # 3. CSV exports of every table
                def to_csv(headers, rows):
                    buf = io.StringIO()
                    w   = csv.writer(buf)
                    w.writerow(headers)
                    w.writerows(rows)
                    return buf.getvalue()

                zf.writestr('docket.csv', to_csv(
                    ['Entry#','Date','Type','Filing Party','Description','File Path','Notes'],
                    [(r['entry_no'], disp(r['date_filed']), r['doc_type'],
                      r['filing_party'], r['description'], r['file_path'], r['notes'])
                     for r in cdb.docket_all()]))

                zf.writestr('deadlines.csv', to_csv(
                    ['Name','Type','Due Date','Status','Docket Ref','Notes'],
                    [(r['name'], r['dtype'], disp(r['due_date']),
                      r['status'], r['docket_ref'], r['notes'])
                     for r in cdb.dl_all()]))

                zf.writestr('discovery.csv', to_csv(
                    ['Type','Set#','Direction','Propounding','Responding',
                     '#Reqs','Served','Due','Responded','Status','Deficiency','Notes'],
                    [(r['rtype'], r['set_no'], r['direction'],
                      r['prop_party'], r['resp_party'], r['num_reqs'],
                      disp(r['date_served']), disp(r['due_date']), disp(r['resp_date']),
                      r['status'], r['deficiency'], r['notes'])
                     for r in cdb.disc_all()]))

                zf.writestr('depositions.csv', to_csv(
                    ['Witness','Role','Affiliation','Notice Date','Depo Date',
                     'Location','Format','Status','Errata Deadline',
                     'Transcript Bates','Topics','Issues'],
                    [(r['witness'], r['role'], r['affiliation'],
                      disp(r['notice_date']), disp(r['depo_date']),
                      r['location'], r['format'], r['status'],
                      disp(r['errata_deadline']), r['transcript_bates'],
                      r['topics'], r['issues'])
                     for r in cdb.depo_all()]))

                zf.writestr('parties.csv', to_csv(
                    ['Name','Role','Affiliation','Contact','Notes'],
                    [(r['name'], r['role'], r['affiliation'],
                      r['contact'], r['notes'])
                     for r in cdb.party_all()]))

                zf.writestr('counsel.csv', to_csv(
                    ['Name','Firm','Phone','Email','Party','Lead','Terminated','Notes'],
                    [(r['name'], r['firm'], r['phone'], r['email'],
                      r['party_name'] or '', 'Yes' if r['is_lead'] else '',
                      disp(r['term_date']), r['notes'])
                     for r in cdb.counsel_all()]))

            messagebox.showinfo('Backup Complete',
                f'Backed up to:\n{Path(save_path).name}\n\n'
                f'The ZIP contains the database and CSV exports of all tables.')
        except Exception as e:
            messagebox.showerror('Backup Failed', f'An error occurred:\n{e}')
        finally:
            if close_after: cdb.close()

    # ── Restore ───────────────────────────────────────────────────────────────
    def _restore_case(self):
        fp = filedialog.askopenfilename(
            title='Select Case Backup to Restore',
            filetypes=[('ZIP backup', '*.zip'),
                       ('SQLite database', '*.db'),
                       ('All files', '*.*')])
        if not fp: return
        fp = Path(fp)
        cases_dir = self._cases_dir()
        ts        = datetime.now().strftime('%Y%m%d_%H%M%S')

        if fp.suffix.lower() == '.zip':
            try:
                with zipfile.ZipFile(fp, 'r') as zf:
                    if 'case_data.db' not in zf.namelist():
                        messagebox.showerror('Invalid Backup',
                            'This ZIP does not appear to be a Case Command backup.\n'
                            '(Expected case_data.db inside)')
                        return
                    dest = cases_dir / f'restored_{ts}.db'
                    with zf.open('case_data.db') as src, \
                         open(dest, 'wb') as dst:
                        dst.write(src.read())
            except Exception as e:
                messagebox.showerror('Error', f'Could not read ZIP:\n{e}'); return
        elif fp.suffix.lower() == '.db':
            dest = cases_dir / f'restored_{ts}.db'
            shutil.copy2(fp, dest)
        else:
            messagebox.showerror('Unsupported',
                'Please select a .zip backup or .db database file.')
            return

        # Read meta from the restored db
        try:
            rdb  = CaseDB(dest)
            meta = rdb.meta()
            rdb.close()
        except Exception as e:
            messagebox.showerror('Error', f'Could not read restored database:\n{e}')
            return

        # Check for duplicate path
        if any(str(dest) == r['db_path'] for r in self.master_db.all_cases()):
            messagebox.showinfo('Already Registered',
                'This case is already in your Case Master.'); return

        cid = self.master_db.add_case(
            case_name   = meta.get('case_name', 'Restored Case'),
            case_number = meta.get('case_number', ''),
            court       = meta.get('court', ''),
            judge       = meta.get('judge', ''),
            db_path     = dest)
        self._refresh_master()
        messagebox.showinfo('Restore Complete',
            f'Case restored:\n{meta.get("case_name", "Unknown")}\n\n'
            f'Find it in the Case Master list.')

    # ═══════════════════════════════════════════════════════════════════════
    # DASHBOARD
    # ═══════════════════════════════════════════════════════════════════════
    def _build_dashboard(self):
        p = self.tab_dash
        self.card_frame = tk.Frame(p, bg=C['bg'])
        self.card_frame.pack(fill='x', padx=16, pady=(14, 8))

        tk.Label(p, text='UPCOMING DEADLINES (next 30 days + overdue)',
                 bg=C['bg'], fg=C['acc'], font=('Helvetica', 10, 'bold'),
                 anchor='w').pack(fill='x', padx=16, pady=(8, 4))

        uf = tk.Frame(p, bg=C['bg2']); uf.pack(fill='both', expand=True, padx=16, pady=(0, 12))
        cols = ('Name', 'Type', 'Due', 'Status')
        self.dash_tv = ttk.Treeview(uf, columns=cols, show='headings',
                                    style=_tv_style('dash'))
        for col, w in [('Name', 260), ('Type', 110), ('Due', 100), ('Status', 120)]:
            self.dash_tv.heading(col, text=col)
            self.dash_tv.column(col, width=w, minwidth=60)
        self.dash_tv.tag_configure('overdue', foreground=C['red'])
        self.dash_tv.tag_configure('soon',    foreground=C['yellow'])
        vsb = ttk.Scrollbar(uf, orient='vertical', command=self.dash_tv.yview)
        self.dash_tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y'); self.dash_tv.pack(fill='both', expand=True)
        _make_sortable(self.dash_tv)

    def _refresh_dashboard(self):
        for w in self.card_frame.winfo_children(): w.destroy()
        if self.active_db is None:
            tk.Label(self.card_frame,
                     text='Open a case from the Case Master tab',
                     bg=C['bg'], fg=C['fg3'],
                     font=('Helvetica', 10)).pack(pady=20)
            self.dash_tv.delete(*self.dash_tv.get_children())
            return
        s = self.active_db.dashboard()
        cards = [
            ('Docket Entries',  str(s['dock_total']),    C['acc']),
            ('Deadlines',       str(s['dl_total']),      C['fg']),
            ('Overdue',         str(s['dl_overdue']),    C['red']    if s['dl_overdue']    else C['fg3']),
            ('Due ≤ 14 days',   str(s['dl_soon']),       C['yellow'] if s['dl_soon']       else C['fg3']),
            ('Discovery Sets',  str(s['disc_total']),    C['fg']),
            ('Disc. Deficient', str(s['disc_deficient']),C['orange'] if s['disc_deficient'] else C['fg3']),
            ('Depositions',     str(s['depo_total']),    C['fg']),
            ('Depos Complete',  str(s['depo_done']),     C['green']),
        ]
        for i, (lbl, val, col) in enumerate(cards):
            cf = tk.Frame(self.card_frame, bg=C['bg3'], padx=16, pady=10)
            cf.grid(row=0, column=i, padx=6, sticky='nsew')
            self.card_frame.columnconfigure(i, weight=1)
            tk.Label(cf, text=val, bg=C['bg3'], fg=col,
                     font=('Helvetica', 22, 'bold')).pack()
            tk.Label(cf, text=lbl, bg=C['bg3'], fg=C['fg2'],
                     font=('Helvetica', 8)).pack()

        self.dash_tv.delete(*self.dash_tv.get_children())
        for r in s['upcoming']:
            tag = deadline_tag(r['due_date']) if r['status'] not in ('Completed', 'Vacated') else 'ok'
            self.dash_tv.insert('', 'end',
                values=(r['name'], r['dtype'], disp(r['due_date']), r['status']),
                tags=(tag,))
        _apply_sort(self.dash_tv)

    # ═══════════════════════════════════════════════════════════════════════
    # DOCKET
    # ═══════════════════════════════════════════════════════════════════════
    def _build_docket_tab(self):
        p = self.tab_docket

        tb = tk.Frame(p, bg=C['bg'], pady=6); tb.pack(fill='x', padx=8)
        _Btn(tb, '+ Add Entry',       lambda: self._dock_add()).pack(side='left', padx=(0, 4))
        _Btn(tb, 'Edit',              lambda: self._dock_edit()).pack(side='left', padx=2)
        _Btn(tb, 'Delete',            lambda: self._dock_del()).pack(side='left', padx=2)
        _Btn(tb, 'Import PACER PDF',  lambda: self._import_pacer(),
             bg=C['mauve'], fg=C['bg']).pack(side='left', padx=(12, 2))
        _Btn(tb, 'Link Docket Folder', lambda: self._link_folder(),
             bg=C['orange'], fg=C['bg']).pack(side='left', padx=2)
        _Btn(tb, 'Export CSV',        lambda: self._export_docket()).pack(side='right')

        ff = tk.Frame(p, bg=C['bg2'], pady=4); ff.pack(fill='x', padx=8)
        tk.Label(ff, text='Filter:', bg=C['bg2'], fg=C['fg2'],
                 font=('Helvetica', 9)).pack(side='left', padx=(8, 4))
        self.dock_filter = tk.StringVar()
        fe = tk.Entry(ff, textvariable=self.dock_filter,
                      bg=C['bg3'], fg=C['fg'], insertbackground=C['fg'],
                      relief='flat', font=('Helvetica', 9), width=30)
        fe.pack(side='left')
        fe.bind('<KeyRelease>', lambda _: self._refresh_docket())

        pw = tk.PanedWindow(p, orient='horizontal', bg=C['bg'], sashwidth=5)
        pw.pack(fill='both', expand=True, padx=8, pady=(4, 8))
        left  = tk.Frame(pw, bg=C['bg2'])
        right = tk.Frame(pw, bg=C['bg2'])
        pw.add(left,  minsize=400, stretch='always')
        pw.add(right, minsize=260, stretch='never')
        pw.paneconfig(left,  width=820)
        pw.paneconfig(right, width=420)

        cols = ('#', 'Date', 'Type', 'Filing Party', 'Description')
        self.dock_tv = ttk.Treeview(left, columns=cols, show='tree headings',
                                    style=_tv_style('dock'))
        self.dock_tv.column('#0', width=16, minwidth=16, stretch=False)
        self.dock_tv.heading('#0', text='')
        for col, w in [('#', 62), ('Date', 90), ('Type', 130),
                       ('Filing Party', 140), ('Description', 310)]:
            self.dock_tv.heading(col, text=col)
            self.dock_tv.column(col, width=w, minwidth=30)
        self.dock_tv.tag_configure('attachment', foreground=C['fg2'])
        vsb = ttk.Scrollbar(left, orient='vertical', command=self.dock_tv.yview)
        self.dock_tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y'); self.dock_tv.pack(fill='both', expand=True)
        self.dock_tv.bind('<<TreeviewSelect>>', self._dock_selected)
        _make_sortable(self.dock_tv)

        # ── Right pane: PDF viewer / entry text ───────────────────────────────
        vlbl = tk.Frame(right, bg=C['bg2'], pady=4); vlbl.pack(fill='x')
        self.dock_view_lbl = tk.Label(vlbl, text='Select an entry to view',
                                      bg=C['bg2'], fg=C['fg3'],
                                      font=('Helvetica', 9), anchor='w')
        self.dock_view_lbl.pack(side='left', padx=8)
        sm = dict(bg=C['bg3'], fg=C['fg'], relief='flat',
                  font=('Helvetica', 8), padx=5, pady=2, cursor='hand2')
        # Page navigation (jumps to that page in the continuous scroll)
        self.dock_pg_next = tk.Button(vlbl, text='pg >',
                                      command=lambda: self._view_pg(1), **sm)
        self.dock_pg_next.pack(side='right', padx=2)
        self.dock_pg_lbl = tk.Label(vlbl, text='', bg=C['bg2'], fg=C['fg3'],
                                    font=('Helvetica', 8))
        self.dock_pg_lbl.pack(side='right')
        self.dock_pg_prev = tk.Button(vlbl, text='< pg',
                                      command=lambda: self._view_pg(-1), **sm)
        self.dock_pg_prev.pack(side='right', padx=2)

        # Zoom controls
        self.dock_zoom_in  = tk.Button(vlbl, text='+',
                                       command=lambda: self._view_zoom(1.25), **sm)
        self.dock_zoom_out = tk.Button(vlbl, text='−',
                                       command=lambda: self._view_zoom(0.8), **sm)
        self.dock_zoom_fit = tk.Button(vlbl, text='Fit',
                                       command=self._view_zoom_fit, **sm)
        self.dock_zoom_lbl = tk.Label(vlbl, text='', bg=C['bg2'], fg=C['fg3'],
                                      font=('Helvetica', 8), width=5)
        self.dock_zoom_in.pack(side='right', padx=(2, 6))
        self.dock_zoom_lbl.pack(side='right')
        self.dock_zoom_out.pack(side='right', padx=2)
        self.dock_zoom_fit.pack(side='right', padx=(6, 2))

        cf = tk.Frame(right, bg=C['bg2']); cf.pack(fill='both', expand=True)
        self.dock_canvas = tk.Canvas(cf, bg=C['bg2'], highlightthickness=0)
        vsbc = ttk.Scrollbar(cf, orient='vertical',   command=self.dock_canvas.yview)
        hsbc = ttk.Scrollbar(cf, orient='horizontal', command=self.dock_canvas.xview)
        self.dock_canvas.configure(yscrollcommand=vsbc.set, xscrollcommand=hsbc.set)
        vsbc.pack(side='right', fill='y')
        hsbc.pack(side='bottom', fill='x')
        self.dock_canvas.pack(fill='both', expand=True)
        # Mouse-wheel: scroll vertically; Ctrl+Wheel: zoom
        self.dock_canvas.bind('<MouseWheel>',       self._canvas_scroll)
        self.dock_canvas.bind('<Button-4>',         self._canvas_scroll)
        self.dock_canvas.bind('<Button-5>',         self._canvas_scroll)
        self.dock_canvas.bind('<Control-MouseWheel>', self._canvas_zoom_wheel)
        self.dock_canvas.bind('<Control-Button-4>',   self._canvas_zoom_wheel)
        self.dock_canvas.bind('<Control-Button-5>',   self._canvas_zoom_wheel)
        # Re-render (re-fit) on pane resize if in Fit mode
        self.dock_canvas.bind('<Configure>', self._canvas_resized)
        # Track current page based on scroll position
        self.dock_canvas.bind('<B1-Motion>', lambda e: self._update_current_page())
        vsbc.bind('<ButtonRelease-1>', lambda e: self._update_current_page())

        # Viewer state
        self._view_page   = 0       # currently "visible" page (1-based display)
        self._view_total  = 0
        self._view_path   = None
        self._zoom_factor   = 1.0     # current zoom factor (1.0 = fit to width)
        self._view_fit    = True    # True = auto-fit to canvas width
        self._page_y_tops = []      # y-coord top of each page in canvas (for nav)
        self._img_refs    = []      # keep PhotoImage refs alive
        self.dock_pg_prev.config(state='disabled')
        self.dock_pg_next.config(state='disabled')
        self.dock_zoom_in.config(state='disabled')
        self.dock_zoom_out.config(state='disabled')
        self.dock_zoom_fit.config(state='disabled')
        self._dock_iid_map = {}

    def _refresh_docket(self):
        self.dock_tv.delete(*self.dock_tv.get_children())
        self._dock_iid_map = {}
        if self.active_db is None: return

        flt      = self.dock_filter.get().lower()
        all_rows = self.active_db.docket_all()

        def _matches(r):
            if not flt: return True
            return (flt in (r['description'] or '').lower()
                    or flt in (r['doc_type'] or '').lower()
                    or flt in (r['entry_no'] or '').lower()
                    or flt in (r['filing_party'] or '').lower())

        def _vals(r, ecf_override=None):
            ecf = ecf_override if ecf_override is not None else r['entry_no']
            return (ecf, disp(r['date_filed']),
                    r['doc_type'], r['filing_party'],
                    (r['description'] or '')[:80])

        if flt:
            matched_parent_ids = {r['id'] for r in all_rows
                                  if not r['parent_id'] and _matches(r)}
            for r in all_rows:
                show = _matches(r) or (r['parent_id'] and
                                       r['parent_id'] in matched_parent_ids)
                if show:
                    iid = self.dock_tv.insert(
                        '', 'end', values=_vals(r),
                        tags=('attachment',) if r['parent_id'] else ())
                    self._dock_iid_map[iid] = r['id']
        else:
            parents  = [r for r in all_rows if not r['parent_id']]
            children = {}
            for r in all_rows:
                if r['parent_id']:
                    children.setdefault(r['parent_id'], []).append(r)

            for r in parents:
                iid = self.dock_tv.insert('', 'end', values=_vals(r))
                self._dock_iid_map[iid] = r['id']
                for child in children.get(r['id'], []):
                    an   = child['attachment_no'] or ''
                    ecf  = (f"{r['entry_no']}-{an}"
                            if an else child['entry_no'])
                    ciid = self.dock_tv.insert(
                        iid, 'end',
                        values=_vals(child, ecf_override=ecf),
                        tags=('attachment',))
                    self._dock_iid_map[ciid] = child['id']
                if r['id'] in children:
                    self.dock_tv.item(iid, open=True)
        _apply_sort(self.dock_tv)

    def _dock_selected(self, _=None):
        sel = self.dock_tv.selection()
        if not sel or self.active_db is None: return
        rid  = self._dock_iid_map.get(sel[0])
        rows = [r for r in self.active_db.docket_all() if r['id'] == rid]
        if not rows: return
        row  = rows[0]
        fp   = row['file_path']
        self.dock_view_lbl.config(
            text=f"#{row['entry_no']}  {disp(row['date_filed'])}  {row['doc_type']}",
            fg=C['acc'])
        if fp and Path(fp).exists():
            self._load_viewer(fp)
        else:
            self._show_entry_text(row)

    def _show_entry_text(self, row):
        """Display full docket entry description as readable text in the right panel."""
        if self._prev_doc:
            try: self._prev_doc.close()
            except: pass
            self._prev_doc = None
        self._img_refs    = []
        self._page_y_tops = []
        self._view_total  = 0
        self._view_page   = 0
        self._view_path   = None
        self.dock_pg_lbl.config(text='')
        self.dock_pg_prev.config(state='disabled')
        self.dock_pg_next.config(state='disabled')
        self.dock_zoom_in.config(state='disabled')
        self.dock_zoom_out.config(state='disabled')
        self.dock_zoom_fit.config(state='disabled')
        self.dock_zoom_lbl.config(text='')

        self.dock_canvas.delete('all')
        self.dock_canvas.update_idletasks()
        cw = max(self.dock_canvas.winfo_width(), 360)   # canvas pixel width
        pad = 14
        wrap = cw - pad * 2

        # Header line
        hdr = (f"ECF #{row['entry_no'] or '—'}   "
               f"{disp(row['date_filed'])}   "
               f"{row['doc_type'] or ''}")
        if row['filing_party']:
            hdr += f"\nFiled by: {row['filing_party']}"
        hdr_id = self.dock_canvas.create_text(
            pad, pad, text=hdr, anchor='nw',
            fill=C['acc'], font=('Helvetica', 9, 'bold'), width=wrap)
        hdr_bbox = self.dock_canvas.bbox(hdr_id)
        y = (hdr_bbox[3] if hdr_bbox else pad + 24) + 6

        # Separator
        self.dock_canvas.create_line(pad, y, cw - pad, y, fill=C['bg4'])
        y += 10

        # Full description text
        desc = (row['description'] or '(No description)').replace('\u2212', '-').replace('\u2013', '-')
        txt_id = self.dock_canvas.create_text(
            pad, y, text=desc, anchor='nw',
            fill=C['fg'], font=('Helvetica', 9), width=wrap)
        txt_bbox = self.dock_canvas.bbox(txt_id)
        y = (txt_bbox[3] if txt_bbox else y + 40) + 10

        # Notes (if any)
        if row['notes']:
            self.dock_canvas.create_line(pad, y, cw - pad, y, fill=C['bg4'])
            y += 8
            self.dock_canvas.create_text(
                pad, y, text='Notes:', anchor='nw',
                fill=C['fg3'], font=('Helvetica', 8, 'bold'), width=wrap)
            y += 16
            notes_id = self.dock_canvas.create_text(
                pad, y, text=row['notes'], anchor='nw',
                fill=C['fg2'], font=('Helvetica', 8), width=wrap)
            nb = self.dock_canvas.bbox(notes_id)
            y = (nb[3] if nb else y + 20) + 10

        # No file indicator
        if not row['file_path']:
            self.dock_canvas.create_text(
                pad, y + 4, text='No PDF attached — use Link Docket Folder to attach one.',
                anchor='nw', fill=C['fg3'], font=('Helvetica', 8, 'italic'), width=wrap)
            y += 22
        elif row['file_path']:
            self.dock_canvas.create_text(
                pad, y + 4, text=f'File not found:\n{row["file_path"]}',
                anchor='nw', fill=C['red'], font=('Helvetica', 8), width=wrap)
            y += 36

        self.dock_canvas.config(scrollregion=(0, 0, cw, y + 20))
        self.dock_canvas.yview_moveto(0)

    def _load_viewer(self, path):
        if self._prev_doc:
            try: self._prev_doc.close()
            except: pass
        try:
            self._prev_doc   = fitz.open(str(path))
            self._view_total = len(self._prev_doc)
            self._view_page  = 0
            self._view_path  = path
            # Reset to fit-to-width whenever a new doc is loaded
            self._view_fit   = True
            self._zoom_factor  = 1.0
            self._render_all_pages()
            self.dock_canvas.yview_moveto(0)
        except Exception as e:
            self._clear_viewer(f'Cannot open file:\n{e}')

    def _render_all_pages(self):
        """Render every page of the current document stacked vertically on the
        canvas so the user can scroll continuously from start to end. Re-draws
        the entire canvas — call this whenever the zoom factor changes or the
        canvas is resized while in Fit mode."""
        if not self._prev_doc:
            return
        self.dock_canvas.delete('all')
        self._img_refs    = []
        self._page_y_tops = []

        # Canvas width available for fitting
        self.dock_canvas.update_idletasks()
        cw = max(self.dock_canvas.winfo_width() - 8, 200)

        # Pick zoom factor
        if self._view_fit:
            # Fit the widest page to the canvas width
            max_w = max(p.rect.width for p in self._prev_doc)
            base  = cw / max_w if max_w else 1.0
            zoom  = base
        else:
            # Explicit zoom: use target preview height as the 1.0 baseline
            # so that "100%" feels similar to the old single-page view.
            avg_h = sum(p.rect.height for p in self._prev_doc) / self._view_total
            base  = (PDF_H / avg_h) if avg_h else 1.0
            zoom  = base * self._zoom_factor

        # Stack each page vertically with a small gap
        gap   = 12
        y     = 4
        max_x = 0
        for i, page in enumerate(self._prev_doc):
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
                ph  = ImageTk.PhotoImage(img)
            except Exception:
                # If a single page fails, skip it but keep going
                self._page_y_tops.append(y)
                continue
            self._img_refs.append(ph)
            self._page_y_tops.append(y)
            self.dock_canvas.create_image(4, y, anchor='nw', image=ph)
            # Page-number label in the bottom-right corner of each page
            self.dock_canvas.create_text(
                4 + pix.width - 4, y + pix.height - 4,
                text=f'{i+1}', anchor='se',
                fill=C['fg3'], font=('Helvetica', 8))
            y     += pix.height + gap
            max_x = max(max_x, pix.width + 8)

        self.dock_canvas.config(scrollregion=(0, 0, max_x, y))

        # Buttons / labels
        self.dock_pg_prev.config(state='normal' if self._view_total > 1 else 'disabled')
        self.dock_pg_next.config(state='normal' if self._view_total > 1 else 'disabled')
        self.dock_zoom_in.config(state='normal')
        self.dock_zoom_out.config(state='normal')
        self.dock_zoom_fit.config(state='normal')
        self._update_zoom_label()
        self._update_page_label()

    def _update_zoom_label(self):
        if self._view_fit:
            self.dock_zoom_lbl.config(text='Fit')
        else:
            self.dock_zoom_lbl.config(text=f'{int(self._zoom_factor * 100)}%')

    def _update_page_label(self):
        if self._view_total:
            self.dock_pg_lbl.config(
                text=f' {self._view_page + 1}/{self._view_total} ')
        else:
            self.dock_pg_lbl.config(text='')

    def _update_current_page(self):
        """Figure out which page is currently in view based on scroll position
        and update the page label."""
        if not self._page_y_tops:
            return
        # Top of the visible region in canvas coordinates
        y0, _ = self.dock_canvas.yview()
        sr    = self.dock_canvas.cget('scrollregion').split()
        if len(sr) < 4:
            return
        try:
            total_h = float(sr[3])
        except ValueError:
            return
        top_y = y0 * total_h
        # Find the last page whose top is <= top_y (the currently "leading" page)
        page = 0
        for i, py in enumerate(self._page_y_tops):
            if py <= top_y + 10:
                page = i
            else:
                break
        if page != self._view_page:
            self._view_page = page
            self._update_page_label()

    def _view_pg(self, delta):
        """Jump the viewport to the previous/next page in the stack."""
        if not self._page_y_tops:
            return
        np = self._view_page + delta
        if 0 <= np < self._view_total:
            self._view_page = np
            sr = self.dock_canvas.cget('scrollregion').split()
            if len(sr) < 4:
                return
            try:
                total_h = float(sr[3])
            except ValueError:
                return
            y_top = self._page_y_tops[np]
            frac  = y_top / total_h if total_h else 0
            self.dock_canvas.yview_moveto(max(0, frac))
            self._update_page_label()

    def _view_zoom(self, factor):
        """Multiply the current zoom factor by `factor` (e.g. 1.25 to zoom in,
        0.8 to zoom out). Clamped to a reasonable range."""
        if not self._prev_doc:
            return
        # Remember the fraction of the document currently at the top so we can
        # restore that scroll position after re-rendering at the new zoom.
        y0, _ = self.dock_canvas.yview()
        self._view_fit  = False
        self._zoom_factor = max(0.25, min(self._zoom_factor * factor, 6.0))
        self._render_all_pages()
        self.dock_canvas.yview_moveto(y0)
        self._update_current_page()

    def _view_zoom_fit(self):
        if not self._prev_doc:
            return
        y0, _ = self.dock_canvas.yview()
        self._view_fit  = True
        self._zoom_factor = 1.0
        self._render_all_pages()
        self.dock_canvas.yview_moveto(y0)
        self._update_current_page()

    def _canvas_resized(self, _=None):
        """Re-render in fit mode when the pane/canvas is resized so that
        pages stay sized to the new width."""
        if self._prev_doc and self._view_fit:
            # Debounce slightly: only re-render if width actually changed.
            w = self.dock_canvas.winfo_width()
            if getattr(self, '_last_cw', None) != w:
                self._last_cw = w
                y0, _ = self.dock_canvas.yview()
                self._render_all_pages()
                self.dock_canvas.yview_moveto(y0)

    def _clear_viewer(self, msg=''):
        if self._prev_doc:
            try: self._prev_doc.close()
            except: pass
            self._prev_doc = None
        self._img_refs    = []
        self._page_y_tops = []
        self._view_total  = 0
        self._view_page   = 0
        self.dock_canvas.delete('all')
        if msg:
            self.dock_canvas.create_text(180, 120, text=msg, fill=C['fg3'],
                                         font=('Helvetica', 10), justify='center')
        self.dock_pg_lbl.config(text='')
        self.dock_pg_prev.config(state='disabled')
        self.dock_pg_next.config(state='disabled')
        self.dock_zoom_in.config(state='disabled')
        self.dock_zoom_out.config(state='disabled')
        self.dock_zoom_fit.config(state='disabled')
        self.dock_zoom_lbl.config(text='')

    def _canvas_scroll(self, e):
        if e.num == 4:   self.dock_canvas.yview_scroll(-1, 'units')
        elif e.num == 5: self.dock_canvas.yview_scroll( 1, 'units')
        else:            self.dock_canvas.yview_scroll(int(-1*(e.delta/120)), 'units')
        # After wheel-scrolling, figure out which page we're on now
        self._update_current_page()

    def _canvas_zoom_wheel(self, e):
        """Ctrl+MouseWheel zooms in/out."""
        # X11 uses Button-4/5 with num; Windows/Mac use delta
        if getattr(e, 'num', 0) == 4 or (getattr(e, 'delta', 0) > 0):
            self._view_zoom(1.1)
        elif getattr(e, 'num', 0) == 5 or (getattr(e, 'delta', 0) < 0):
            self._view_zoom(0.9)
        return 'break'   # prevent default scroll

    def _dock_add(self):
        if not self._require_case(): return
        data = docket_dialog(self, db=self.active_db)
        if data: self.active_db.docket_add(**data); self._refresh_all()

    def _dock_edit(self):
        if not self._require_case(): return
        sel = self.dock_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select an entry first.'); return
        rid = self._dock_iid_map.get(sel[0])
        row = next((r for r in self.active_db.docket_all() if r['id'] == rid), None)
        if row:
            data = docket_dialog(self, row, db=self.active_db)
            if data: self.active_db.docket_upd(rid, **data); self._refresh_all()

    def _dock_del(self):
        if not self._require_case(): return
        sel = self.dock_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select an entry first.'); return
        if messagebox.askyesno('Delete', 'Delete this docket entry?'):
            self.active_db.docket_del(self._dock_iid_map[sel[0]]); self._refresh_all()

    def _import_pacer(self):
        if not self._require_case(): return
        fp = filedialog.askopenfilename(title='Select PACER Docket PDF',
                                        filetypes=[('PDF', '*.pdf'), ('All', '*.*')])
        if fp:
            pacer_import_dialog(self, self.active_db, fp)
            # Sync updated meta back to master
            self.master_db.sync_meta(self.active_id, self.active_db.meta())
            self._set_title()
            self._refresh_all()
            self._refresh_master()

    def _link_folder(self):
        if not self._require_case(): return
        link_folder_dialog(self, self.active_db)
        self._refresh_all()

    def _export_docket(self):
        if not self._require_case(): return
        fp = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')],
            initialfile=f'docket_{date.today()}.csv')
        if fp:
            with open(fp, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Entry#', 'Date', 'Type', 'Filing Party', 'Description', 'Notes'])
                for r in self.active_db.docket_all():
                    w.writerow([r['entry_no'], disp(r['date_filed']), r['doc_type'],
                                r['filing_party'], r['description'], r['notes']])
            messagebox.showinfo('Exported', f'Docket exported to {Path(fp).name}')

    # ═══════════════════════════════════════════════════════════════════════
    # DEADLINES
    # ═══════════════════════════════════════════════════════════════════════
    def _build_dl_tab(self):
        p = self.tab_dl
        tb = tk.Frame(p, bg=C['bg'], pady=6); tb.pack(fill='x', padx=8)
        _Btn(tb, '+ Add',    lambda: self._dl_add()).pack(side='left', padx=(0, 4))
        _Btn(tb, 'Edit',     lambda: self._dl_edit()).pack(side='left', padx=2)
        _Btn(tb, 'Delete',   lambda: self._dl_del()).pack(side='left', padx=2)
        _Btn(tb, 'Export CSV', lambda: self._export_dl()).pack(side='right')
        tk.Label(p, text='  Overdue = red   ·   Due ≤ 14 days = yellow',
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8), anchor='w'
                 ).pack(fill='x', padx=8)
        f = tk.Frame(p, bg=C['bg2']); f.pack(fill='both', expand=True, padx=8, pady=6)
        cols = ('Name', 'Type', 'Due Date', 'Status', 'Docket Ref', 'Notes')
        self.dl_tv = ttk.Treeview(f, columns=cols, show='headings', style=_tv_style('dl'))
        for col, w in [('Name', 240), ('Type', 100), ('Due Date', 90),
                       ('Status', 110), ('Docket Ref', 80), ('Notes', 200)]:
            self.dl_tv.heading(col, text=col); self.dl_tv.column(col, width=w, minwidth=40)
        self.dl_tv.tag_configure('overdue', foreground=C['red'])
        self.dl_tv.tag_configure('soon',    foreground=C['yellow'])
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.dl_tv.yview)
        self.dl_tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y'); self.dl_tv.pack(fill='both', expand=True)
        self._dl_iid_map = {}
        _make_sortable(self.dl_tv)

    def _refresh_dl(self):
        self.dl_tv.delete(*self.dl_tv.get_children())
        self._dl_iid_map = {}
        if self.active_db is None: return
        for r in self.active_db.dl_all():
            tag = deadline_tag(r['due_date']) if r['status'] not in ('Completed', 'Vacated') else 'ok'
            iid = self.dl_tv.insert('', 'end',
                values=(r['name'], r['dtype'], disp(r['due_date']),
                        r['status'], r['docket_ref'], r['notes']), tags=(tag,))
            self._dl_iid_map[iid] = r['id']
        _apply_sort(self.dl_tv)

    def _dl_add(self):
        if not self._require_case(): return
        data = deadline_dialog(self)
        if data: self.active_db.dl_add(**data); self._refresh_all()

    def _dl_edit(self):
        if not self._require_case(): return
        sel = self.dl_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a deadline first.'); return
        rid = self._dl_iid_map[sel[0]]
        row = next((r for r in self.active_db.dl_all() if r['id'] == rid), None)
        if row:
            data = deadline_dialog(self, row)
            if data: self.active_db.dl_upd(rid, **data); self._refresh_all()

    def _dl_del(self):
        if not self._require_case(): return
        sel = self.dl_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a deadline first.'); return
        if messagebox.askyesno('Delete', 'Delete this deadline?'):
            self.active_db.dl_del(self._dl_iid_map[sel[0]]); self._refresh_all()

    def _export_dl(self):
        if not self._require_case(): return
        fp = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')],
            initialfile=f'deadlines_{date.today()}.csv')
        if fp:
            with open(fp, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Name', 'Type', 'Due Date', 'Status', 'Docket Ref', 'Notes'])
                for r in self.active_db.dl_all():
                    w.writerow([r['name'], r['dtype'], disp(r['due_date']),
                                r['status'], r['docket_ref'], r['notes']])
            messagebox.showinfo('Exported', f'Deadlines exported to {Path(fp).name}')

    # ═══════════════════════════════════════════════════════════════════════
    # DISCOVERY
    # ═══════════════════════════════════════════════════════════════════════
    def _build_disc_tab(self):
        p = self.tab_disc
        tb = tk.Frame(p, bg=C['bg'], pady=6); tb.pack(fill='x', padx=8)
        _Btn(tb, '+ Add Set',  lambda: self._disc_add()).pack(side='left', padx=(0, 4))
        _Btn(tb, 'Edit Set',   lambda: self._disc_edit()).pack(side='left', padx=2)
        _Btn(tb, 'Delete Set', lambda: self._disc_del()).pack(side='left', padx=2)
        _Btn(tb, 'Export CSV', lambda: self._export_disc()).pack(side='right')
        ff = tk.Frame(p, bg=C['bg2'], pady=3); ff.pack(fill='x', padx=8)
        tk.Label(ff, text='Show:', bg=C['bg2'], fg=C['fg2'],
                 font=('Helvetica', 9)).pack(side='left', padx=(8, 4))
        self.disc_filter = tk.StringVar(value='All')
        for v in ['All', 'RFP', 'ROG', 'RFA', 'Deficient', 'Pending']:
            tk.Radiobutton(ff, text=v, variable=self.disc_filter, value=v,
                           bg=C['bg2'], fg=C['fg'], selectcolor=C['bg3'],
                           activebackground=C['bg2'], font=('Helvetica', 9),
                           command=self._refresh_disc).pack(side='left', padx=4)
        f = tk.Frame(p, bg=C['bg2']); f.pack(fill='both', expand=True, padx=8, pady=(4, 8))
        cols = ('Type', 'Set #', 'Dir', 'Propounding', 'Responding',
                '# Reqs', 'Date Served', 'Due Date', 'Resp Date', 'Status', 'Deficiency')
        self.disc_tv = ttk.Treeview(f, columns=cols, show='headings', style=_tv_style('disc'))
        widths = [50, 50, 90, 140, 140, 60, 90, 90, 90, 110, 110]
        for col, w in zip(cols, widths):
            self.disc_tv.heading(col, text=col); self.disc_tv.column(col, width=w, minwidth=30)
        self.disc_tv.tag_configure('deficient', foreground=C['orange'])
        self.disc_tv.tag_configure('overdue',   foreground=C['red'])
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.disc_tv.yview)
        hsb = ttk.Scrollbar(f, orient='horizontal', command=self.disc_tv.xview)
        self.disc_tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x')
        self.disc_tv.pack(fill='both', expand=True)
        self._disc_iid_map = {}
        _make_sortable(self.disc_tv)

    def _refresh_disc(self):
        self.disc_tv.delete(*self.disc_tv.get_children())
        self._disc_iid_map = {}
        if self.active_db is None: return
        flt = self.disc_filter.get()
        for r in self.active_db.disc_all():
            if flt == 'RFP' and r['rtype'] != 'RFP': continue
            if flt == 'ROG' and r['rtype'] != 'ROG': continue
            if flt == 'RFA' and r['rtype'] != 'RFA': continue
            if flt == 'Deficient' and r['deficiency'] in ('None', 'Resolved'): continue
            if flt == 'Pending'   and r['status'] != 'Pending': continue
            tag = 'deficient' if r['deficiency'] not in ('None', 'Resolved') else \
                  deadline_tag(r['due_date']) if r['status'] == 'Pending' else 'ok'
            iid = self.disc_tv.insert('', 'end', values=(
                r['rtype'], r['set_no'], r['direction'],
                r['prop_party'], r['resp_party'], r['num_reqs'],
                disp(r['date_served']), disp(r['due_date']), disp(r['resp_date']),
                r['status'], r['deficiency']), tags=(tag,))
            self._disc_iid_map[iid] = r['id']
        _apply_sort(self.disc_tv)

    def _disc_add(self):
        if not self._require_case(): return
        m = self.active_db.meta(); days = int(m.get('default_response_days', '30'))
        data = disc_set_dialog(self, default_days=days)
        if data: self.active_db.disc_add(**data); self._refresh_all()

    def _disc_edit(self):
        if not self._require_case(): return
        sel = self.disc_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a set first.'); return
        rid = self._disc_iid_map[sel[0]]
        row = next((r for r in self.active_db.disc_all() if r['id'] == rid), None)
        if row:
            m = self.active_db.meta(); days = int(m.get('default_response_days', '30'))
            data = disc_set_dialog(self, row, days)
            if data: self.active_db.disc_upd(rid, **data); self._refresh_all()

    def _disc_del(self):
        if not self._require_case(): return
        sel = self.disc_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a set first.'); return
        if messagebox.askyesno('Delete', 'Delete this discovery set?'):
            self.active_db.disc_del(self._disc_iid_map[sel[0]]); self._refresh_all()

    def _export_disc(self):
        if not self._require_case(): return
        fp = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')],
            initialfile=f'discovery_{date.today()}.csv')
        if fp:
            with open(fp, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Type', 'Set#', 'Direction', 'Propounding', 'Responding',
                            '#Reqs', 'Served', 'Due', 'Responded',
                            'Status', 'Deficiency', 'Notes'])
                for r in self.active_db.disc_all():
                    w.writerow([r['rtype'], r['set_no'], r['direction'],
                                r['prop_party'], r['resp_party'], r['num_reqs'],
                                disp(r['date_served']), disp(r['due_date']),
                                disp(r['resp_date']), r['status'],
                                r['deficiency'], r['notes']])
            messagebox.showinfo('Exported', f'Discovery exported to {Path(fp).name}')

    # ═══════════════════════════════════════════════════════════════════════
    # DEPOSITIONS
    # ═══════════════════════════════════════════════════════════════════════
    def _build_depo_tab(self):
        p = self.tab_depo
        tb = tk.Frame(p, bg=C['bg'], pady=6); tb.pack(fill='x', padx=8)
        _Btn(tb, '+ Add',    lambda: self._depo_add()).pack(side='left', padx=(0, 4))
        _Btn(tb, 'Edit',     lambda: self._depo_edit()).pack(side='left', padx=2)
        _Btn(tb, 'Delete',   lambda: self._depo_del()).pack(side='left', padx=2)
        _Btn(tb, 'Export CSV', lambda: self._export_depo()).pack(side='right')
        f = tk.Frame(p, bg=C['bg2']); f.pack(fill='both', expand=True, padx=8, pady=6)
        cols = ('Witness', 'Role', 'Affiliation', 'Notice Date',
                'Depo Date', 'Format', 'Status', 'Errata Deadline', 'Issues')
        self.depo_tv = ttk.Treeview(f, columns=cols, show='headings',
                                    style=_tv_style('depo'))
        widths = [160, 90, 140, 90, 90, 90, 130, 110, 160]
        for col, w in zip(cols, widths):
            self.depo_tv.heading(col, text=col); self.depo_tv.column(col, width=w, minwidth=40)
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.depo_tv.yview)
        hsb = ttk.Scrollbar(f, orient='horizontal', command=self.depo_tv.xview)
        self.depo_tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x')
        self.depo_tv.pack(fill='both', expand=True)
        self._depo_iid_map = {}
        _make_sortable(self.depo_tv)

    def _refresh_depo(self):
        self.depo_tv.delete(*self.depo_tv.get_children())
        self._depo_iid_map = {}
        if self.active_db is None: return
        for r in self.active_db.depo_all():
            iid = self.depo_tv.insert('', 'end', values=(
                r['witness'], r['role'], r['affiliation'],
                disp(r['notice_date']), disp(r['depo_date']), r['format'],
                r['status'], disp(r['errata_deadline']),
                (r['issues'] or '')[:60]))
            self._depo_iid_map[iid] = r['id']
        _apply_sort(self.depo_tv)

    def _depo_add(self):
        if not self._require_case(): return
        data = depo_dialog(self)
        if data: self.active_db.depo_add(**data); self._refresh_all()

    def _depo_edit(self):
        if not self._require_case(): return
        sel = self.depo_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a deposition first.'); return
        rid = self._depo_iid_map[sel[0]]
        row = next((r for r in self.active_db.depo_all() if r['id'] == rid), None)
        if row:
            data = depo_dialog(self, row)
            if data: self.active_db.depo_upd(rid, **data); self._refresh_all()

    def _depo_del(self):
        if not self._require_case(): return
        sel = self.depo_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a deposition first.'); return
        if messagebox.askyesno('Delete', 'Delete this deposition?'):
            self.active_db.depo_del(self._depo_iid_map[sel[0]]); self._refresh_all()

    def _export_depo(self):
        if not self._require_case(): return
        fp = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')],
            initialfile=f'depositions_{date.today()}.csv')
        if fp:
            with open(fp, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Witness', 'Role', 'Affiliation', 'Notice Date', 'Depo Date',
                            'Location', 'Format', 'Status', 'Errata Deadline',
                            'Transcript Bates', 'Topics', 'Issues'])
                for r in self.active_db.depo_all():
                    w.writerow([r['witness'], r['role'], r['affiliation'],
                                disp(r['notice_date']), disp(r['depo_date']),
                                r['location'], r['format'], r['status'],
                                disp(r['errata_deadline']), r['transcript_bates'],
                                r['topics'], r['issues']])
            messagebox.showinfo('Exported', f'Depositions exported to {Path(fp).name}')

    # ═══════════════════════════════════════════════════════════════════════
    # PARTIES & COUNSEL
    # ═══════════════════════════════════════════════════════════════════════
    def _build_party_tab(self):
        p  = self.tab_party
        pw = tk.PanedWindow(p, orient='vertical', bg=C['bg'], sashwidth=5)
        pw.pack(fill='both', expand=True, padx=8, pady=8)

        ptop = tk.Frame(pw, bg=C['bg']); pw.add(ptop, minsize=180, stretch='always')
        tb1  = tk.Frame(ptop, bg=C['bg'], pady=4); tb1.pack(fill='x')
        tk.Label(tb1, text='PARTIES', bg=C['bg'], fg=C['acc'],
                 font=('Helvetica', 9, 'bold')).pack(side='left', padx=4)
        _Btn(tb1, '+ Add',  lambda: self._party_add()).pack(side='left', padx=(8, 2))
        _Btn(tb1, 'Edit',   lambda: self._party_edit()).pack(side='left', padx=2)
        _Btn(tb1, 'Delete', lambda: self._party_del()).pack(side='left', padx=2)
        pf = tk.Frame(ptop, bg=C['bg2']); pf.pack(fill='both', expand=True)
        cols1 = ('Name', 'Role', 'Affiliation', 'Contact', 'Notes')
        self.party_tv = ttk.Treeview(pf, columns=cols1, show='headings',
                                     style=_tv_style('pty'))
        for col, w in [('Name', 180), ('Role', 110), ('Affiliation', 180),
                       ('Contact', 160), ('Notes', 200)]:
            self.party_tv.heading(col, text=col); self.party_tv.column(col, width=w, minwidth=40)
        vsb1 = ttk.Scrollbar(pf, orient='vertical', command=self.party_tv.yview)
        self.party_tv.configure(yscrollcommand=vsb1.set)
        vsb1.pack(side='right', fill='y'); self.party_tv.pack(fill='both', expand=True)
        self._party_iid_map = {}
        _make_sortable(self.party_tv)

        cbot = tk.Frame(pw, bg=C['bg']); pw.add(cbot, minsize=180, stretch='always')
        tb2  = tk.Frame(cbot, bg=C['bg'], pady=4); tb2.pack(fill='x')
        tk.Label(tb2, text='COUNSEL', bg=C['bg'], fg=C['acc'],
                 font=('Helvetica', 9, 'bold')).pack(side='left', padx=4)
        _Btn(tb2, '+ Add',  lambda: self._counsel_add()).pack(side='left', padx=(8, 2))
        _Btn(tb2, 'Edit',   lambda: self._counsel_edit()).pack(side='left', padx=2)
        _Btn(tb2, 'Delete', lambda: self._counsel_del()).pack(side='left', padx=2)
        cf = tk.Frame(cbot, bg=C['bg2']); cf.pack(fill='both', expand=True)
        cols2 = ('Name', 'Firm', 'Phone', 'Email', 'Party', 'Lead', 'Terminated')
        self.counsel_tv = ttk.Treeview(cf, columns=cols2, show='headings',
                                       style=_tv_style('csl'))
        for col, w in [('Name', 160), ('Firm', 180), ('Phone', 100), ('Email', 180),
                       ('Party', 140), ('Lead', 50), ('Terminated', 90)]:
            self.counsel_tv.heading(col, text=col); self.counsel_tv.column(col, width=w, minwidth=40)
        self.counsel_tv.tag_configure('terminated', foreground=C['fg3'])
        vsb2 = ttk.Scrollbar(cf, orient='vertical', command=self.counsel_tv.yview)
        self.counsel_tv.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side='right', fill='y'); self.counsel_tv.pack(fill='both', expand=True)
        self._counsel_iid_map = {}
        _make_sortable(self.counsel_tv)

    def _refresh_parties(self):
        self.party_tv.delete(*self.party_tv.get_children())
        self._party_iid_map = {}
        self.counsel_tv.delete(*self.counsel_tv.get_children())
        self._counsel_iid_map = {}
        if self.active_db is None: return
        for r in self.active_db.party_all():
            iid = self.party_tv.insert('', 'end', values=(
                r['name'], r['role'], r['affiliation'], r['contact'], r['notes']))
            self._party_iid_map[iid] = r['id']
        for r in self.active_db.counsel_all():
            tag = 'terminated' if r['term_date'] else 'ok'
            iid = self.counsel_tv.insert('', 'end', values=(
                r['name'], r['firm'], r['phone'], r['email'],
                r['party_name'] or '', 'Yes' if r['is_lead'] else '',
                disp(r['term_date'])), tags=(tag,))
            self._counsel_iid_map[iid] = r['id']
        _apply_sort(self.party_tv)
        _apply_sort(self.counsel_tv)

    def _party_add(self):
        if not self._require_case(): return
        data = party_dialog(self)
        if data: self.active_db.party_add(**data); self._refresh_parties()

    def _party_edit(self):
        if not self._require_case(): return
        sel = self.party_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a party first.'); return
        rid = self._party_iid_map[sel[0]]
        row = next((r for r in self.active_db.party_all() if r['id'] == rid), None)
        if row:
            data = party_dialog(self, row)
            if data: self.active_db.party_upd(rid, **data); self._refresh_parties()

    def _party_del(self):
        if not self._require_case(): return
        sel = self.party_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a party first.'); return
        if messagebox.askyesno('Delete', 'Delete this party?'):
            self.active_db.party_del(self._party_iid_map[sel[0]]); self._refresh_parties()

    def _counsel_add(self):
        if not self._require_case(): return
        data = counsel_dialog(self, self.active_db)
        if data: self.active_db.counsel_add(**data); self._refresh_parties()

    def _counsel_edit(self):
        if not self._require_case(): return
        sel = self.counsel_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a counsel entry first.'); return
        rid = self._counsel_iid_map[sel[0]]
        row = next((r for r in self.active_db.counsel_all() if r['id'] == rid), None)
        if row:
            data = counsel_dialog(self, self.active_db, row)
            if data: self.active_db.counsel_upd(rid, **data); self._refresh_parties()

    def _counsel_del(self):
        if not self._require_case(): return
        sel = self.counsel_tv.selection()
        if not sel: messagebox.showinfo('Select', 'Select a counsel entry first.'); return
        if messagebox.askyesno('Delete', 'Delete this counsel?'):
            self.active_db.counsel_del(self._counsel_iid_map[sel[0]]); self._refresh_parties()

    # ═══════════════════════════════════════════════════════════════════════
    # SEARCH
    # ═══════════════════════════════════════════════════════════════════════
    def _do_search(self):
        if not self._require_case(): return
        q = self.search_var.get().strip()
        if not q: return
        results = self.active_db.search(q)

        top = tk.Toplevel(self)
        top.title(f'Search Results: "{q}"')
        top.configure(bg=C['bg'])
        top.geometry(f'760x420+{self.winfo_rootx()+80}+{self.winfo_rooty()+80}')
        tk.Label(top, text=f'{len(results)} result(s) for "{q}"',
                 bg=C['bg'], fg=C['acc'], font=('Helvetica', 10, 'bold'), anchor='w'
                 ).pack(fill='x', padx=12, pady=(10, 4))
        f   = tk.Frame(top, bg=C['bg2']); f.pack(fill='both', expand=True, padx=12, pady=(0, 8))
        tv  = ttk.Treeview(f, columns=('Tab', 'Match'), show='headings',
                            style=_tv_style('srch'))
        tv.heading('Tab',   text='Section'); tv.column('Tab',   width=100, minwidth=80)
        tv.heading('Match', text='Match');   tv.column('Match', width=600, minwidth=200)
        vsb = ttk.Scrollbar(f, orient='vertical', command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y'); tv.pack(fill='both', expand=True)

        TAB_MAP = {'Docket': 2, 'Deadlines': 3, 'Discovery': 4,
                   'Depositions': 5, 'Parties': 6}
        for tab, _, label in results:
            tv.insert('', 'end', values=(tab, label))

        def jump(event):
            sel = tv.selection()
            if not sel: return
            vals = tv.item(sel[0], 'values')
            idx  = TAB_MAP.get(vals[0], 0)
            self.nb.select(idx); top.destroy()
        tv.bind('<Double-1>', jump)
        tk.Label(top, text='Double-click a result to jump to that tab.',
                 bg=C['bg'], fg=C['fg3'], font=('Helvetica', 8)).pack(pady=(0, 6))

    # ═══════════════════════════════════════════════════════════════════════
    # REFRESH ALL
    # ═══════════════════════════════════════════════════════════════════════
    def _refresh_all(self):
        self._refresh_dashboard()
        self._refresh_docket()
        self._refresh_dl()
        self._refresh_disc()
        self._refresh_depo()
        self._refresh_parties()

    def destroy(self):
        self._close_active_case()
        self.master_db.close()
        super().destroy()


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app = CaseCommand()
    app.mainloop()

# ── entry point ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app = CaseCommand()
    app.mainloop()
